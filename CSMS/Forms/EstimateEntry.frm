VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO301B~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSEstimateEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estimate Data Entry"
   ClientHeight    =   7050
   ClientLeft      =   1635
   ClientTop       =   1905
   ClientWidth     =   10380
   ClipControls    =   0   'False
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
   Icon            =   "EstimateEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   10380
   Begin VB.Frame fraDiscount_ 
      Caption         =   "Enter Discount Percentage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1245
      Left            =   3443
      TabIndex        =   46
      Top             =   7635
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame fraAddAccessories_ 
      Appearance      =   0  'Flat
      Caption         =   "Add/Edit Accessories"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   825
      Left            =   2130
      TabIndex        =   64
      Top             =   7380
      Width           =   2205
   End
   Begin VB.Frame fraAddParts_ 
      Appearance      =   0  'Flat
      Caption         =   "Add/Edit Parts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   885
      Left            =   150
      TabIndex        =   41
      Top             =   7320
      Width           =   1785
   End
   Begin VB.Frame fraAddMaterials_ 
      Appearance      =   0  'Flat
      Caption         =   "Add/Edit Materials"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   5070
      TabIndex        =   42
      Top             =   7350
      Width           =   1995
      Begin VB.TextBox txtMatPOCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   1230
         MaxLength       =   2
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   5190
         Width           =   375
      End
      Begin VB.Label Label45 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PO code"
         Enabled         =   0   'False
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
         Left            =   420
         TabIndex        =   44
         Top             =   5220
         Width           =   1185
      End
   End
   Begin VB.Frame fraAddJobs_ 
      Appearance      =   0  'Flat
      Caption         =   "Add/Edit Jobs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   885
      Left            =   7260
      TabIndex        =   40
      Top             =   7320
      Width           =   2955
   End
   Begin VB.PictureBox pic3 
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
      Height          =   795
      Left            =   7470
      ScaleHeight     =   765
      ScaleWidth      =   2835
      TabIndex        =   47
      Top             =   6210
      Width           =   2865
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Customer Vehicle Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   52
         Left            =   120
         TabIndex        =   50
         Top             =   510
         Width           =   2280
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "F7   - Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   51
         Left            =   120
         TabIndex        =   49
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "F2   - Participation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   50
         Left            =   120
         TabIndex        =   48
         Top             =   30
         Width           =   1440
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
      Height          =   900
      Left            =   150
      ScaleHeight     =   900
      ScaleWidth      =   7380
      TabIndex        =   54
      Top             =   6210
      Width           =   7380
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   6510
         MouseIcon       =   "EstimateEntry.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   5790
         MouseIcon       =   "EstimateEntry.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   5070
         MouseIcon       =   "EstimateEntry.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   795
         Left            =   4350
         MouseIcon       =   "EstimateEntry.frx":1EA0
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":1FF2
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Move to Last Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   795
         Left            =   3630
         MouseIcon       =   "EstimateEntry.frx":2342
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":2494
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Move to First Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   2910
         MouseIcon       =   "EstimateEntry.frx":27F2
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":2944
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   2190
         MouseIcon       =   "EstimateEntry.frx":2C3E
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":2D90
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   1470
         MouseIcon       =   "EstimateEntry.frx":30E8
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":323A
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload"
         Height          =   795
         Left            =   750
         MouseIcon       =   "EstimateEntry.frx":3599
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":36EB
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Upload to RO"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   60
         MouseIcon       =   "EstimateEntry.frx":476D
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":48BF
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Visible         =   0   'False
         Width           =   705
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
      Height          =   825
      Left            =   5790
      ScaleHeight     =   825
      ScaleWidth      =   1605
      TabIndex        =   51
      Top             =   6210
      Width           =   1605
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   870
         MouseIcon       =   "EstimateEntry.frx":4BEA
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":4D3C
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   150
         MouseIcon       =   "EstimateEntry.frx":507A
         MousePointer    =   99  'Custom
         Picture         =   "EstimateEntry.frx":51CC
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   60
      TabIndex        =   21
      Top             =   0
      Width           =   10275
      Begin VB.TextBox cboModel 
         Height          =   315
         Left            =   5430
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   2610
         Width           =   2085
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5190
         TabIndex        =   73
         Top             =   1470
         Width           =   405
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5190
         TabIndex        =   72
         Top             =   900
         Width           =   405
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1470
         Width           =   3975
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   8070
         Top             =   1350
      End
      Begin VB.TextBox txtRONO 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox txtInvoiceNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtPart_amt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11250
         MaxLength       =   10
         TabIndex        =   16
         Top             =   2100
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtDte_recd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         MaxLength       =   10
         TabIndex        =   11
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtDte_comp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2010
         MaxLength       =   10
         TabIndex        =   12
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtDte_Rel 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   13
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CheckBox chkParticipat 
         Caption         =   "Participation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   14
         Top             =   1470
         Width           =   1845
      End
      Begin Crystal.CrystalReport rptEstimate 
         Left            =   9660
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Estimate Print Out"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.ComboBox cboRecd_by 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5700
         TabIndex        =   10
         Text            =   "cboRecd_by"
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   4515
      End
      Begin VB.TextBox txtEstimateNo 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtNiym 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   900
         Width           =   3975
      End
      Begin VB.TextBox txtPlate_No 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   3
         Top             =   2610
         Width           =   2265
      End
      Begin VB.TextBox txtAcct_No 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   4
         Top             =   900
         Width           =   1005
      End
      Begin VB.TextBox txtROType 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11250
         MaxLength       =   1
         TabIndex        =   5
         Top             =   1410
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtSvc_No 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11250
         MaxLength       =   1
         TabIndex        =   6
         Top             =   1770
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtTerm 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   7
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox txtKm_rdg 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8010
         MaxLength       =   9
         TabIndex        =   8
         Top             =   2040
         Width           =   1005
      End
      Begin VB.TextBox txtSektion 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9120
         MaxLength       =   3
         TabIndex        =   9
         Top             =   2040
         Width           =   1035
      End
      Begin VB.TextBox txtParticipat 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   15
         Top             =   1470
         Width           =   1005
      End
      Begin VB.TextBox txtCertific8 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2490
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   17
         Top             =   2610
         Width           =   2835
      End
      Begin VB.TextBox txtMake 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7620
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2610
         Width           =   2535
      End
      Begin VB.Label labDetId 
         BackColor       =   &H000000FF&
         Caption         =   "Label48"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   8760
         TabIndex        =   70
         Top             =   1560
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   58
         Left            =   180
         TabIndex        =   69
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Repair Order no"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   57
         Left            =   2130
         TabIndex        =   66
         Top             =   150
         Width           =   1305
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   17
         Left            =   3810
         TabIndex        =   45
         Top             =   150
         Width           =   885
      End
      Begin VB.Label labID 
         BackColor       =   &H000000FF&
         Caption         =   "Label18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8760
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estimate no"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   37
         Top             =   150
         Width           =   975
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   5640
         TabIndex        =   36
         Top             =   690
         Width           =   720
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   150
         TabIndex        =   35
         Top             =   2400
         Width           =   705
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROType"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   10470
         TabIndex        =   34
         Top             =   1500
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   10500
         TabIndex        =   33
         Top             =   1860
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Term"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   5670
         TabIndex        =   32
         Top             =   150
         Width           =   780
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM Reading"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   7
         Left            =   8040
         TabIndex        =   31
         Top             =   1830
         Width           =   960
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   8
         Left            =   9150
         TabIndex        =   30
         Top             =   1830
         Width           =   915
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Advisor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   5700
         TabIndex        =   29
         Top             =   1830
         Width           =   1140
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Recorded"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   10
         Left            =   150
         TabIndex        =   28
         Top             =   1830
         Width           =   1200
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Completed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   11
         Left            =   2040
         TabIndex        =   27
         Top             =   1830
         Width           =   1320
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Released"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   12
         Left            =   3840
         TabIndex        =   26
         Top             =   1830
         Width           =   1170
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   15
         Left            =   10470
         TabIndex        =   25
         Top             =   2190
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Warranty Certificate Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   14
         Left            =   2520
         TabIndex        =   24
         Top             =   2400
         Width           =   2340
      End
      Begin VB.Label labCAPv 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5460
         TabIndex        =   23
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   13
         Left            =   7710
         TabIndex        =   22
         Top             =   2400
         Width           =   1500
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   150
         TabIndex        =   39
         Top             =   690
         Width           =   1350
      End
      Begin VB.Label lblSTATUS 
         Alignment       =   2  'Center
         Caption         =   "STATUS"
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
         Height          =   285
         Left            =   7170
         TabIndex        =   63
         Top             =   390
         Width           =   2925
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   60
      TabIndex        =   19
      Top             =   3030
      Width           =   10275
      Begin TabDlg.SSTab SSTab1_ 
         Height          =   1245
         Left            =   10380
         TabIndex        =   20
         Top             =   300
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   2196
         _Version        =   393216
         Tabs            =   5
         Tab             =   3
         TabsPerRow      =   5
         TabHeight       =   706
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Details"
         TabPicture(0)   =   "EstimateEntry.frx":551C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "F3-Jobs"
         TabPicture(1)   =   "EstimateEntry.frx":5538
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "F4-Parts"
         TabPicture(2)   =   "EstimateEntry.frx":5554
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "F5-Materials"
         TabPicture(3)   =   "EstimateEntry.frx":5570
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).ControlCount=   0
         TabCaption(4)   =   "F6 - Accessories"
         TabPicture(4)   =   "EstimateEntry.frx":558C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
      End
      Begin XtremeSuiteControls.TabControl SSTab1 
         Height          =   2925
         Left            =   30
         TabIndex        =   93
         Top             =   150
         Width           =   10185
         _Version        =   655364
         _ExtentX        =   17965
         _ExtentY        =   5159
         _StockProps     =   64
         Appearance      =   2
         Color           =   4
         PaintManager.Layout=   2
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.FixedTabWidth=   125
         ItemCount       =   5
         Item(0).Caption =   "Details"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "fraDetails"
         Item(1).Caption =   "F3 - Jobs"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "fraJobs"
         Item(2).Caption =   "F4 - Parts"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "fraParts"
         Item(3).Caption =   "F5 - Materials"
         Item(3).ControlCount=   1
         Item(3).Control(0)=   "fraMaterials"
         Item(4).Caption =   "F6 - Accessories"
         Item(4).ControlCount=   1
         Item(4).Control(0)=   "lstAccessories"
         Begin VB.Frame fraMaterials 
            BackColor       =   &H00808080&
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
            Height          =   2505
            Left            =   -69940
            TabIndex        =   100
            Top             =   360
            Visible         =   0   'False
            Width           =   10065
            Begin MSComctlLib.ListView lstMaterials 
               Height          =   2505
               Left            =   0
               TabIndex        =   205
               Top             =   0
               Width           =   10065
               _ExtentX        =   17754
               _ExtentY        =   4419
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HoverSelection  =   -1  'True
               _Version        =   393217
               ForeColor       =   0
               BackColor       =   16777215
               Appearance      =   1
               MousePointer    =   99
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "EstimateEntry.frx":55A8
               NumItems        =   9
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "LINE #"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "MAT. CODE"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "DESCRIPTION"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Text            =   "QTY"
                  Object.Width           =   1147
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "UNITPRICE"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "AMOUNT"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   6
                  Text            =   "WSC"
                  Object.Width           =   1147
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   7
                  Text            =   "DISC."
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "ID"
                  Object.Width           =   2
               EndProperty
            End
         End
         Begin VB.Frame fraParts 
            BackColor       =   &H00808080&
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
            Height          =   2505
            Left            =   -69940
            TabIndex        =   98
            Top             =   360
            Visible         =   0   'False
            Width           =   10065
            Begin MSComctlLib.ListView lstParts 
               Height          =   2505
               Left            =   0
               TabIndex        =   204
               Top             =   0
               Width           =   10065
               _ExtentX        =   17754
               _ExtentY        =   4419
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HoverSelection  =   -1  'True
               _Version        =   393217
               ForeColor       =   0
               BackColor       =   16777215
               Appearance      =   1
               MousePointer    =   99
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "EstimateEntry.frx":570A
               NumItems        =   9
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "LINE #"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "PART NUMBER"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "DESCRIPTION"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Text            =   "QTY"
                  Object.Width           =   1147
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "UNITPRICE"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "AMOUNT"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   6
                  Text            =   "WSC"
                  Object.Width           =   1147
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   7
                  Text            =   "DISC."
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "ID"
                  Object.Width           =   2
               EndProperty
            End
         End
         Begin VB.Frame fraJobs 
            BackColor       =   &H00808080&
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
            Height          =   2505
            Left            =   -69940
            TabIndex        =   96
            Top             =   360
            Visible         =   0   'False
            Width           =   10065
            Begin MSComctlLib.ListView lstJobs 
               Height          =   2505
               Left            =   0
               TabIndex        =   203
               Top             =   0
               Width           =   10065
               _ExtentX        =   17754
               _ExtentY        =   4419
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HoverSelection  =   -1  'True
               _Version        =   393217
               ForeColor       =   0
               BackColor       =   16777215
               Appearance      =   1
               MousePointer    =   99
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "EstimateEntry.frx":586C
               NumItems        =   7
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "LINE #"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "JOB CODE"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "JOB DESCRIPTION"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "AMOUNT"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   4
                  Text            =   "WC"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "DISCOUNT"
                  Object.Width           =   2328
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "ID"
                  Object.Width           =   2
               EndProperty
            End
         End
         Begin VB.Frame fraDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
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
            ForeColor       =   &H80000008&
            Height          =   2505
            Left            =   60
            TabIndex        =   94
            Top             =   360
            Width           =   10065
            Begin MSFlexGridLib.MSFlexGrid grdDetails 
               Height          =   2505
               Left            =   0
               TabIndex        =   95
               Top             =   0
               Width           =   10065
               _ExtentX        =   17754
               _ExtentY        =   4419
               _Version        =   393216
               Rows            =   5
               Cols            =   9
               ForeColor       =   0
               ForeColorFixed  =   0
               BackColorSel    =   -2147483633
               ForeColorSel    =   16777215
               BackColorBkg    =   -2147483633
               TextStyleFixed  =   3
               SelectionMode   =   1
               Appearance      =   0
               MousePointer    =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "EstimateEntry.frx":59CE
            End
         End
         Begin MSComctlLib.ListView lstAccessories 
            Height          =   2505
            Left            =   -69940
            TabIndex        =   206
            Top             =   360
            Visible         =   0   'False
            Width           =   10065
            _ExtentX        =   17754
            _ExtentY        =   4419
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "EstimateEntry.frx":5CE8
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LINE #"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ACC. CODE"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "DESCRIPTION"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "QTY"
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "UNITPRICE"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "AMOUNT"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "WSC"
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "DISC."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
   End
   Begin VB.PictureBox fraAddAccessories 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   2903
      ScaleHeight     =   4545
      ScaleWidth      =   4545
      TabIndex        =   176
      Top             =   1238
      Width           =   4575
      Begin VB.CommandButton cmdAccDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   855
         Left            =   180
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":5E4A
         Style           =   1  'Graphical
         TabIndex        =   190
         Top             =   3630
         Width           =   765
      End
      Begin VB.CommandButton cmdAccCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   855
         Left            =   3660
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":6ECC
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   3630
         Width           =   765
      End
      Begin VB.ComboBox cboAccessories 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   186
         Text            =   "cboDescription"
         Top             =   1410
         Width           =   4335
      End
      Begin VB.ComboBox cboAccCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1230
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   185
         Top             =   780
         Width           =   2685
      End
      Begin VB.ComboBox cboAccChargeto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "EstimateEntry.frx":7F4E
         Left            =   1230
         List            =   "EstimateEntry.frx":7F50
         TabIndex        =   184
         Text            =   "cboChargeTo"
         Top             =   1770
         Width           =   585
      End
      Begin VB.TextBox txtAccQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1230
         MaxLength       =   5
         TabIndex        =   183
         Text            =   "0.0"
         Top             =   2130
         Width           =   555
      End
      Begin VB.TextBox txtAccUnitPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   182
         Text            =   "0.00"
         Top             =   2490
         Width           =   1545
      End
      Begin VB.TextBox txtAccAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   181
         Text            =   "0.00"
         Top             =   2850
         Width           =   1545
      End
      Begin VB.TextBox txtAccDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1230
         MaxLength       =   3
         TabIndex        =   180
         Text            =   "0"
         Top             =   3210
         Width           =   555
      End
      Begin VB.TextBox txtACCLineNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1230
         TabIndex        =   179
         Text            =   "Text1"
         Top             =   420
         Width           =   645
      End
      Begin VB.TextBox txtAccPOCODE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   178
         Top             =   1950
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdAccSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   855
         Left            =   2940
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":7F52
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   3630
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   1230
         MaxLength       =   2
         TabIndex        =   187
         Text            =   "Text1"
         Top             =   1410
         Width           =   345
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   42
         Left            =   450
         TabIndex        =   200
         Top             =   2940
         Width           =   660
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   45
         Left            =   240
         TabIndex        =   199
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   41
         Left            =   390
         TabIndex        =   198
         Top             =   3300
         Width           =   720
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   46
         Left            =   180
         TabIndex        =   197
         Top             =   1170
         Width           =   945
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Line No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   48
         Left            =   450
         TabIndex        =   196
         Top             =   450
         Width           =   660
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   47
         Left            =   450
         TabIndex        =   195
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label57 
         BackColor       =   &H000000C0&
         Caption         =   "Part Code"
         Enabled         =   0   'False
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
         Left            =   3150
         TabIndex        =   194
         Top             =   2580
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   44
         Left            =   420
         TabIndex        =   193
         Top             =   2190
         Width           =   675
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   43
         Left            =   330
         TabIndex        =   192
         Top             =   2580
         Width           =   780
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "( % )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   56
         Left            =   1890
         TabIndex        =   191
         Top             =   3270
         Width           =   345
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption5 
         Height          =   345
         Left            =   -60
         TabIndex        =   177
         Top             =   -30
         Width           =   5265
         _Version        =   655364
         _ExtentX        =   9287
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   " Add/Edit Accessories"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox fraAddParts 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4545
      Left            =   2918
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   152
      Top             =   1253
      Width           =   4545
      Begin VB.CommandButton cmdPartsDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   855
         Left            =   120
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":A024
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   3570
         Width           =   795
      End
      Begin VB.CommandButton cmdPartsCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   855
         Left            =   3630
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":B0A6
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   3570
         Width           =   795
      End
      Begin VB.TextBox txtPartsLineNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   162
         Text            =   "0"
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox txtPartDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   161
         Text            =   "0"
         Top             =   3150
         Width           =   555
      End
      Begin VB.TextBox txtPartAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   160
         Text            =   "0.00"
         Top             =   2790
         Width           =   1575
      End
      Begin VB.TextBox txtUnitPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   159
         Text            =   "0.00"
         Top             =   2430
         Width           =   1545
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   158
         Text            =   "0.0"
         Top             =   2070
         Width           =   555
      End
      Begin VB.TextBox txtPartCode 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   157
         Text            =   "Text1"
         Top             =   1800
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.ComboBox cboChargeTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "EstimateEntry.frx":C128
         Left            =   1200
         List            =   "EstimateEntry.frx":C12A
         TabIndex        =   156
         Text            =   "cboChargeTo"
         Top             =   1710
         Width           =   585
      End
      Begin VB.ComboBox cboPartNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   155
         Top             =   720
         Width           =   2685
      End
      Begin VB.ComboBox cboDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   154
         Text            =   "cboDescription"
         Top             =   1350
         Width           =   4335
      End
      Begin VB.CommandButton cmdPartsSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   855
         Left            =   2880
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":C12C
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   3570
         Width           =   765
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   38
         Left            =   300
         TabIndex        =   175
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   37
         Left            =   390
         TabIndex        =   174
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label Label32 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part Code"
         Enabled         =   0   'False
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
         Left            =   270
         TabIndex        =   173
         Top             =   1380
         Width           =   1305
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   34
         Left            =   420
         TabIndex        =   172
         Top             =   750
         Width           =   630
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Line No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   33
         Left            =   420
         TabIndex        =   171
         Top             =   390
         Width           =   660
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   35
         Left            =   150
         TabIndex        =   170
         Top             =   1110
         Width           =   945
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   40
         Left            =   360
         TabIndex        =   169
         Top             =   3210
         Width           =   720
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   36
         Left            =   210
         TabIndex        =   168
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   39
         Left            =   420
         TabIndex        =   167
         Top             =   2880
         Width           =   660
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "( % )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   55
         Left            =   1800
         TabIndex        =   166
         Top             =   3210
         Width           =   345
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   345
         Left            =   -60
         TabIndex        =   153
         Top             =   -30
         Width           =   5415
         _Version        =   655364
         _ExtentX        =   9551
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   " Add/Edit Parts"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox fraAddMaterials 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   3323
      ScaleHeight     =   4635
      ScaleWidth      =   3705
      TabIndex        =   130
      Top             =   1193
      Width           =   3735
      Begin VB.CommandButton cmdMatDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   795
         Left            =   120
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":E1FE
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   3630
         Width           =   825
      End
      Begin VB.TextBox txtMatLineNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   140
         Text            =   "Text1"
         Top             =   450
         Width           =   525
      End
      Begin VB.CommandButton cmdMatCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   2820
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":F280
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   3630
         Width           =   825
      End
      Begin VB.ComboBox cboMatChargeTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1200
         TabIndex        =   138
         Text            =   "cboMatChargeTo"
         Top             =   1770
         Width           =   585
      End
      Begin VB.TextBox txtMatDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   137
         Text            =   "0"
         Top             =   3180
         Width           =   585
      End
      Begin VB.TextBox txtMatAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   136
         Text            =   "0.00"
         Top             =   2820
         Width           =   1545
      End
      Begin VB.TextBox txtMatUnitPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   135
         Text            =   "0.00"
         Top             =   2460
         Width           =   1545
      End
      Begin VB.TextBox txtMatQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   134
         Text            =   "0.0"
         Top             =   2130
         Width           =   555
      End
      Begin VB.ComboBox cboMaterial 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   133
         Text            =   "cboMaterial"
         Top             =   1410
         Width           =   3525
      End
      Begin VB.ComboBox cboMatCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   132
         Text            =   "cboMatCode"
         Top             =   810
         Width           =   2415
      End
      Begin VB.CommandButton cmdMatSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   795
         Left            =   2040
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":10302
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   3630
         Width           =   795
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   24
         Left            =   450
         TabIndex        =   151
         Top             =   2880
         Width           =   660
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   21
         Left            =   210
         TabIndex        =   150
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   25
         Left            =   360
         TabIndex        =   149
         Top             =   3240
         Width           =   720
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   20
         Left            =   120
         TabIndex        =   148
         Top             =   1170
         Width           =   660
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Line No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   18
         Left            =   420
         TabIndex        =   147
         Top             =   510
         Width           =   660
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   22
         Left            =   420
         TabIndex        =   146
         Top             =   2190
         Width           =   675
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   23
         Left            =   330
         TabIndex        =   145
         Top             =   2550
         Width           =   780
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mat. Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   19
         Left            =   240
         TabIndex        =   144
         Top             =   870
         Width           =   825
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "( % )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   53
         Left            =   1860
         TabIndex        =   143
         Top             =   3240
         Width           =   345
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   315
         Left            =   -60
         TabIndex        =   131
         Top             =   0
         Width           =   4335
         _Version        =   655364
         _ExtentX        =   7646
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   " Add/Edit Materials"
         ForeColor       =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   4194304
      End
   End
   Begin VB.PictureBox fraAddJobs 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   2378
      ScaleHeight     =   5145
      ScaleWidth      =   5595
      TabIndex        =   104
      Top             =   938
      Width           =   5625
      Begin VB.CommandButton cmdJobCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   855
         Left            =   4710
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":123D4
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   4200
         Width           =   795
      End
      Begin VB.TextBox txtJobPostCode 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   3420
         TabIndex        =   117
         Text            =   "Text1"
         Top             =   1590
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtJobLineNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   116
         Text            =   "0"
         Top             =   390
         Width           =   585
      End
      Begin VB.CommandButton cmdJobDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   855
         Left            =   180
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":13456
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   4200
         Width           =   795
      End
      Begin VB.ComboBox cboJobChargeTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1200
         TabIndex        =   114
         Text            =   "cboJobChargeTo"
         Top             =   1470
         Width           =   585
      End
      Begin VB.TextBox txtJobDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   113
         Text            =   "0"
         Top             =   2580
         Width           =   465
      End
      Begin VB.TextBox txtJobRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   112
         Text            =   "0.00"
         Top             =   2220
         Width           =   1425
      End
      Begin VB.TextBox txtJobDetail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   915
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   111
         Text            =   "EstimateEntry.frx":144D8
         Top             =   3210
         Width           =   5325
      End
      Begin VB.OptionButton optByDescription 
         Caption         =   "By Job &Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3510
         TabIndex        =   110
         Top             =   390
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.OptionButton optByCode 
         Caption         =   "By &Job Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1950
         TabIndex        =   109
         Top             =   390
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox cboJcode 
         Height          =   330
         Left            =   1200
         TabIndex        =   108
         Top             =   750
         Width           =   1935
      End
      Begin VB.TextBox cboJobCode 
         Height          =   330
         Left            =   1200
         TabIndex        =   107
         Top             =   1110
         Width           =   4305
      End
      Begin VB.TextBox txtHRS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   106
         Text            =   "0.00"
         Top             =   1860
         Width           =   885
      End
      Begin VB.CommandButton cmdJobSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   855
         Left            =   3930
         MaskColor       =   &H0000FFFF&
         Picture         =   "EstimateEntry.frx":144DE
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   4200
         Width           =   795
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Line No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   31
         Left            =   420
         TabIndex        =   129
         Top             =   450
         Width           =   660
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   28
         Left            =   210
         TabIndex        =   128
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   27
         Left            =   360
         TabIndex        =   127
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   26
         Left            =   360
         TabIndex        =   126
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label labCAP 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Job Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   32
         Left            =   180
         TabIndex        =   125
         Top             =   2970
         Width           =   1785
      End
      Begin VB.Label labJobDet_Vol 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "det_vol"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   510
         TabIndex        =   124
         Top             =   4320
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   30
         Left            =   330
         TabIndex        =   123
         Top             =   840
         Width           =   780
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Desc."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   29
         Left            =   300
         TabIndex        =   122
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "( % )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   54
         Left            =   1770
         TabIndex        =   121
         Top             =   2640
         Width           =   345
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Std Hrs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   62
         Left            =   465
         TabIndex        =   120
         Top             =   1920
         Width           =   600
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   -30
         TabIndex        =   105
         Top             =   0
         Width           =   5685
         _Version        =   655364
         _ExtentX        =   10028
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   " Add/Edit Jobs"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox fraDiscount 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   3443
      ScaleHeight     =   1395
      ScaleWidth      =   3465
      TabIndex        =   201
      Top             =   2813
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtDiscAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   990
         MaxLength       =   8
         TabIndex        =   97
         Text            =   "Tex"
         Top             =   450
         Width           =   1275
      End
      Begin wizButton.cmd cmdCancelDisk 
         Height          =   375
         Left            =   1680
         TabIndex        =   99
         Top             =   960
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "EstimateEntry.frx":165B0
      End
      Begin wizButton.cmd cmdOkDisc 
         Height          =   375
         Left            =   690
         TabIndex        =   101
         Top             =   960
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         TX              =   "&Ok"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "EstimateEntry.frx":165CC
      End
      Begin VB.Label labCAP 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   49
         Left            =   2280
         TabIndex        =   102
         Top             =   510
         Width           =   345
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption6 
         Height          =   315
         Left            =   -90
         TabIndex        =   202
         Top             =   0
         Width           =   3585
         _Version        =   655364
         _ExtentX        =   6324
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   " Enter Discount Percentage"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox PicInsurance 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   3173
      ScaleHeight     =   4425
      ScaleWidth      =   4005
      TabIndex        =   74
      Top             =   1320
      Visible         =   0   'False
      Width           =   4035
      Begin VB.Frame fraParticipation 
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
         Height          =   2355
         Left            =   120
         TabIndex        =   78
         Top             =   1200
         Width           =   3765
         Begin VB.TextBox txtPartLabor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Left            =   1860
            MaxLength       =   13
            TabIndex        =   83
            Text            =   "0.00"
            Top             =   240
            Width           =   1755
         End
         Begin VB.TextBox txtPartParts 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Left            =   1860
            MaxLength       =   13
            TabIndex        =   82
            Text            =   "0.00"
            Top             =   660
            Width           =   1755
         End
         Begin VB.TextBox txtPartMaterials 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Left            =   1860
            MaxLength       =   13
            TabIndex        =   81
            Text            =   "0.00"
            Top             =   1080
            Width           =   1755
         End
         Begin VB.TextBox txtPartAccessories 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Left            =   1860
            MaxLength       =   13
            TabIndex        =   80
            Text            =   "0.00"
            Top             =   1500
            Width           =   1755
         End
         Begin VB.TextBox txtPartTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Height          =   315
            Left            =   1860
            MaxLength       =   13
            TabIndex        =   79
            Text            =   "0.00"
            Top             =   1890
            Width           =   1755
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Labor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   50
            Left            =   690
            TabIndex        =   88
            Top             =   270
            Width           =   480
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Parts"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   51
            Left            =   690
            TabIndex        =   87
            Top             =   690
            Width           =   435
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Materials"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   49
            Left            =   690
            TabIndex        =   86
            Top             =   1110
            Width           =   765
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Accessories"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   52
            Left            =   690
            TabIndex        =   85
            Top             =   1530
            Width           =   1050
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   53
            Left            =   690
            TabIndex        =   84
            Top             =   1920
            Width           =   405
         End
      End
      Begin VB.CheckBox chkAllowManDist 
         Caption         =   "Enable Manual Distribution"
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
         Height          =   240
         Left            =   120
         TabIndex        =   77
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtLOAAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   1980
         MaxLength       =   13
         TabIndex        =   76
         Text            =   "0.00"
         Top             =   450
         Width           =   1755
      End
      Begin VB.CommandButton cmdPartClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2040
         TabIndex        =   75
         Top             =   3720
         Width           =   915
      End
      Begin VB.CommandButton cmdPartSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1140
         TabIndex        =   89
         Top             =   3720
         Width           =   915
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   345
         Left            =   0
         TabIndex        =   91
         Top             =   -30
         Width           =   4065
         _Version        =   655364
         _ExtentX        =   7170
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "Input Insurance Participation"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         Alignment       =   1
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "LOA Amount :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   48
         Left            =   90
         TabIndex        =   90
         Top             =   480
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmCSMSEstimateEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsEsti_HD                                           As ADODB.Recordset
Dim rsCusmas                                            As ADODB.Recordset
Dim rsEmpNo                                             As ADODB.Recordset
Dim rsS_Model                                           As ADODB.Recordset
Dim rsROJOBS                                            As ADODB.Recordset
Dim rsROJOBS2                                           As ADODB.Recordset
Dim rsRemarks                                           As ADODB.Recordset
Dim rsEsti_Det                                          As ADODB.Recordset
Dim rsORD_HIST                                          As ADODB.Recordset
Dim rsDAYTRAN                                           As ADODB.Recordset
Dim rsOrd_Hd                                            As ADODB.Recordset
Dim rsTdayTran                                          As ADODB.Recordset
Dim rsPartMas                                           As ADODB.Recordset
Dim rsMatMas                                            As ADODB.Recordset
Dim rsMATISS                                            As ADODB.Recordset
Dim rsDNPP                                              As ADODB.Recordset

Dim promt                                               As Boolean
Dim JobTotal                                            As Double
Dim JobComTotal                                         As Double
Dim JobSalesTotal                                       As Double
Dim JobWarTotal                                         As Double
Dim JobDiscTotal                                        As Double
Dim JobVatTotal                                         As Double

Dim PartsTotal                                          As Double
Dim PartsComTotal                                       As Double
Dim PartsSalesTotal                                     As Double
Dim PartsWarTotal                                       As Double
Dim PartsDiscTotal                                      As Double
Dim PartsVatTotal                                       As Double

Dim MatTotal                                            As Double
Dim MatComTotal                                         As Double
Dim MatSalesTotal                                       As Double
Dim MatWarTotal                                         As Double
Dim MatDiscTotal                                        As Double
Dim MatVatTotal                                         As Double

Dim ACCTotal                                            As Double
Dim ACCComTotal                                         As Double
Dim ACCSalesTotal                                       As Double
Dim ACCWarTotal                                         As Double
Dim ACCDiscTotal                                        As Double
Dim ACCVatTotal                                         As Double

Dim COMTotal                                            As Double
Dim SALESTotal                                          As Double
Dim WARTotal                                            As Double
Dim VATTotal                                            As Double
Dim ROTotal                                             As Double

Dim AddorEdit                                           As String
Dim kcnt                                                As Integer
Dim Mcnt                                                As Integer
Dim Pcnt                                                As Integer
Dim Acnt                                                As Integer
Dim DiscTotal                                           As Double
Dim PrevRoNumber                                        As String

Dim JobInsTotal                                         As Double
Dim PartsInsTotal                                       As Double
Dim MatInsTotal                                         As Double
Dim AccInsTotal                                         As Double
Dim INSTotal                                            As Double

Dim WithEvents frm                                      As frmCSMSROCusveh
Attribute frm.VB_VarHelpID = -1
Dim WithEvents FRMx                                     As frmCSMS_MasterSearchCustomer
Attribute FRMx.VB_VarHelpID = -1
Dim WithEvents frmApp                                   As frmCSMS_UploadEstimate
Attribute frmApp.VB_VarHelpID = -1

Function RECALL_STOREMEMVARS()
    Call StoreMemVars
    frmCSMSQUESTION.lblDISCOUNT.Caption = DiscTotal
End Function

Function AddJobs()
    'SSTab1.Tab = 1
    SSTab1.SelectedItem = 1
    Call SendToBack
    
    fraAddJobs.ZOrder 0
    fraAddJobs.Enabled = True
    AddorEdit = "ADD"
    Call InitJobs
    Call optByCode_Click
    optByCode.Value = True
End Function

Function SetPartDesc(ppp As String)
    Set rsPartMas = New ADODB.Recordset
    Set rsPartMas = gconDMIS.Execute("Select partno,partdesc from PMIS_PartMas where partno = '" & ppp & "'")
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartDesc = Null2String(rsPartMas!PartDesc)
    Else
        SetPartDesc = cboDescription.Text
    End If
End Function

Function setJobCode(jjj As String)
    If jjj <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select jcode,desc1,std_mhrs from CSMS_Jobs where desc1 = '" & jjj & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
            setJobCode = Null2String(rsROJOBS!JCode)
            labJobDet_Vol.Caption = N2Str2IntZero(rsROJOBS!std_mhrs)
        Else
            setJobCode = ""
            labJobDet_Vol.Caption = 0
        End If
    End If
End Function

Function setJobDesc(jjj As String)
    If jjj <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select jcode,desc1,std_mhrs from CSMS_Jobs where jcode = '" & jjj & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
            setJobDesc = Null2String(rsROJOBS!desc1)
            labJobDet_Vol.Caption = N2Str2IntZero(rsROJOBS!std_mhrs)
        Else
            setJobDesc = ""
            labJobDet_Vol.Caption = 0
        End If
    End If
End Function

Function setJobPOcode(ppp As String)
    If ppp <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select desc1,std_mhrs from CSMS_Jobs where desc1 = '" & ppp & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then

            labJobDet_Vol.Caption = N2Str2IntZero(rsROJOBS!std_mhrs)
        Else
            setJobPOcode = ""
            labJobDet_Vol.Caption = 0
        End If
    End If
End Function

Function setJobDetail(ppp As String)
    If ppp <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select desc1,detail,std_mhrs from CSMS_Jobs where desc1 = '" & ppp & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
            setJobDetail = Null2String(rsROJOBS!DETAIL)
            labJobDet_Vol.Caption = N2Str2IntZero(rsROJOBS!std_mhrs)
        Else
            setJobDetail = ""
            labJobDet_Vol.Caption = 0
        End If
    End If
End Function

Function setJobRate(ppp As String)
    If ppp <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select desc1,flatrate,std_mhrs from CSMS_Jobs where desc1 = '" & ppp & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
            setJobRate = Null2String(rsROJOBS!FLATRATE)
            labJobDet_Vol.Caption = N2Str2IntZero(rsROJOBS!std_mhrs)
        Else
            setJobRate = 0#
            labJobDet_Vol.Caption = 0
        End If
    End If
End Function

Function SetMake(mmm As String) As String
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select DESCRIPTION from CSMS_CUSVEH where ltrim(rtrim(PLATE_NO)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        SetMake = Null2String(rsS_Model!Description)
    End If
    Set rsS_Model = Nothing
End Function

Function SetSA(emp As String) As String
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("Select code,naym from CSMS_vw_EmpNo where code = '" & emp & "'")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetSA = Null2String(rsEmpNo!NAYM)
End Function

Function SetCodeSA(nam As String)
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("Select code,naym from CSMS_vw_EmpNo where naym = '" & nam & "'")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetCodeSA = Null2String(rsEmpNo!Code)
End Function

Function StoreJobsEntry(ByVal ID As Variant)
    Dim retVal                                         As Boolean
    Set rsEsti_Det = New ADODB.Recordset
    rsEsti_Det.Open "select det_hrs,id,LINE_NO,detcde,DETDSC,pocode,wcode,det_amt,discrate,estimateno,livil,detail from CSMS_EstDETAILS where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        labDetID.Caption = rsEsti_Det!ID
        txtJobLineNo.Text = Null2String(rsEsti_Det!LINE_NO)
        cboJcode.Text = Null2String(rsEsti_Det!DETCDE)
        cboJobCode.Text = Null2String(rsEsti_Det!DETDSC)
        txtJobPostCode.Text = Null2String(rsEsti_Det!pocode)
        cboJobChargeTo.Text = Null2String(rsEsti_Det!wCode)
        txtHRS.Text = NumericVal(rsEsti_Det!DET_HRS)
        txtJobRate.Text = N2Str2Zero(rsEsti_Det!DET_AMT)
        txtJobDiscount.Text = N2Str2Zero(rsEsti_Det!discrate)
        txtJobDetail.Text = Null2String(rsEsti_Det!DETAIL)
    End If
End Function

Function StorePartsEntry(ByVal ID As Variant)
    Dim retVal                                         As Boolean
    Set rsEsti_Det = New ADODB.Recordset
    rsEsti_Det.Open "select id,LINE_NO,detcde,detdsc,pocode,detvol,detprc,det_amt,wcode,discrate from CSMS_EstDETAILS where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        labDetID.Caption = rsEsti_Det!ID
        txtPartsLineNo.Text = Null2String(rsEsti_Det!LINE_NO)
        cboPartNo.Text = Null2String(rsEsti_Det!DETCDE)
        cboDescription.Text = Null2String(rsEsti_Det!DETDSC)
        txtPartCode.Text = Null2String(rsEsti_Det!pocode)
        txtQTY.Text = N2Str2Zero(rsEsti_Det!detvol)
        txtUnitPrice.Text = Null2String(rsEsti_Det!DetPrc)
        txtPartAmount.Text = N2Str2Zero(rsEsti_Det!DET_AMT)
        cboChargeTo.Text = Null2String(rsEsti_Det!wCode)
        txtPartDiscount.Text = N2Str2IntZero(rsEsti_Det!discrate)
    End If
End Function

Function StoreMatEntry(ByVal ID As String)
    Dim retVal                                         As Boolean
    Set rsEsti_Det = New ADODB.Recordset
    rsEsti_Det.Open "select id,LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discrate,pocode from CSMS_EstDETAILS where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        labDetID.Caption = rsEsti_Det!ID
        txtMatLineNo.Text = Null2String(rsEsti_Det!LINE_NO)
        cboMatCode.Text = Null2String(rsEsti_Det!DETCDE)
        cboMaterial.Text = Null2String(rsEsti_Det!DETDSC)
        txtMatQty.Text = N2Str2Zero(rsEsti_Det!detvol)
        txtMatUnitPrice.Text = Null2String(rsEsti_Det!DetPrc)
        txtMatAmount.Text = N2Str2Zero(rsEsti_Det!DET_AMT)
        cboMatChargeTo.Text = Null2String(rsEsti_Det!wCode)
        txtMatDiscount.Text = N2Str2Zero(rsEsti_Det!discrate)
        txtMatPOCode.Text = Null2String(rsEsti_Det!pocode)
    End If
End Function

Function StoreAccEntry(ByVal ID As Variant)

    Dim retVal                                         As Boolean
    Set rsEsti_Det = New ADODB.Recordset
    rsEsti_Det.Open "select id,LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discrate,pocode from CSMS_EstDETAILS where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        labDetID.Caption = rsEsti_Det!ID
        txtAccLineNo.Text = Null2String(rsEsti_Det!LINE_NO)
        cboAccCode.Text = Null2String(rsEsti_Det!DETCDE)
        cboAccessories.Text = Null2String(rsEsti_Det!DETDSC)
        txtAccQty.Text = N2Str2Zero(rsEsti_Det!detvol)
        txtAccUnitPrice.Text = Null2String(rsEsti_Det!DetPrc)
        txtAccAmount.Text = N2Str2Zero(rsEsti_Det!DET_AMT)
        cboAccChargeTo.Text = Null2String(rsEsti_Det!wCode)
        txtAccDiscount.Text = N2Str2Zero(rsEsti_Det!discrate)
        txtAccPOCODE.Text = Null2String(rsEsti_Det!pocode)
    End If
End Function

Function SetMatCode(mmm As String) As String
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        rsMatMas.Open "select matcde,matdsc from CSMS_MatMas where matdsc = '" & mmm & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then
            SetMatCode = Null2String(rsMatMas!MATCDE)
        Else
            SetMatCode = cboMatCode.Text
        End If
    End If
End Function

Function SetMatDisc(mmm As String) As String
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        rsMatMas.Open "select matcde,matdsc from CSMS_MatMas where matcde = '" & mmm & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then
            SetMatDisc = Null2String(rsMatMas!MatDsc)
        Else
            If mmm = "MISC" Then
                SetMatDisc = "MISCELLANEOUS CHARGES"
            Else
                SetMatDisc = cboMaterial.Text
            End If
        End If
    End If
End Function

Function SetMatPrice(mmm As String) As String
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        rsMatMas.Open "select matcde,s_price from CSMS_MatMas where matcde = '" & mmm & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then
            SetMatPrice = Null2String(rsMatMas!s_price)
        End If
    End If
End Function

Function SetMatPOCode(mmm As String) As String
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        SetMatPOCode = ""
    End If
End Function

Function SetAccCode(mmm As String) As String
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        rsMatMas.Open "select sTOCKNO,STOCKDESC from PMIS_STOCKMAS where STOCKDESC = '" & mmm & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then
            SetAccCode = Null2String(rsMatMas!STOCKNO)
        Else
            SetAccCode = cboAccCode.Text
        End If
    End If
End Function

Function SetAccDisc(mmm As String) As String
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        rsMatMas.Open "select STOCKNO,STOCKDESC from pmis_STOCKMAS where STOCKNO = '" & mmm & "' AND TYPE = 'A'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then
            SetAccDisc = Null2String(rsMatMas!STOCKDESC)
        Else
            SetAccDisc = cboAccessories.Text
        End If
    End If
End Function

Function SetAccPOCode(mmm As String) As String
End Function

Function SetAccPrice(mmm As String) As String
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        rsMatMas.Open "select STOCKNO,SRP from PMIS_STOCKMAS where STOCKNO = '" & mmm & "' AND TYPE = 'A'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then
            SetAccPrice = Null2String(rsMatMas!SRP)
        End If
    End If
End Function

Function SetPartDisc(xx As String) As String
    promt = False
    If xx <> "" Then
        Set rsPartMas = New ADODB.Recordset
        rsPartMas.Open "Select partno,partdesc from PMIS_PartMas where partno = '" & xx & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            SetPartDisc = Null2String(rsPartMas!PartDesc)
        Else
            Set rsDNPP = New ADODB.Recordset
            rsDNPP.Open "Select partnumber,descriptio from pmis_dnpp where partnumber = '" & xx & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not (rsDNPP.EOF And rsDNPP.BOF) Then
                SetPartDisc = Null2String(rsDNPP!DESCRIPTIO)
            End If
        End If
    End If
End Function

Function SetPartPrice(ppp As String) As String
    If ppp <> "" Then
        Set rsPartMas = New ADODB.Recordset
        rsPartMas.Open "Select partno,SRP from PMIS_PartMas where partno = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            SetPartPrice = Null2String(rsPartMas!SRP)
        Else
          Set rsDNPP = New ADODB.Recordset
          rsDNPP.Open "Select Partnumber, srp from pmis_dnpp where Partnumber = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
          If Not (rsDNPP.EOF And rsDNPP.BOF) Then
              SetPartPrice = Null2String(rsDNPP!SRP)
          End If
        End If
    End If
End Function

Sub PRINTESTIDISC()
    Screen.MousePointer = 11
    rptEstimate.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptEstimate.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptEstimate, CSMS_REPORT_PATH & "estimatedisc.rpt", "{esti_hd.estimateno} = '" & txtEstimateno.Text & "'", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Sub PRINTESTI()
    Screen.MousePointer = 11
    rptEstimate.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptEstimate.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptEstimate, CSMS_REPORT_PATH & "estimate.rpt", "{esti_hd.estimateno} = '" & txtEstimateno.Text & "'", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set rsEsti_HD = New ADODB.Recordset
    Set rsEsti_HD = gconDMIS.Execute("select * from CSMS_EstHd order by id asc")
End Sub

Sub initMemvars()
    txtEstimateno.Text = vbNullString
    txtInvoiceNo.Text = vbNullString
    txtROType.Text = "0"
    txtSvc_No.Text = "Q"
    txtAcct_No.Text = vbNullString
    txtNiym.Text = vbNullString
    txtPlate_No.Text = vbNullString
    txtMake.Text = vbNullString
    txtTerm.Text = "CSH"
    txtSektion.Text = vbNullString
    txtKm_rdg.Text = vbNullString
    txtDte_recd.Text = LOGDATE
    txtCertific8.Text = vbNullString
    txtDte_comp.Text = vbNullString
    txtDte_Rel.Text = vbNullString
    txtPart_amt.Text = vbNullString
    txtParticipat.Text = vbNullString
    
    JobTotal = 0:   JobWarTotal = 0:    JobDiscTotal = 0:   JobVatTotal = 0
    PartsTotal = 0: PartsWarTotal = 0:  PartsDiscTotal = 0: PartsVatTotal = 0
    MatTotal = 0:   MatWarTotal = 0:    MatDiscTotal = 0:   MatVatTotal = 0
    ACCTotal = 0:   ACCWarTotal = 0:    ACCDiscTotal = 0:   ACCVatTotal = 0
    ROTotal = 0
    Pcnt = 0
    ESTIKCNT = 0
    Mcnt = 0
    Acnt = 0

    Call clearDetailsgrd
    Call InitGrid
    
    lstJObs.Sorted = True: lstJObs.ListItems.Clear: lstJObs.Refresh
    lstParts.Sorted = True: lstParts.ListItems.Clear: lstParts.Refresh
    lstMaterials.Sorted = True: lstMaterials.ListItems.Clear: lstMaterials.Refresh
End Sub

Sub InitJobs()
    txtJobLineNo.Text = Format(ESTIKCNT + 1, "00")
    
    txtJobPostCode.Text = ""
    cboJobChargeTo.Clear
    cboJobChargeTo.AddItem ""
    cboJobChargeTo.AddItem "W"
    cboJobChargeTo.AddItem "S"
    cboJobChargeTo.AddItem "C"
    txtJobDiscount.Text = "0"
    txtJobDetail.Text = ""
End Sub

Sub InitParts()
    txtPartsLineNo.Text = Format(Pcnt + 1, "00")
    txtPartCode.Text = "01"
    cboPartNo.ListIndex = 0
    txtQTY.Text = 1
    txtUnitPrice.Text = 0#
    txtPartAmount.Text = 0#
    cboChargeTo.Clear
    cboChargeTo.AddItem ""
    cboChargeTo.AddItem "W"
    cboChargeTo.AddItem "S"
    cboChargeTo.AddItem "C"
    txtPartDiscount.Text = 0#
End Sub

Sub InitAccessories()
    On Error Resume Next
    txtAccLineNo.Text = Format(Acnt + 1, "00")
    cboAccCode.ListIndex = 0
    txtAccQty.Text = 1
    txtAccUnitPrice.Text = 0#
    txtAccAmount.Text = 0#
    cboAccChargeTo.Clear
    cboAccChargeTo.AddItem ""
    cboAccChargeTo.AddItem "W"
    cboAccChargeTo.AddItem "S"
    cboAccChargeTo.AddItem "C"
    txtAccDiscount.Text = 0#
    txtAccPOCODE.Text = "01"
End Sub

Sub InitMaterials()
    On Error Resume Next
    txtMatLineNo.Text = Format(Mcnt + 1, "00")
    cboMatCode.ListIndex = 0
    txtMatQty.Text = 1
    txtMatUnitPrice.Text = 0#
    txtMatAmount.Text = 0#
    cboMatChargeTo.Clear
    cboMatChargeTo.AddItem ""
    cboMatChargeTo.AddItem "W"
    cboMatChargeTo.AddItem "S"
    cboMatChargeTo.AddItem "C"
    txtMatDiscount.Text = 0#
    txtMatPOCode.Text = "01"
End Sub

Sub AddParts()
    SSTab1.SelectedItem = 2
    Call SendToBack
    fraAddParts.ZOrder 0
    fraAddParts.Enabled = True
    AddorEdit = "ADD"
    Call InitParts
    
    On Error Resume Next
    cboPartNo.SetFocus
End Sub

Sub AddMaterials()
    SSTab1.SelectedItem = 3
    Call SendToBack
    fraAddMaterials.ZOrder 0
    fraAddMaterials.Enabled = True
    AddorEdit = "ADD"
    Call InitMaterials
    
    On Error Resume Next
    cboMatCode.SetFocus
End Sub

Sub AddAccessories()
    SSTab1.SelectedItem = 4
    Call SendToBack
    
    fraAddAccessories.ZOrder 0
    fraAddAccessories.Enabled = True
    AddorEdit = "ADD"
    InitAccessories
    On Error Resume Next
    cboAccCode.SetFocus
End Sub

Sub SearchEstimateNo(XXX As String)
    Call rsRefresh
    On Error GoTo ErrorCode
    rsEsti_HD.Bookmark = rsFind(rsEsti_HD.Clone, "estimateno", XXX).Bookmark
    SendToBack
    StoreMemVars
    Exit Sub

ErrorCode:
    ShowCantFind XXX
    Resume Next
End Sub

Sub InitCbo()
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("select naym from CSMS_vw_EmpNo order by naym asc")

    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then
        cboRecd_by.Clear
        cboRecd_by.Text = Null2String(rsEmpNo!NAYM)
        Do While Not rsEmpNo.EOF
            cboRecd_by.AddItem Null2String(rsEmpNo!NAYM)
            rsEmpNo.MoveNext
        Loop
    End If

    Set rsPartMas = New ADODB.Recordset
    Set rsPartMas = gconDMIS.Execute("select STOCKNO from PMIS_PartMas where type = 'P' order by STOCKNO asc")
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        rsPartMas.MoveFirst
        cboPartNo.Clear
        Do While Not rsPartMas.EOF
            cboPartNo.AddItem Null2String(rsPartMas!STOCKNO)
            rsPartMas.MoveNext
        Loop
    End If
    
    Set rsPartMas = New ADODB.Recordset
    Set rsPartMas = gconDMIS.Execute("select stockdesc from PMIS_PartMas where type = 'P' order by stockdesc asc")
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        rsPartMas.MoveFirst
        cboDescription.Clear
        Do While Not rsPartMas.EOF
            cboDescription.AddItem Null2String(rsPartMas!STOCKDESC)
            rsPartMas.MoveNext
        Loop
    End If
'--------------------------------------------------------------------------------------------------------------------
    Set rsDNPP = New ADODB.Recordset
    Set rsDNPP = gconDMIS.Execute("select partnumber from PMIS_dnpp where partnumber not in (Select stockdesc from pmis_stockmas where [type] = 'P') order by partnumber asc")
    If Not (rsDNPP.EOF And rsDNPP.BOF) Then
        rsDNPP.MoveFirst
        Do While Not rsDNPP.EOF
            cboPartNo.AddItem Null2String(rsDNPP!partnumber)
            rsDNPP.MoveNext
        Loop
    End If
    
    Set rsDNPP = New ADODB.Recordset
    Set rsDNPP = gconDMIS.Execute("select  descriptio from PMIS_dnpp where  partnumber not in (Select stockdesc from pmis_stockmas where [type] = 'P') order by descriptio asc")
    If Not (rsDNPP.EOF And rsDNPP.BOF) Then
        rsDNPP.MoveFirst
        Do While Not rsDNPP.EOF
            cboDescription.AddItem Null2String(rsDNPP!DESCRIPTIO)
            rsDNPP.MoveNext
        Loop
    End If
'--------------------------------------------------------------------------------------------------------------------
    
    Set rsMatMas = New ADODB.Recordset
    Set rsMatMas = gconDMIS.Execute("select STOCKNO from PMIS_STOCKmas WHERE TYPE = 'M' order by STOCKNO asc")
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        rsMatMas.MoveFirst
        cboMatCode.Clear
        cboMatCode.AddItem Null2String("MISC")
        Do While Not rsMatMas.EOF
            cboMatCode.AddItem Null2String(rsMatMas!STOCKNO)
            rsMatMas.MoveNext
        Loop
    End If

    Set rsMatMas = New ADODB.Recordset
    Set rsMatMas = gconDMIS.Execute("select stockdesc from PMIS_STOCKmas WHERE TYPE = 'M' order by stockdesc asc")
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        rsMatMas.MoveFirst
        cboMaterial.Clear
        Do While Not rsMatMas.EOF
            cboMaterial.AddItem Null2String(rsMatMas!STOCKDESC)
            rsMatMas.MoveNext
        Loop
    End If

    Set rsMatMas = New ADODB.Recordset
    Set rsMatMas = gconDMIS.Execute("select STOCKNO from PMIS_STOCKmas WHERE TYPE = 'A' order by STOCKNO asc")
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        rsMatMas.MoveFirst
        cboAccCode.Clear
        Do While Not rsMatMas.EOF
            cboAccCode.AddItem Null2String(rsMatMas!STOCKNO)
            rsMatMas.MoveNext
        Loop
    End If
    
    Set rsMatMas = New ADODB.Recordset
    Set rsMatMas = gconDMIS.Execute("select stockdesc from PMIS_STOCKmas WHERE TYPE = 'A' order by stockdesc asc")
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        rsMatMas.MoveFirst
        cboAccessories.Clear
        Do While Not rsMatMas.EOF
            cboAccessories.AddItem Null2String(rsMatMas!STOCKDESC)
            rsMatMas.MoveNext
        Loop
    End If
    
End Sub






Sub StoreMemVars()
    If Not rsEsti_HD.EOF And Not rsEsti_HD.BOF Then
        labid.Caption = NumericVal(rsEsti_HD!ID)
        txtInvoiceNo.Text = Null2String(rsEsti_HD!INVOICE)
        txtDte_Rel.Text = Null2String(rsEsti_HD!dte_rel)
        txtEstimateno.Text = Null2String(rsEsti_HD!EstimateNo)
        txtROType.Text = Null2String(rsEsti_HD!ROTYPE)
        txtSvc_No.Text = Null2String(rsEsti_HD!svc_no)
        txtAcct_No.Text = Null2String(rsEsti_HD!ACCT_NO)
        If IsNull(rsEsti_HD!ACCT_NO) = False Then SetAdres (rsEsti_HD!ACCT_NO)
        txtNiym.Text = Null2String(rsEsti_HD!NIYM)
        txtPlate_No.Text = Null2String(rsEsti_HD!PLATE_NO)
        cboModel.Text = Null2String(rsEsti_HD!Model)
        txtMake.Text = SetMake(Null2String(rsEsti_HD!PLATE_NO))
        txtTerm.Text = Null2String(rsEsti_HD!TERM)
        txtSektion.Text = Null2String(rsEsti_HD!sektion)
        cboRecd_by.Text = SetSA(Null2String(rsEsti_HD!RECD_BY))
        txtKm_rdg.Text = Null2String(rsEsti_HD!km_rdg)
        txtDte_recd.Text = Null2String(rsEsti_HD!DTE_RECD)
        txtCertific8.Text = Null2String(rsEsti_HD!certific8)
        txtDte_comp.Text = Null2String(rsEsti_HD!dte_comp)
        txtPart_amt.Text = Null2String(rsEsti_HD!part_amt)
        If Null2String(rsEsti_HD!TRANSTYPE) = "R" Or Left(Null2String(rsEsti_HD!REP_OR), 1) = "R" Then
            lblSTATUS.Caption = "** UPLOADED TO RO **"
            txtROno.Text = Null2String(getro(txtEstimateno.Text))
            Call getinvoiceandpayterm
            cmdUpload.Enabled = False
            cmdEdit.Enabled = False
        Else
            lblSTATUS.Caption = ""
            cmdUpload.Enabled = True
            txtROno.Text = ""
            cmdEdit.Enabled = True
        End If
        If Null2String(rsEsti_HD!participat) = "" Then
            chkParticipat.Value = 0
            txtParticipat.Text = ""
            Text2.Text = ""
        Else
            chkParticipat.Value = 1
            txtParticipat.Text = Null2String(rsEsti_HD!participat)
            Text2.Text = Null2String(rsEsti_HD!INSCDE)
        End If
        
        DoEvents

        Call FillJobs
        Call FillParts
        Call FillMaterials
        Call FillAccessories
        Call FillDetails

        txtLOAAmount.Text = Format(NumericVal(rsEsti_HD!INSAMT), MAXIMUM_DIGIT)
        txtPartLabor.Text = Format(NumericVal(rsEsti_HD!PARTLABOR), MAXIMUM_DIGIT)
        txtPartParts.Text = Format(NumericVal(rsEsti_HD!PARTPARTS), MAXIMUM_DIGIT)
        txtPartMaterials.Text = Format(NumericVal(rsEsti_HD!PARTMATERIALS), MAXIMUM_DIGIT)
        txtPartAccessories.Text = Format(NumericVal(rsEsti_HD!PARTACCESSORIES), MAXIMUM_DIGIT)
        txtPartTotal.Text = Format(NumericVal(rsEsti_HD!INSAMT), MAXIMUM_DIGIT)
        
        DoEvents
    Else
        cmdFirst.Enabled = False: cmdLast.Enabled = False: cmdPrevious.Enabled = False
        cmdNext.Enabled = False:  cmdPrint.Enabled = False
    End If
End Sub

Sub SetAdres(CCC As String)
    Set rsCusmas = New ADODB.Recordset
    Set rsCusmas = gconDMIS.Execute("Select cuscde,cusadd,cusphon1 from ALL_CUSMAS where cuscde = '" & CCC & "'")
    If Not rsCusmas.EOF And Not rsCusmas.BOF Then
        If Null2String(rsCusmas!cusphon1) <> "" Then
            txtAddress.Text = Null2String(rsCusmas!Cusadd) & " --- " & Null2String(rsCusmas!cusphon1)
        Else
            txtAddress.Text = Null2String(rsCusmas!Cusadd)
        End If
    Else
        txtAddress.Text = ""
    End If
End Sub

Sub clearDetailsgrd()
    Dim i, r                                           As Integer
    grdDetails.Rows = 7: grdDetails.Row = 1
    For r = 0 To grdDetails.Rows - 1
        grdDetails.Row = r
        For i = 0 To grdDetails.Cols - 1
            grdDetails.Col = i: grdDetails.Text = ""
        Next
    Next
    grdDetails.Col = 0: grdDetails.Text = "No Entry": grdDetails.Col = 2
End Sub

Sub InitGrid()
    With grdDetails
        .Rows = 8
        .ColWidth(0) = 1350
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 1

        .Row = 0
        .Col = 1: .Text = "Customer"
        .Col = 2: .Text = "Company"
        .Col = 3: .Text = "Sales"
        .Col = 4: .Text = "Warranty"
        .Col = 5: .Text = "Insurance"
        .Col = 6: .Text = "Discount"
        .Col = 7: .Text = "Vat"
        .Col = 8: .Text = "ID"
        .Col = 0

        .Row = 2: .Text = "Labor"
        .Row = 3: .Text = "Parts"
        .Row = 4: .Text = "Materials"
        .Row = 5: .Text = "Accessories"
        .Row = 6: .Text = "TOTAL"
        .Row = 7: .Text = "RO Amount"
    End With
    grdDetails.RemoveItem 1
End Sub

Sub FillDetails()
    Screen.MousePointer = 11
    
    JobInsTotal = N2Str2Zero(rsEsti_HD!PARTLABOR)
    PartsInsTotal = N2Str2Zero(rsEsti_HD!PARTPARTS)
    MatInsTotal = N2Str2Zero(rsEsti_HD!PARTMATERIALS)
    AccInsTotal = N2Str2Zero(rsEsti_HD!PARTACCESSORIES)
    INSTotal = JobInsTotal + PartsInsTotal + MatInsTotal + AccInsTotal
    
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT - INSTotal
    COMTotal = JobComTotal + PartsComTotal + MatComTotal + ACCComTotal
    SALESTotal = JobSalesTotal + PartsSalesTotal + MatSalesTotal + ACCSalesTotal
    WARTotal = JobWarTotal + PartsWarTotal + MatWarTotal + ACCWarTotal
  

    DiscTotal = N2Str2Zero(rsEsti_HD!l_discount) + N2Str2Zero(rsEsti_HD!p_discount) + N2Str2Zero(rsEsti_HD!m_discount) + N2Str2Zero(rsEsti_HD!a_discount)
    VATTotal = N2Str2Zero(rsEsti_HD!l_taxval) + N2Str2Zero(rsEsti_HD!p_taxval) + N2Str2Zero(rsEsti_HD!m_taxval) + N2Str2Zero(rsEsti_HD!A_taxval)
    
    TOTJOBAMT = TOTJOBAMT - JobInsTotal
    TOTPARTSAMT = TOTPARTSAMT - PartsInsTotal
    TOTMATAMT = TOTMATAMT - MatInsTotal
    TOTACCAMT = TOTACCAMT - AccInsTotal
    
    
    Call InitGrid
    Set rsEsti_Det = New ADODB.Recordset
    Set rsEsti_Det = gconDMIS.Execute("select estimateno from CSMS_EstDETAILS where estimateno = " & N2Str2Null(rsEsti_HD!EstimateNo) & " order by LINE_NO asc")
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        With grdDetails
            .Rows = 7
            .Col = 1
            .Row = 1
            .Text = Format(TOTJOBAMT - N2Str2Zero(rsEsti_HD!l_discount), MAXIMUM_DIGIT)
            .Row = 2
            .Text = Format(TOTPARTSAMT - N2Str2Zero(rsEsti_HD!p_discount), MAXIMUM_DIGIT)
            .Row = 3
            .Text = Format(TOTMATAMT - N2Str2Zero(rsEsti_HD!m_discount), MAXIMUM_DIGIT)
            .Row = 4
            .Text = Format(TOTACCAMT - N2Str2Zero(rsEsti_HD!a_discount), MAXIMUM_DIGIT)
            .Row = 5
            .Text = Format(ROTotal - DiscTotal, MAXIMUM_DIGIT)
            .Row = 6
            .Text = Format(N2Str2Zero(ROTotal + INSTotal + COMTotal + SALESTotal + WARTotal), MAXIMUM_DIGIT)

            .Col = 2
            .Row = 1
            .Text = Format(JobComTotal, MAXIMUM_DIGIT)
            .Row = 2
            .Text = Format(PartsComTotal, MAXIMUM_DIGIT)
            .Row = 3
            .Text = Format(MatComTotal, MAXIMUM_DIGIT)
            .Row = 4
            .Text = Format(ACCComTotal, MAXIMUM_DIGIT)
            .Row = 5
            .Text = Format(COMTotal, MAXIMUM_DIGIT)

            .Col = 3
            .Row = 1
            .Text = Format(JobSalesTotal, MAXIMUM_DIGIT)
            .Row = 2
            .Text = Format(PartsSalesTotal, MAXIMUM_DIGIT)
            .Row = 3
            .Text = Format(MatSalesTotal, MAXIMUM_DIGIT)
            .Row = 4
            .Text = Format(ACCSalesTotal, MAXIMUM_DIGIT)
            .Row = 5
            .Text = Format(SALESTotal, MAXIMUM_DIGIT)

            .Col = 4
            .Row = 1
            .Text = Format(JobWarTotal, MAXIMUM_DIGIT)
            .Row = 2
            .Text = Format(PartsWarTotal, MAXIMUM_DIGIT)
            .Row = 3
            .Text = Format(MatWarTotal, MAXIMUM_DIGIT)
            .Row = 4
            .Text = Format(ACCWarTotal, MAXIMUM_DIGIT)
            .Row = 5
            .Text = Format(WARTotal, MAXIMUM_DIGIT)

            .Col = 5
            .Row = 1
            .Text = Format(N2Str2Zero(rsEsti_HD!PARTLABOR), MAXIMUM_DIGIT)
            .Row = 2
            .Text = Format(N2Str2Zero(rsEsti_HD!PARTPARTS), MAXIMUM_DIGIT)
            .Row = 3
            .Text = Format(N2Str2Zero(rsEsti_HD!PARTMATERIALS), MAXIMUM_DIGIT)
            .Row = 4
            .Text = Format(N2Str2Zero(rsEsti_HD!PARTACCESSORIES), MAXIMUM_DIGIT)
            .Row = 5
            .Text = Format(N2Str2Zero(rsEsti_HD!INSAMT), MAXIMUM_DIGIT)
            
            .Col = 6
            .Row = 1
            .Text = Format(N2Str2Zero(rsEsti_HD!l_discount), MAXIMUM_DIGIT)
            .Row = 2
            .Text = Format(N2Str2Zero(rsEsti_HD!p_discount), MAXIMUM_DIGIT)
            .Row = 3
            .Text = Format(N2Str2Zero(rsEsti_HD!m_discount), MAXIMUM_DIGIT)
            .Row = 4
            .Text = Format(N2Str2Zero(rsEsti_HD!a_discount), MAXIMUM_DIGIT)
            .Row = 5
            .Text = Format(DiscTotal, MAXIMUM_DIGIT)

            .Col = 7
            .Row = 1
            .Text = Format(N2Str2Zero(rsEsti_HD!l_taxval), MAXIMUM_DIGIT)
            .Row = 2
            .Text = Format(N2Str2Zero(rsEsti_HD!p_taxval), MAXIMUM_DIGIT)
            .Row = 3
            .Text = Format(N2Str2Zero(rsEsti_HD!m_taxval), MAXIMUM_DIGIT)
            .Row = 4
            .Text = Format(N2Str2Zero(rsEsti_HD!A_taxval), MAXIMUM_DIGIT)
            .Row = 5
            .Text = Format(VATTotal, MAXIMUM_DIGIT)
        End With
    Else
        Call clearDetailsgrd
        Call InitGrid
    End If
    Screen.MousePointer = 0
End Sub

Sub FillJobs()
    Me.lstJObs.Sorted = True: Me.lstJObs.ListItems.Clear
    Dim Item                                           As ListItem
    Set rsEsti_Det = New ADODB.Recordset
    Set rsEsti_Det = gconDMIS.Execute("select id,LINE_NO,detcde,detdsc,det_amt,wcode,discount_2 from CSMS_estdetails where estimateno = " & N2Str2Null(rsEsti_HD!EstimateNo) & " and livil = '1' order by LINE_NO asc")
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        Do While Not rsEsti_Det.EOF
            Set Item = lstJObs.ListItems.Add(, , Null2String(rsEsti_Det!ID))
            Item.SubItems(1) = Null2String(rsEsti_Det!LINE_NO)
            Item.SubItems(2) = Null2String(rsEsti_Det!DETCDE)
            Item.SubItems(3) = Null2String(rsEsti_Det!DETDSC)
            Item.SubItems(4) = Format(NumericVal(rsEsti_Det!DET_AMT), MAXIMUM_DIGIT)
            Item.SubItems(5) = Null2String(rsEsti_Det!wCode)
            Item.SubItems(6) = Format(NumericVal(rsEsti_Det!Discount_2), MAXIMUM_DIGIT)
            rsEsti_Det.MoveNext
        Loop
    End If

    Me.lstJObs.Sorted = False: Me.lstJObs.Refresh
    TOTJOBAMT = 0: TOTJOBDISC = 0: TOTJOBDISCVAL = 0: TOTJOBTAX = 0
    ESTIKCNT = 0: JobComTotal = 0: JobSalesTotal = 0: JobWarTotal = 0
    
    Set rsEsti_Det = New ADODB.Recordset
    Set rsEsti_Det = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_EstDETAILS where estimateno = " & N2Str2Null(rsEsti_HD!EstimateNo) & " and livil = '1' order by LINE_NO asc")
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        Screen.MousePointer = 11
        rsEsti_Det.MoveFirst
        Do While Not rsEsti_Det.EOF
            ESTIKCNT = ESTIKCNT + 1
            If Null2String(rsEsti_Det!wCode) = "C" Then
                JobComTotal = JobComTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            ElseIf Null2String(rsEsti_Det!wCode) = "S" Then JobSalesTotal = JobSalesTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            ElseIf Null2String(rsEsti_Det!wCode) = "W" Then JobWarTotal = JobWarTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            Else
                TOTJOBAMT = TOTJOBAMT + N2Str2Zero(rsEsti_Det!DET_AMT)
                TOTJOBDISC = TOTJOBDISC + N2Str2Zero(rsEsti_Det!Discount_2)
                TOTJOBDISCVAL = TOTJOBDISCVAL + N2Str2Zero(rsEsti_Det!disval)
                TOTJOBTAX = TOTJOBTAX + N2Str2Zero(rsEsti_Det!TAXVAL)
            End If
            rsEsti_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsEsti_Det = Nothing
    TOTJOBAMT = Round(TOTJOBAMT, 2)
    TOTJOBDISC = Round(TOTJOBDISC, 2)
    TOTJOBDISCVAL = Round(TOTJOBDISCVAL, 2)
    TOTJOBTAX = Round(TOTJOBTAX, 2)
End Sub

Sub FillParts()
    Me.lstParts.Sorted = True: Me.lstParts.ListItems.Clear
    Dim Item                                           As ListItem

    Set rsEsti_Det = New ADODB.Recordset
    Set rsEsti_Det = gconDMIS.Execute("select id,LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discount_2 from CSMS_estdetails where estimateno = '" & rsEsti_HD!EstimateNo & "' and livil = '2' order by LINE_NO asc")
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        Do While Not rsEsti_Det.EOF
            Set Item = lstParts.ListItems.Add(, , Null2String(rsEsti_Det!ID))
            Item.SubItems(1) = Null2String(rsEsti_Det!LINE_NO)
            Item.SubItems(2) = Null2String(rsEsti_Det!DETCDE)
            Item.SubItems(3) = Null2String(rsEsti_Det!DETDSC)
            Item.SubItems(4) = NumericVal(rsEsti_Det!detvol)
            Item.SubItems(5) = Format(NumericVal(rsEsti_Det!DetPrc))
            Item.SubItems(6) = Format(NumericVal(rsEsti_Det!DET_AMT))
            Item.SubItems(7) = Null2String(rsEsti_Det!wCode)
            Item.SubItems(8) = Format(NumericVal(rsEsti_Det!Discount_2))


            rsEsti_Det.MoveNext
        Loop
    End If
    Me.lstParts.Sorted = False: Me.lstParts.Refresh
    TOTPARTSAMT = 0: TOTPARTSDISC = 0: TOTPARTSDISCVAL = 0: TOTPARTSTAX = 0
    Pcnt = 0: PartsComTotal = 0: PartsSalesTotal = 0: PartsWarTotal = 0
    
    Set rsEsti_Det = New ADODB.Recordset
    Set rsEsti_Det = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_EstDETAILS where estimateno = '" & rsEsti_HD!EstimateNo & "' and livil = '2' order by LINE_NO asc")
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        rsEsti_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsEsti_Det.EOF
            Pcnt = Pcnt + 1
            If Null2String(rsEsti_Det!wCode) = "C" Then
                PartsComTotal = PartsComTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            ElseIf Null2String(rsEsti_Det!wCode) = "S" Then PartsSalesTotal = PartsSalesTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            ElseIf Null2String(rsEsti_Det!wCode) = "W" Then PartsWarTotal = PartsWarTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            Else
                TOTPARTSAMT = TOTPARTSAMT + N2Str2Zero(rsEsti_Det!DET_AMT)
                TOTPARTSDISC = TOTPARTSDISC + N2Str2Zero(rsEsti_Det!Discount_2)
                TOTPARTSDISCVAL = TOTPARTSDISCVAL + N2Str2Zero(rsEsti_Det!disval)
                TOTPARTSTAX = TOTPARTSTAX + N2Str2Zero(rsEsti_Det!TAXVAL)
            End If
            rsEsti_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsEsti_Det = Nothing
    TOTPARTSAMT = Round(TOTPARTSAMT, 2)
    TOTPARTSDISC = Round(TOTPARTSDISC, 2)
    TOTPARTSDISCVAL = Round(TOTPARTSDISCVAL, 2)
    TOTPARTSTAX = Round(TOTPARTSTAX, 2)
End Sub

Sub FillMaterials()
    Me.lstMaterials.Sorted = True: Me.lstMaterials.ListItems.Clear
    Set rsEsti_Det = New ADODB.Recordset
    
    Set rsEsti_Det = gconDMIS.Execute("select LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discount_2,id from CSMS_estdetails where estimateno = '" & rsEsti_HD!EstimateNo & "' and livil = '3' order by LINE_NO asc")
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        Listview_Loadval Me.lstMaterials.ListItems, rsEsti_Det
    End If

    Me.lstMaterials.Sorted = False: Me.lstMaterials.Refresh
    TOTMATAMT = 0: TOTMATDISC = 0: TOTMATDISCVAL = 0: TOTMATTAX = 0
    Mcnt = 0: MatComTotal = 0: MatSalesTotal = 0: MatWarTotal = 0
    
    Set rsEsti_Det = New ADODB.Recordset
    Set rsEsti_Det = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_EstDETAILS where estimateno = '" & rsEsti_HD!EstimateNo & "' and livil = '3' order by LINE_NO asc")
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        Screen.MousePointer = 11
        rsEsti_Det.MoveFirst
        Do While Not rsEsti_Det.EOF
            Mcnt = Mcnt + 1
            If Null2String(rsEsti_Det!wCode) = "C" Then
                MatComTotal = MatComTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            ElseIf Null2String(rsEsti_Det!wCode) = "S" Then MatSalesTotal = MatSalesTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            ElseIf Null2String(rsEsti_Det!wCode) = "W" Then MatWarTotal = MatWarTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            Else
                TOTMATAMT = TOTMATAMT + N2Str2Zero(rsEsti_Det!DET_AMT)
                TOTMATDISC = TOTMATDISC + N2Str2Zero(rsEsti_Det!Discount_2)
                TOTMATDISCVAL = TOTMATDISCVAL + N2Str2Zero(rsEsti_Det!disval)
                TOTMATTAX = TOTMATTAX + N2Str2Zero(rsEsti_Det!TAXVAL)
            End If
            rsEsti_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsEsti_Det = Nothing
    TOTMATAMT = Round(TOTMATAMT, 2)
    TOTMATDISC = Round(TOTMATDISC, 2)
    TOTMATDISCVAL = Round(TOTMATDISCVAL, 2)
    TOTMATTAX = Round(TOTMATTAX, 2)
End Sub

Sub FillAccessories()
    Me.lstAccessories.Sorted = True: Me.lstAccessories.ListItems.Clear
    
    Set rsEsti_Det = New ADODB.Recordset
    Set rsEsti_Det = gconDMIS.Execute("select id,LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discount_2 from CSMS_estdetails where estimateno = '" & rsEsti_HD!EstimateNo & "' and livil = '4' order by LINE_NO asc")
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        Listview_Loadval Me.lstAccessories.ListItems, rsEsti_Det
    End If

    Me.lstAccessories.Sorted = False: Me.lstAccessories.Refresh
    TOTACCAMT = 0: TOTACCDISC = 0: TOTACCDISCVAL = 0: TOTACCTAX = 0
    Acnt = 0: ACCComTotal = 0: ACCSalesTotal = 0: ACCWarTotal = 0
    
    Set rsEsti_Det = New ADODB.Recordset
    Set rsEsti_Det = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_EstDETAILS where estimateno = '" & rsEsti_HD!EstimateNo & "' and livil = '4' order by LINE_NO asc")
    If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
        Screen.MousePointer = 11
        rsEsti_Det.MoveFirst
        Do While Not rsEsti_Det.EOF
            Acnt = Acnt + 1
            If Null2String(rsEsti_Det!wCode) = "C" Then
                ACCComTotal = ACCComTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            ElseIf Null2String(rsEsti_Det!wCode) = "S" Then ACCSalesTotal = ACCSalesTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            ElseIf Null2String(rsEsti_Det!wCode) = "W" Then ACCWarTotal = ACCWarTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
            Else
                TOTACCAMT = TOTACCAMT + N2Str2Zero(rsEsti_Det!DET_AMT)
                TOTACCDISC = TOTACCDISC + N2Str2Zero(rsEsti_Det!Discount_2)
                TOTACCDISCVAL = TOTACCDISCVAL + N2Str2Zero(rsEsti_Det!disval)
                TOTACCTAX = TOTACCTAX + N2Str2Zero(rsEsti_Det!TAXVAL)
            End If
            rsEsti_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsEsti_Det = Nothing
    TOTACCAMT = Round(TOTACCAMT, 2)
    TOTACCDISC = Round(TOTACCDISC, 2)
    TOTACCDISCVAL = Round(TOTACCDISCVAL, 2)
    TOTACCTAX = Round(TOTACCTAX, 2)
End Sub

Sub SendToBack()
    fraAddJobs.ZOrder 1
    fraAddJobs.Enabled = False
    fraAddParts.ZOrder 1
    fraAddParts.Enabled = False
    fraAddMaterials.ZOrder 1
    fraAddMaterials.Enabled = False
    fraAddAccessories.ZOrder 1
    fraAddAccessories.Enabled = False
End Sub

Sub SendToBackDisc()
    fraDiscount.ZOrder 1: txtDiscAmt.Text = 0
End Sub

Sub SendToFrontDisc()
    fraDiscount.ZOrder 0
    fraDiscount.Visible = True
    txtDiscAmt.Text = 5
    On Error Resume Next
    txtDiscAmt.SetFocus
End Sub

Private Sub cboAccCode_LostFocus()
    cboAccCode.Text = UCase(cboAccCode.Text)
    cboAccessories.Text = SetAccDisc(cboAccCode.Text)
    txtAccUnitPrice.Text = SetAccPrice(cboAccCode.Text)
    txtAccPOCODE.Text = ""                            'SetMatPOCode(cboMatCode.Text)
    txtAccAmount.Text = txtAccUnitPrice.Text
End Sub

Private Sub cboAccessories_LostFocus()
    If cboAccessories.Text <> "" Then
        cboAccCode.Text = SetAccCode(cboAccessories.Text)
        txtAccUnitPrice.Text = SetAccPrice(cboAccCode.Text)
        txtAccPOCODE.Text = ""
        txtAccAmount.Text = txtAccUnitPrice.Text
    End If
End Sub

Private Sub chkAllowManDist_Click()
    If chkAllowManDist.Value = 1 Then
        fraParticipation.Enabled = True
        txtLOAAmount.Enabled = False
    Else
        fraParticipation.Enabled = False
        txtLOAAmount.Enabled = True
    End If
End Sub

Function StoreParticipationEntry(ByVal RO_NO As String)
    Dim rsRO_DET                        As New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select * from CSMS_ESTHD where ESTIMATENO = " & N2Str2Null(RO_NO) & "")
    If Not (rsRO_DET.EOF And rsRO_DET.BOF) Then
        chkAllowManDist.Value = 0
        fraParticipation.Enabled = False
        txtLOAAmount.Enabled = True
        txtPartLabor.Text = ToDoubleNumber(N2Str2Zero(rsRO_DET!PARTLABOR))
        txtPartParts.Text = ToDoubleNumber(N2Str2Zero(rsRO_DET!PARTPARTS))
        txtPartMaterials.Text = ToDoubleNumber(N2Str2Zero(rsRO_DET!PARTMATERIALS))
        txtPartAccessories.Text = ToDoubleNumber(N2Str2Zero(rsRO_DET!PARTACCESSORIES))
        Call SetTotalParticipation
        txtLOAAmount.Text = txtPartTotal.Text
    End If
    Set rsRO_DET = Nothing
End Function

Sub SetTotalParticipation()
    txtPartTotal.Text = ToDoubleNumber(NumericVal(txtPartLabor.Text) + NumericVal(txtPartParts.Text) + NumericVal(txtPartMaterials.Text) + NumericVal(txtPartAccessories.Text))
    If chkAllowManDist.Value = 1 Then
        txtLOAAmount.Text = ToDoubleNumber(txtPartTotal.Text)
        If NumericVal(txtLOAAmount.Text) > ROTotal + INSTotal Then
            MsgBox "Warning: LOA Amount should not Exceed Repair Order Total Amount.", vbCritical, "Not Allowed!"
            txtLOAAmount.Text = NumericVal(txtPartTotal.Text)
            cmdPartSave.Enabled = False
            Exit Sub
        Else
            txtPartTotal.Text = ToDoubleNumber(NumericVal(txtPartLabor.Text) + NumericVal(txtPartParts.Text) + NumericVal(txtPartMaterials.Text) + NumericVal(txtPartAccessories.Text))
            txtLOAAmount.Text = ToDoubleNumber(txtPartTotal.Text)
            cmdPartSave.Enabled = True
        End If
    Else
        txtPartTotal.Text = ToDoubleNumber(NumericVal(txtPartLabor.Text) + NumericVal(txtPartParts.Text) + NumericVal(txtPartMaterials.Text) + NumericVal(txtPartAccessories.Text))
        cmdPartSave.Enabled = True
    End If
End Sub

Private Sub cmdAccCancel_Click()
    Call SendToBack
    cmdCancel.Value = True
    
    Call EnableFrame(True)
End Sub

Private Sub cmdAccDelete_Click()
    If MsgBox("Delete This Materials, Are you Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
        
    SQL_STATEMENT = "delete from CSMS_EstDETAILS where id = " & labDetID.Caption
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT----------------------------------------------------------
        Call NEW_LogAudit("XX", "JOB ESTIMATE", SQL_STATEMENT, labid, "ACC", "ACC CODE: " & cboAccCode, "", labDetID)
    'NEW LOG AUDIT----------------------------------------------------------
    Call ShowDeletedMsg

    Dim cnt                                            As Integer
    Dim rsEsti_detDup                                  As New ADODB.Recordset
    Set rsEsti_detDup = gconDMIS.Execute("select id,LINE_NO from CSMS_EstDETAILS where estimateno = " & N2Str2Null(rsEsti_HD!EstimateNo) & " and livil = '4' order by LINE_NO asc")
    If Not rsEsti_detDup.EOF And Not rsEsti_detDup.BOF Then
        cnt = 0
        rsEsti_detDup.MoveFirst
        Do While Not rsEsti_detDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update CSMS_EstDETAILS set LINE_NO = " & N2Str2Null(Format(cnt, "00")) & " where id = " & rsEsti_detDup!ID
            rsEsti_detDup.MoveNext
        Loop
    End If

    Call FillAccessories
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_Esti_Hd set" & _
        " ACCESSORIES = " & TOTACCAMT - TOTACCTAX & "," & _
        " A_amtvalue = " & TOTACCAMT & "," & _
        " A_disc = " & TOTACCDISCVAL & "," & _
        " A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
        " A_taxval = " & TOTACCTAX & "," & _
        " A_discount = " & TOTACCDISC & "," & _
        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
        " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
        " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------------------------
        Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "ACC", "EST NO: " & txtEstimateno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------------------------

    Call rsRefresh
    rsEsti_HD.Find "id = " & labid.Caption
    cmdMatCancel.Value = True
End Sub

Private Sub cmdAccSave_Click()
    Screen.MousePointer = 11

    Dim ACCESTIMATENO                                   As String
    Dim ACCLEVEL                                        As String
    Dim ACCLINE_NO                                      As String
    Dim ACCDETCDE                                       As String
    Dim ACCDETDSC                                       As String
    Dim ACCDETUNT                                       As String
    Dim ACCDETVOL                                       As Double
    Dim ACCDETPRC                                       As Double
    Dim ACCDETAMT                                       As Double
    Dim ACCCODE                                         As String
    Dim ACCWCODE                                        As String
    Dim ACCTAXRATE                                      As Double
    Dim ACCDISCRATE                                     As Double
    Dim ACCTAXVAL                                       As Double
    Dim ACCDISVAL                                       As Double
    Dim ACCPOCODE                                       As String
    Dim ACCRep_Or2                                      As String
    Dim ACCDETAIL                                       As String
    Dim ACCDET_AMT                                      As Double
    Dim ACCDIS_VAL                                      As Double
    Dim ACCDISCOUNT_2                                   As Double


    If RTrim(LTrim(cboAccCode.Text)) = "" Then
        MsgBox "Accessory number cannot be blank", vbInformation + vbOKOnly
        cboAccCode.SetFocus
        Exit Sub
    End If

    ACCDISVAL = 0: ACCTAXVAL = 0: ACCDETAMT = 0
    ACCDIS_VAL = 0: ACCDISCOUNT_2 = 0: ACCDISCRATE = 0

    ACCESTIMATENO = N2Str2Null(txtEstimateno.Text)
    ACCLEVEL = "'4'"
    ACCLINE_NO = N2Str2Null(Format(txtAccLineNo.Text, "00"))
    ACCDETCDE = N2Str2Null(cboAccCode.Text)
    ACCDETDSC = N2Str2Null(cboAccessories.Text)
    ACCDETUNT = "NULL"
    ACCDETVOL = NumericVal(txtAccQty.Text)
    ACCDETPRC = NumericVal(txtAccUnitPrice.Text)
    ACCDETAMT = NumericVal(txtAccAmount.Text) / ConvertToBIRDecimalFormat(VAT_RATE)
    ACCCODE = "NULL"
    ACCWCODE = N2Str2Null(cboAccChargeTo.Text)
    ACCTAXRATE = (VAT_RATE / 100)
    ACCDISCRATE = NumericVal(txtAccDiscount.Text) / 100
    ACCDISVAL = (ACCDETPRC * ACCDISCRATE) - ((ACCDETPRC * ACCDISCRATE) * ACCTAXRATE)
    ACCPOCODE = N2Str2Null(txtAccPOCODE.Text)
    ACCRep_Or2 = "NULL"
    ACCDETAIL = "NULL"
    ACCDET_AMT = NumericVal(txtAccAmount.Text)
    ACCDIS_VAL = ACCDISVAL * ACCTAXRATE
    ACCDISCOUNT_2 = ACCDET_AMT * ACCDISCRATE
    ACCTAXVAL = (ACCDET_AMT - ACCDISCOUNT_2) * ACCTAXRATE
    If NumericVal(txtAccDiscount.Text) > 0 Then
        If ACCDISCOUNT_2 > TOTACCAMT Then
            MsgBox "Invalid Discount.", vbInformation + vbOKOnly
            On Error Resume Next
            txtAccDiscount.SetFocus
            Exit Sub
        End If
    End If
    
    If gconDMIS.Execute("Select count(*) from pmis_stockmas where [type] = 'A' and stockno = " & ACCDETCDE & "").Fields(0).Value = 0 Then
        MsgBox "Accessory number does't exist.", vbInformation + vbOKOnly
        cboAccCode.SetFocus
        Exit Sub
    End If
    
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into CSMS_EstDETAILS " & _
            "(TRANSTYPE,estimateno,livil,LINE_NO,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,Rep_Or2,detail,det_amt,dis_val,discount_2, REP_OR)" & _
            " values ('E'," & ACCESTIMATENO & ", " & ACCLEVEL & ", " & ACCLINE_NO & "," & _
            " " & ACCDETCDE & "," & ACCDETDSC & "," & _
            " " & ACCDETUNT & ", " & ACCDETVOL & "," & _
            " " & ACCDETPRC & ", " & ACCDETAMT & ", " & ACCCODE & _
            ", " & ACCWCODE & ", " & ACCTAXRATE * 100 & ", " & ACCDISCRATE * 100 & _
            ", " & ACCTAXVAL & ", " & ACCDISVAL & ", " & ACCPOCODE & _
            ", " & ACCRep_Or2 & ", " & ACCDETAIL & ", " & ACCDET_AMT & _
            ", " & ACCDIS_VAL & ", " & ACCDISCOUNT_2 & _
            ", " & ACCESTIMATENO & ")"
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT---------------------------------------------------------
            Call NEW_LogAudit("AA", "JOB ESTIMATE", SQL_STATEMENT, labid, "ACC", "ACC CODE: " & cboAccCode, "", "")
        'NEW LOG AUDIT---------------------------------------------------------

        Call ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CSMS_EstDETAILS set" & _
            " REP_OR = " & ACCESTIMATENO & _
            ", estimateno = " & ACCESTIMATENO & "," & _
            " livil = " & ACCLEVEL & "," & _
            " LINE_NO = " & ACCLINE_NO & "," & _
            " detcde = " & ACCDETCDE & "," & _
            " detdsc = " & ACCDETDSC & "," & _
            " detunt = " & ACCDETUNT & "," & _
            " detvol = " & ACCDETVOL & "," & _
            " detprc = " & ACCDETPRC & "," & _
            " detamt = " & ACCDETAMT & "," & _
            " code = " & ACCCODE & "," & _
            " wcode = " & ACCWCODE & "," & _
            " taxrate = " & ACCTAXRATE * 100 & "," & _
            " discrate = " & ACCDISCRATE * 100 & "," & _
            " taxval = " & ACCTAXVAL & "," & _
            " disval = " & ACCDISVAL & "," & _
            " pocode = " & ACCPOCODE & "," & _
            " Rep_Or2 = " & ACCRep_Or2 & "," & _
            " detail = " & ACCDETAIL & "," & _
            " det_amt = " & ACCDET_AMT & "," & _
            " dis_val = " & ACCDIS_VAL & "," & _
            " discount_2 = " & ACCDISCOUNT_2 & _
            " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT---------------------------------------------------------
            Call NEW_LogAudit("EE", "JOB ESTIMATE", SQL_STATEMENT, labid, "ACC", "ACC CODE: " & cboAccCode, "", "")
        'NEW LOG AUDIT---------------------------------------------------------

        Call ShowSuccessFullyUpdated
    End If

    Call FillAccessories
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_EstHd set" & _
        " ACCESSORIES = " & TOTACCAMT - TOTACCTAX & "," & _
        " A_amtvalue = " & TOTACCAMT & "," & _
        " A_disc = " & TOTACCDISCVAL & "," & _
        " A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
        " A_taxval = " & TOTACCTAX & "," & _
        " A_discount = " & TOTACCDISC & "," & _
        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
        " wm_amt = " & 0 & "," & _
        " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
        " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT---------------------------------------------------------
        Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "ACC", "EST NO: " & txtEstimateno, "", "")
    'NEW LOG AUDIT---------------------------------------------------------

    Call rsRefresh
    rsEsti_HD.Find "id = " & labid.Caption
    Call StoreMemVars
    cmdAccCancel.Value = True
    Screen.MousePointer = 0

    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
    Exit Sub
End Sub

Private Sub cmdCancelDisk_Click()
    Call SendToBackDisc
    
    Call EnableFrame(True)
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "acess_delete", "JOB ESTIMATE") = False Then Exit Sub
    
    If lblSTATUS = "** UPLOADED TO RO **" Then
        MsgBox "You cannot delete this estimate. Estimate already uploaded to Repair order.", vbInformation, "Info."
    Else
        If CheckEstimateStatus(txtEstimateno) = "UPLOADED" Then
            MsgBox "You cannot delete this estimate. Estimate already uploaded to Repair order. " & vbCrLf & " kindly refresh your sJob Estimate Module to display fresh data.", vbInformation, "Info."
            Exit Sub
        ElseIf CheckEstimateStatus(txtEstimateno) = "NOT FOUND" Then
            MsgBox "Estimate Record not found. kindly refresh your Job Estimate module to display fresh data", vbCritical, "Info"
            Exit Sub
        ElseIf CheckEstimateStatus(txtEstimateno) = "NOT UPLOADED" Then
            If MsgBox("Delete this estimate record", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
            
            gconDMIS.Execute ("DELETE FROM CSMS_ESTHD WHERE ESTIMATENO = '" & txtEstimateno & "'")
            gconDMIS.Execute ("DELETE FROM CSMS_ESTDETAILS WHERE ESTIMATENO = '" & txtEstimateno & "'")
            gconDMIS.Execute ("DELETE FROM CSMS_REPOR WHERE ESTIMATENO = '" & txtEstimateno & "'")
            gconDMIS.Execute ("DELETE FROM CSMS_REPAIRORDER WHERE ESTIMATENO = '" & txtEstimateno & "'")
            gconDMIS.Execute ("DELETE FROM CSMS_RO_DET WHERE ESTIMATENO = '" & txtEstimateno & "'")
            gconDMIS.Execute ("DELETE FROM CSMS_PMS_JOB_DET WHERE ESTIMATENO = '" & txtEstimateno & "'")
            
            Call ShowDeletedMsg
            Call Refresh
            Call StoreMemVars
        End If
    End If
End Sub

Private Sub cmdOkDisc_Click()
    Screen.MousePointer = 11
    Dim varESTIMATENO                                  As String
    Dim varDISVAL                                      As Double
    
    varESTIMATENO = N2Str2Null(rsEsti_HD!EstimateNo)
    varDISVAL = NumericVal(txtDiscAmt.Text) / 100

    If SSTab1.SelectedItem = 1 Then
        Dim JOBID                                       As Long
        Dim JOBESTIMATENO                               As String
        Dim JOBLEVEL                                    As String
        Dim JOBLINE_NO                                  As String
        Dim JOBDETCDE                                   As String
        Dim JOBDETDSC                                   As String
        Dim JOBDETUNT                                   As String
        Dim JOBDETVOL                                   As Double
        Dim JOBDETPRC                                   As Double
        Dim JOBDETAMT                                   As Double
        Dim JOBCODE                                     As String
        Dim JOBWCODE                                    As String
        Dim JOBTAXRATE                                  As Double
        Dim JOBDISCRATE                                 As Double
        Dim JOBTAXVAL                                   As Double
        Dim JOBDISVAL                                   As Double
        Dim JOBPOCODE                                   As String
        Dim JOBRep_Or2                                  As String
        Dim JOBDETAIL                                   As String
        Dim JOBDET_AMT                                  As Double
        Dim JOBDIS_VAL                                  As Double
        Dim JOBDISCOUNT_2                               As Double
        Dim JOBREMARKS                                  As String

        Set rsEsti_Det = New ADODB.Recordset
        Set rsEsti_Det = gconDMIS.Execute("Select * from CSMS_EstDETAILS where estimateno = " & varESTIMATENO & " and livil = '1' order by LINE_NO asc")
        If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
            rsEsti_Det.MoveFirst
            Do While Not rsEsti_Det.EOF
                JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
                JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0

                JOBID = rsEsti_Det!ID
                JOBESTIMATENO = varESTIMATENO
                JOBLEVEL = "'1'"
                JOBLINE_NO = Format(N2Str2Null(rsEsti_Det!LINE_NO), "00")
                JOBDETCDE = N2Str2Null(rsEsti_Det!DETCDE)
                JOBDETDSC = N2Str2Null(rsEsti_Det!DETDSC)
                JOBDETUNT = N2Str2Null(rsEsti_Det!detunt)
                JOBDETVOL = N2Str2IntZero(rsEsti_Det!detvol)
                JOBDETPRC = N2Str2Zero(rsEsti_Det!DetPrc)
                JOBCODE = N2Str2Null(rsEsti_Det!Code)
                JOBWCODE = N2Str2Null(rsEsti_Det!wCode)
                JOBTAXRATE = (VAT_RATE / 100)
                JOBDISCRATE = varDISVAL
                JOBDETAMT = JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE)
                JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
                JOBPOCODE = N2Str2Null(rsEsti_Det!pocode)
                JOBRep_Or2 = "NULL"
                JOBDETAIL = N2Str2Null(rsEsti_Det!DETAIL)
                JOBDET_AMT = JOBDETPRC
                JOBDIS_VAL = JOBDISVAL * JOBTAXRATE
                JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
                JOBTAXVAL = (JOBDETAMT - JOBDISCOUNT_2) * JOBTAXRATE

                SQL_STATEMENT = "update CSMS_EstDETAILS set" & _
                    " estimateno = " & JOBESTIMATENO & "," & _
                    " livil = " & JOBLEVEL & "," & _
                    " LINE_NO = " & JOBLINE_NO & "," & _
                    " detcde = " & JOBDETCDE & "," & _
                    " detdsc = " & JOBDETDSC & "," & _
                    " detunt = " & JOBDETUNT & "," & _
                    " detvol = " & JOBDETVOL & "," & _
                    " detprc = " & JOBDETPRC & "," & _
                    " detamt = " & JOBDETAMT & "," & _
                    " code = " & JOBCODE & "," & _
                    " wcode = " & JOBWCODE & "," & _
                    " taxrate = " & (JOBTAXRATE * 100) & "," & _
                    " discrate = " & (JOBDISCRATE * 100) & "," & _
                    " taxval = " & JOBTAXVAL & "," & _
                    " disval = " & JOBDISVAL & "," & _
                    " pocode = " & JOBPOCODE & "," & _
                    " Rep_Or2 = " & JOBRep_Or2 & "," & _
                    " detail = " & JOBDETAIL & "," & _
                    " det_amt = " & JOBDET_AMT & "," & _
                    " dis_val = " & JOBDIS_VAL & "," & _
                    " discount_2 = " & JOBDISCOUNT_2 & _
                    " where id = " & JOBID
                gconDMIS.Execute SQL_STATEMENT
                
                'NEW LOG AUDIT---------------------------------------------------------------
                    Call NEW_LogAudit("EE", "JOB ESTIMATE", SQL_STATEMENT, labid, "JOBS", "JOB CODE: " & Null2String(JOBDETCDE) & " - DISCOUNT", "", Null2String(JOBID))
                'NEW LOG AUDIT---------------------------------------------------------------

                rsEsti_Det.MoveNext
            Loop

            Call FillJobs
            ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
            SQL_STATEMENT = "update CSMS_EstHd set" & _
                " labor = " & TOTJOBAMT - TOTJOBTAX & "," & _
                " l_amtvalue = " & TOTJOBAMT & "," & _
                " l_disc = " & TOTJOBDISCVAL & "," & _
                " l_disc2 = " & TOTJOBDISC * (VAT_RATE / 100) & "," & _
                " l_taxval = " & TOTJOBTAX & "," & _
                " l_discount = " & TOTJOBDISC & "," & _
                " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                " wl_amt = " & 0 & "," & _
                " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                " where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT---------------------------------------------------------------
                Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "", "EST NO: " & txtEstimateno & " - DISCOUNT", "", "")
            'NEW LOG AUDIT---------------------------------------------------------------
        End If
    ElseIf SSTab1.SelectedItem = 2 Then
        Dim PARTSID                                     As Long
        Dim PARTSESTIMATENO                             As String
        Dim PARTSLEVEL                                  As String
        Dim PARTSLINE_NO                                As String
        Dim PARTSDETCDE                                 As String
        Dim PARTSDETDSC                                 As String
        Dim PARTSDETUNT                                 As String
        Dim PARTSDETVOL                                 As Double
        Dim PARTSDETPRC                                 As Double
        Dim PARTSDETAMT                                 As Double
        Dim PARTSCODE                                   As String
        Dim PARTSWCODE                                  As String
        Dim PARTSTAXRATE                                As Double
        Dim PARTSDISCRATE                               As Double
        Dim PARTSTAXVAL                                 As Double
        Dim PARTSDISVAL                                 As Double
        Dim PARTSPOCODE                                 As String
        Dim PARTSRep_Or2                                As String
        Dim PARTSDETAIL                                 As String
        Dim PARTSDET_AMT                                As Double
        Dim PARTSDIS_VAL                                As Double
        Dim PARTSDISCOUNT_2                             As Double
        Dim PARTSREMARKS                                As String

        Set rsEsti_Det = New ADODB.Recordset
        Set rsEsti_Det = gconDMIS.Execute("select * from CSMS_EstDETAILS where estimateno = " & varESTIMATENO & " and livil = '2' order by LINE_NO asc")
        If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
            rsEsti_Det.MoveFirst
            Do While Not rsEsti_Det.EOF
                PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

                PARTSID = rsEsti_Det!ID
                PARTSESTIMATENO = varESTIMATENO
                PARTSLEVEL = "'2'"
                PARTSLINE_NO = Format(N2Str2Null(rsEsti_Det!LINE_NO), "00")
                PARTSDETCDE = N2Str2Null(rsEsti_Det!DETCDE)
                PARTSDETDSC = N2Str2Null(rsEsti_Det!DETDSC)
                PARTSDETUNT = N2Str2Null(rsEsti_Det!detunt)
                PARTSDETVOL = N2Str2Zero(rsEsti_Det!detvol)
                PARTSDETPRC = N2Str2Zero(rsEsti_Det!DetPrc)
                PARTSDETAMT = N2Str2Zero(rsEsti_Det!DETAMT)
                PARTSCODE = N2Str2Null(rsEsti_Det!Code)
                PARTSWCODE = N2Str2Null(rsEsti_Det!wCode)
                PARTSTAXRATE = (VAT_RATE / 100)
                PARTSDISCRATE = varDISVAL
                PARTSDISVAL = (PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE)
                PARTSPOCODE = N2Str2Null(rsEsti_Det!pocode)
                PARTSRep_Or2 = "NULL"
                PARTSDETAIL = "NULL"
                PARTSDET_AMT = N2Str2Zero(rsEsti_Det!DET_AMT)
                PARTSDIS_VAL = PARTSDISVAL * PARTSTAXRATE
                PARTSDISCOUNT_2 = PARTSDET_AMT * PARTSDISCRATE
                PARTSTAXVAL = (PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE

                SQL_STATEMENT = "update CSMS_EstDETAILS set" & _
                              " estimateno = " & PARTSESTIMATENO & "," & _
                              " livil = " & PARTSLEVEL & "," & _
                              " LINE_NO = " & PARTSLINE_NO & "," & _
                              " detcde = " & PARTSDETCDE & "," & _
                              " detdsc = " & PARTSDETDSC & "," & _
                              " detunt = " & PARTSDETUNT & "," & _
                              " detvol = " & PARTSDETVOL & "," & _
                              " detprc = " & PARTSDETPRC & "," & _
                              " detamt = " & PARTSDETAMT & "," & _
                              " code = " & PARTSCODE & "," & _
                              " wcode = " & PARTSWCODE & "," & _
                              " taxrate = " & PARTSTAXRATE * 100 & "," & _
                              " discrate = " & PARTSDISCRATE * 100 & "," & _
                              " taxval = " & PARTSTAXVAL & "," & _
                              " disval = " & PARTSDISVAL & "," & _
                              " pocode = " & PARTSPOCODE & "," & _
                              " Rep_Or2 = " & PARTSRep_Or2 & "," & _
                              " detail = " & PARTSDETAIL & "," & _
                              " det_amt = " & PARTSDET_AMT & "," & _
                              " dis_val = " & PARTSDIS_VAL & "," & _
                              " discount_2 = " & PARTSDISCOUNT_2 & _
                              " where id = " & PARTSID
                gconDMIS.Execute SQL_STATEMENT
                
                'NEW LOG AUDIT---------------------------------------------------------------
                    Call NEW_LogAudit("EE", "JOB ESTIMATE", SQL_STATEMENT, labid, "PARTS", "PART CODE: " & Null2String(PARTSDETCDE) & " - DISCOUNT", "", Null2String(PARTSID))
                'NEW LOG AUDIT---------------------------------------------------------------

                rsEsti_Det.MoveNext
            Loop
            
            Call FillParts
            ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
            SQL_STATEMENT = "update CSMS_EstHd set" & _
                " parts = " & TOTPARTSAMT - TOTPARTSTAX & "," & _
                " p_amtvalue = " & TOTPARTSAMT & "," & _
                " p_disc = " & TOTPARTSDISCVAL & "," & _
                " p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
                " p_taxval = " & TOTPARTSTAX & "," & _
                " p_discount = " & TOTPARTSDISC & "," & _
                " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC + TOTACCDISC & "," & _
                " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                " wp_amt = " & 0 & "," & _
                " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                " where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT---------------------------------------------------------------
                Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "", "EST NO: " & txtEstimateno & " - DISCOUNT", "", "")
            'NEW LOG AUDIT---------------------------------------------------------------
        End If
    ElseIf SSTab1.SelectedItem = 3 Then
        Dim MATID                                       As Long
        Dim MATESTIMATENO                               As String
        Dim MATLEVEL                                    As String
        Dim MATLINE_NO                                  As String
        Dim MATDETCDE                                   As String
        Dim MATDETDSC                                   As String
        Dim MATDETUNT                                   As String
        Dim MATDETVOL                                   As Double
        Dim MATDETPRC                                   As Double
        Dim MATDETAMT                                   As Double
        Dim MatCode                                     As String
        Dim MATWCODE                                    As String
        Dim MATTAXRATE                                  As Double
        Dim MATDISCRATE                                 As Double
        Dim MATTAXVAL                                   As Double
        Dim MATDISVAL                                   As Double
        Dim MATPOCODE                                   As String
        Dim MATRep_Or2                                  As String
        Dim MATDETAIL                                   As String
        Dim MATDET_AMT                                  As Double
        Dim MATDIS_VAL                                  As Double
        Dim MATDISCOUNT_2                               As Double

        Set rsEsti_Det = New ADODB.Recordset
        Set rsEsti_Det = gconDMIS.Execute("select * from CSMS_EstDETAILS where estimateno = " & varESTIMATENO & " and livil = '3' order by LINE_NO asc")
        If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
            rsEsti_Det.MoveFirst
            Do While Not rsEsti_Det.EOF
                MATDISVAL = 0: MATTAXVAL = 0: MATDETAMT = 0
                MATDIS_VAL = 0: MATDISCOUNT_2 = 0: MATDISCRATE = 0

                MATID = rsEsti_Det!ID
                MATESTIMATENO = varESTIMATENO
                MATLEVEL = "'3'"
                MATLINE_NO = Format(N2Str2Null(rsEsti_Det!LINE_NO), "00")
                MATDETCDE = N2Str2Null(rsEsti_Det!DETCDE)
                MATDETDSC = N2Str2Null(rsEsti_Det!DETDSC)
                MATDETUNT = N2Str2Null(rsEsti_Det!detunt)
                MATDETVOL = N2Str2Zero(rsEsti_Det!detvol)
                MATDETPRC = N2Str2Zero(rsEsti_Det!DetPrc)
                MATDETAMT = N2Str2Zero(rsEsti_Det!DETAMT)
                MatCode = N2Str2Null(rsEsti_Det!Code)
                MATWCODE = N2Str2Null(rsEsti_Det!wCode)
                MATTAXRATE = (VAT_RATE / 100)
                MATDISCRATE = varDISVAL
                MATDISVAL = (MATDETPRC * MATDISCRATE) - ((MATDETPRC * MATDISCRATE) * MATTAXRATE)
                MATPOCODE = N2Str2Null(rsEsti_Det!pocode)
                MATRep_Or2 = "NULL"
                MATDETAIL = "NULL"
                MATDET_AMT = N2Str2Zero(rsEsti_Det!DET_AMT)
                MATDIS_VAL = MATDISVAL * MATTAXRATE
                MATDISCOUNT_2 = MATDET_AMT * MATDISCRATE
                MATTAXVAL = (MATDETAMT - MATDISCOUNT_2) * MATTAXRATE

                SQL_STATEMENT = "update CSMS_EstDETAILS set" & _
                              " estimateno = " & MATESTIMATENO & "," & _
                              " livil = " & MATLEVEL & "," & _
                              " LINE_NO = " & MATLINE_NO & "," & _
                              " detcde = " & MATDETCDE & "," & _
                              " detdsc = " & MATDETDSC & "," & _
                              " detunt = " & MATDETUNT & "," & _
                              " detvol = " & MATDETVOL & "," & _
                              " detprc = " & MATDETPRC & "," & _
                              " detamt = " & MATDETAMT & "," & _
                              " code = " & MatCode & "," & _
                              " wcode = " & MATWCODE & "," & _
                              " taxrate = " & MATTAXRATE * 100 & "," & _
                              " discrate = " & MATDISCRATE * 100 & "," & _
                              " taxval = " & MATTAXVAL & "," & _
                              " disval = " & MATDISVAL & "," & _
                              " pocode = " & MATPOCODE & "," & _
                              " Rep_Or2 = " & MATRep_Or2 & "," & _
                              " detail = " & MATDETAIL & "," & _
                              " det_amt = " & MATDET_AMT & "," & _
                              " dis_val = " & MATDIS_VAL & "," & _
                              " discount_2 = " & MATDISCOUNT_2 & _
                              " where id = " & MATID
                gconDMIS.Execute SQL_STATEMENT
                
                'NEW LOG AUDIT---------------------------------------------------------------
                    Call NEW_LogAudit("EE", "JOB ESTIMATE", SQL_STATEMENT, labid, "MAT", "MAT CODE: " & Null2String(MATDETCDE) & " - DISCOUNT", "", Null2String(MATID))
                'NEW LOG AUDIT---------------------------------------------------------------

                rsEsti_Det.MoveNext
            Loop
            
            Call FillMaterials
            ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
            SQL_STATEMENT = "update CSMS_EstHd set" & _
                " material = " & TOTMATAMT - TOTMATTAX & "," & _
                " m_amtvalue = " & TOTMATAMT & "," & _
                " m_disc = " & TOTMATDISCVAL & "," & _
                " m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
                " m_taxval = " & TOTMATTAX & "," & _
                " m_discount = " & TOTMATDISC & "," & _
                " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC + TOTACCDISC & "," & _
                " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                " wm_amt = " & 0 & "," & _
                " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                " where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT---------------------------------------------------------------
            Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "", "EST NO: " & txtEstimateno & " - DISCOUNT", "", "")
            'NEW LOG AUDIT---------------------------------------------------------------
        End If
    ElseIf SSTab1.SelectedItem = 4 Then                            'FOR ACCESSORIES
        Dim ACCID                                       As Long
        Dim ACCESTIMATENO                               As String
        Dim ACCLEVEL                                    As String
        Dim ACCLINE_NO                                  As String
        Dim ACCDETCDE                                   As String
        Dim ACCDETDSC                                   As String
        Dim ACCDETUNT                                   As String
        Dim ACCDETVOL                                   As Double
        Dim ACCDETPRC                                   As Double
        Dim ACCDETAMT                                   As Double
        Dim ACCCODE                                     As String
        Dim ACCWCODE                                    As String
        Dim ACCTAXRATE                                  As Double
        Dim ACCDISCRATE                                 As Double
        Dim ACCTAXVAL                                   As Double
        Dim ACCDISVAL                                   As Double
        Dim ACCPOCODE                                   As String
        Dim ACCRep_Or2                                  As String
        Dim ACCDETAIL                                   As String
        Dim ACCDET_AMT                                  As Double
        Dim ACCDIS_VAL                                  As Double
        Dim ACCDISCOUNT_2                               As Double
        Dim ACCREMARKS                                  As String

        Set rsEsti_Det = New ADODB.Recordset
        Set rsEsti_Det = gconDMIS.Execute("Select * from CSMS_EstDETAILS where estimateno = " & varESTIMATENO & " and livil = '1' order by LINE_NO asc")
        If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
            rsEsti_Det.MoveFirst
            Do While Not rsEsti_Det.EOF
                ACCDISVAL = 0: ACCTAXVAL = 0: ACCDETAMT = 0
                ACCDIS_VAL = 0: ACCDISCOUNT_2 = 0: ACCDISCRATE = 0

                ACCID = rsEsti_Det!ID
                ACCESTIMATENO = varESTIMATENO
                ACCLEVEL = "'4'"
                ACCLINE_NO = Format(N2Str2Null(rsEsti_Det!LINE_NO), "00")
                ACCDETCDE = N2Str2Null(rsEsti_Det!DETCDE)
                ACCDETDSC = N2Str2Null(rsEsti_Det!DETDSC)
                ACCDETUNT = N2Str2Null(rsEsti_Det!detunt)
                ACCDETVOL = N2Str2IntZero(rsEsti_Det!detvol)
                ACCDETPRC = N2Str2Zero(rsEsti_Det!DetPrc)
                ACCCODE = N2Str2Null(rsEsti_Det!Code)
                ACCWCODE = N2Str2Null(rsEsti_Det!wCode)
                ACCTAXRATE = (VAT_RATE / 100)
                ACCDISCRATE = varDISVAL
                ACCDETAMT = ACCDETPRC / ConvertToBIRDecimalFormat(VAT_RATE)
                ACCDISVAL = (ACCDETPRC * ACCDISCRATE) - ((ACCDETPRC * ACCDISCRATE) * ACCTAXRATE)
                ACCPOCODE = N2Str2Null(rsEsti_Det!pocode)
                ACCRep_Or2 = "NULL"
                ACCDETAIL = N2Str2Null(rsEsti_Det!DETAIL)
                ACCDET_AMT = ACCDETPRC
                ACCDIS_VAL = ACCDISVAL * ACCTAXRATE
                ACCDISCOUNT_2 = ACCDET_AMT * ACCDISCRATE
                ACCTAXVAL = (ACCDETAMT - ACCDISCOUNT_2) * ACCTAXRATE

                SQL_STATEMENT = "update CSMS_EstDETAILS set" & _
                              " estimateno = " & ACCESTIMATENO & "," & _
                              " livil = " & ACCLEVEL & "," & _
                              " LINE_NO = " & ACCLINE_NO & "," & _
                              " detcde = " & ACCDETCDE & "," & _
                              " detdsc = " & ACCDETDSC & "," & _
                              " detunt = " & ACCDETUNT & "," & _
                              " detvol = " & ACCDETVOL & "," & _
                              " detprc = " & ACCDETPRC & "," & _
                              " detamt = " & ACCDETAMT & "," & _
                              " code = " & ACCCODE & "," & _
                              " wcode = " & ACCWCODE & "," & _
                              " taxrate = " & (ACCTAXRATE * 100) & "," & _
                              " discrate = " & (ACCDISCRATE * 100) & "," & _
                              " taxval = " & ACCTAXVAL & "," & _
                              " disval = " & ACCDISVAL & "," & _
                              " pocode = " & ACCPOCODE & "," & _
                              " Rep_Or2 = " & ACCRep_Or2 & "," & _
                              " detail = " & ACCDETAIL & "," & _
                              " det_amt = " & ACCDET_AMT & "," & _
                              " dis_val = " & ACCDIS_VAL & "," & _
                              " discount_2 = " & ACCDISCOUNT_2 & _
                              " where id = " & ACCID
                gconDMIS.Execute SQL_STATEMENT
                
                'NEW LOG AUDIT---------------------------------------------------------------
                    Call NEW_LogAudit("EE", "JOB ESTIMATE", SQL_STATEMENT, labid, "ACC", "ACC CODE: " & Null2String(ACCDETCDE) & " - DISCOUNT", "", Null2String(ACCID))
                'NEW LOG AUDIT---------------------------------------------------------------

                rsEsti_Det.MoveNext
            Loop

            Call FillAccessories
            ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
            SQL_STATEMENT = "update CSMS_EstHd set" & _
                " ACCESSORIES = " & TOTACCAMT - TOTACCTAX & "," & _
                " A_amtvalue = " & TOTACCAMT & "," & _
                " A_disc = " & TOTACCDISCVAL & "," & _
                " A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
                " A_taxval = " & TOTACCTAX & "," & _
                " A_discount = " & TOTACCDISC & "," & _
                " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                " wl_amt = " & 0 & "," & _
                " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                " where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT---------------------------------------------------------------
                Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "", "EST NO: " & txtEstimateno & " - DISCOUNT", "", "")
            'NEW LOG AUDIT---------------------------------------------------------------
        End If
    End If
    
    Call ShowSuccessFullyUpdated
    Screen.MousePointer = 0
    
    Call rsRefresh
    rsEsti_HD.Find "id = " & labid.Caption
    Call StoreMemVars
    Call cmdCancelDisk_Click
End Sub

Sub UpdateParticipation()
    Screen.MousePointer = 11
    SQL_STATEMENT = "Update CSMS_ESTHD Set " & _
        " PartLabor = " & NumericVal(txtPartLabor) & "," & _
        " PartParts = " & NumericVal(txtPartParts) & "," & _
        " PartMaterials = " & NumericVal(txtPartMaterials) & "," & _
        " Partaccessories = " & NumericVal(txtPartAccessories) & "," & _
        " INSAMT = " & NumericVal(txtPartTotal) & _
        " Where ID = " & labid.Caption & ""
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        'Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labID, "", "RO NO: " & txtRep_Or & " - INSURANCE", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call FillJobs
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT - (NumericVal(txtPartTotal))
    SQL_STATEMENT = "update CSMS_ESTHD set" & _
                  " labor = " & Round(TOTJOBAMT - TOTJOBTAX, 2) - (NumericVal(txtPartLabor)) & "," & _
                  " l_amtvalue = " & Round(TOTJOBAMT, 2) - (NumericVal(txtPartLabor)) & "," & _
                  " l_disc = " & Round(TOTJOBDISCVAL, 2) & "," & _
                  " l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
                  " l_taxval = " & Round(TOTJOBTAX, 2) & "," & _
                  " l_discount = " & Round(TOTJOBDISC, 2) & "," & _
                  " amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & "," & _
                  " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX, 2) & "," & _
                  " wl_amt = " & 0 & "," & _
                  " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                  " where ID = " & labid.Caption & ""
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        'Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labID, "", "RO NO: " & txtRep_Or & " - LABOR INSURANCE", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call FillParts
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT - (NumericVal(txtPartTotal))
    SQL_STATEMENT = "update CSMS_ESTHD set" & _
                  " parts = " & TOTPARTSAMT - TOTPARTSTAX - (NumericVal(txtPartParts)) & "," & _
                  " p_amtvalue = " & TOTPARTSAMT - NumericVal(txtPartParts) & "," & _
                  " p_disc = " & TOTPARTSDISCVAL & "," & _
                  " p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
                  " p_taxval = " & TOTPARTSTAX & "," & _
                  " p_discount = " & TOTPARTSDISC & "," & _
                  " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                  " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                  " wp_amt = " & 0 & "," & _
                  " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                  " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        'Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labID, "", "RO NO: " & txtRep_Or & " - PART INSURANCE", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call FillMaterials
    ROTotal = (TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT - (NumericVal(txtPartTotal)))
    SQL_STATEMENT = "update CSMS_ESTHD set" & _
                  " material = " & TOTMATAMT - TOTMATTAX - NumericVal(txtPartMaterials) & "," & _
                  " m_amtvalue = " & TOTMATAMT - NumericVal(txtPartMaterials) & "," & _
                  " m_disc = " & TOTMATDISCVAL & "," & _
                  " m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
                  " m_taxval = " & TOTMATTAX & "," & _
                  " m_discount = " & TOTMATDISC & "," & _
                  " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                  " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                  " wm_amt = " & 0 & "," & _
                  " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                  " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        'Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labID, "", "RO NO: " & txtRep_Or & " - MAT INSURANCE", "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    
    Call FillAccessories
    ROTotal = (TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT - (NumericVal(txtPartTotal)))
    SQL_STATEMENT = "update CSMS_ESTHD set" & _
        " Accessories = " & TOTACCAMT - TOTACCTAX - NumericVal(txtPartAccessories) & "," & _
        " A_amtvalue = " & TOTACCAMT - NumericVal(txtPartAccessories) & "," & _
        " A_disc = " & TOTACCDISCVAL & "," & _
        " A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
        " A_taxval = " & TOTACCTAX & "," & _
        " A_discount = " & TOTACCDISC & "," & _
        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
        " WA_amt = " & 0 & "," & _
        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
        " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        'Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labID, "", "RO NO: " & txtRep_Or & " - ACC INSURANCE", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call rsRefresh
    rsEsti_HD.Find "ESTIMATENO = '" & txtEstimateno.Text & "'"
    Call StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub cmdPartClose_Click()
    PicInsurance.Visible = False
    Call EnableFrame(True)
End Sub

Private Sub cmdPartSave_Click()
    Call UpdateParticipation
    
    Call cmdPartClose_Click
    Call ShowSuccessFullyUpdated
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "JOB ESTIMATE") = False Then Exit Sub

    PrintSQLReport rptEstimate, CSMS_REPORT_PATH & "PrintEstimate.rpt", "{repor.ESTIMATENO} = '" & txtEstimateno & "'", CSMS_REPORT_CONNECTION, 1
End Sub

Private Sub cboDescription_LostFocus()
    cboPartNo.Text = setpartcode(cboDescription.Text)
    cboDescription.Text = SetPartDisc(cboPartNo.Text)
    txtUnitPrice.Text = SetPartPrice(cboPartNo.Text)
    txtPartAmount.Text = NumericVal(txtQTY.Text) * NumericVal(txtUnitPrice.Text)
End Sub
Function setpartcode(XXX As String)

Set rsPartMas = New ADODB.Recordset
rsPartMas.Open "Select stockno, stockdesc from pmis_stockmas where stockdesc = '" & XXX & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly

If Not (rsPartMas.EOF And rsPartMas.BOF) Then
    setpartcode = Null2String(rsPartMas!STOCKNO)
Else
    Set rsDNPP = New ADODB.Recordset
    rsDNPP.Open "Select partnumber, descriptio from pmis_dnpp where descriptio = '" & XXX & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not (rsDNPP.EOF And rsDNPP.BOF) Then
        setpartcode = Null2String(rsDNPP!partnumber)
    End If
End If

End Function
Private Sub cboMaterial_LostFocus()
        If cboMaterial.Text <> "" Then
        cboMatCode.Text = SetMatCode(cboMaterial.Text)
        cboMaterial.Text = SetMatDisc(cboMatCode.Text)
        txtMatUnitPrice.Text = SetMatPrice(cboMatCode.Text)
        txtMatPOCode.Text = SetMatPOCode(cboMatCode.Text)
        txtMatAmount.Text = txtMatUnitPrice.Text
    End If
End Sub

Private Sub cmdCancel_Click()
    SendToBack
    Frame1.Enabled = False
    Frame2.Enabled = True
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "JOB ESTIMATE") = False Then Exit Sub
    
    If CheckEstimateStatus(txtEstimateno) = "UPLOADED" Then
        MsgBox "Estimate already uploaded to Repair order. " & vbCrLf & "kindly refresh your Job Estimate Module to display fresh data.", vbInformation, "Info."
        Exit Sub
    End If
    AddorEdit = "EDIT"

    Frame1.Enabled = True
    Frame2.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Call SendToBack
    frmCSMSEstiSearchCustomer.Show vbModal
End Sub

Private Sub cmdFirst_Click()
    rsEsti_HD.MoveFirst
    Call StoreMemVars
End Sub

Private Sub cmdJobCancel_Click()
    Call SendToBack
    cmdCancel.Value = True
    'Call FillJobs
    'Call FillDetails
    
    Call EnableFrame(True)
End Sub

Private Sub cmdJobDelete_Click()
    If MsgBox("Delete This Job, Are you Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
        
    SQL_STATEMENT = "delete from CSMS_EstDETAILS where id = " & labDetID.Caption
    gconDMIS.Execute SQL_STATEMENT
    
    Call ShowDeletedMsg
    'NEW LOG AUDIT---------------------------------------------------------
        Call NEW_LogAudit("XX", "JOB ESTIMATE", SQL_STATEMENT, labid, "JOBS", "JOB CODE: " & cboJobCode, "", labDetID)
    'NEW LOG AUDIT---------------------------------------------------------

    Dim cnt                                            As Integer
    Dim rsEsti_detDup                                  As New ADODB.Recordset
    Set rsEsti_detDup = gconDMIS.Execute("select id,LINE_NO from CSMS_EstDETAILS where estimateno = " & N2Str2Null(rsEsti_HD!EstimateNo) & " and livil = '1' order by LINE_NO asc")
    If Not rsEsti_detDup.EOF And Not rsEsti_detDup.BOF Then
        cnt = 0
        rsEsti_detDup.MoveFirst
        Do While Not rsEsti_detDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update CSMS_EstDETAILS set LINE_NO = " & N2Str2Null(Format(cnt, "00")) & " where id = " & rsEsti_detDup!ID
            rsEsti_detDup.MoveNext
        Loop
    End If
    
    Call FillJobs
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_Esti_Hd set" & _
                  " labor = " & TOTJOBAMT - TOTJOBTAX & "," & _
                  " l_amtvalue = " & TOTJOBAMT & "," & _
                  " l_disc = " & TOTJOBDISCVAL & "," & _
                  " l_disc2 = " & TOTJOBDISC * (VAT_RATE / 100) & "," & _
                  " l_taxval = " & TOTJOBTAX & "," & _
                  " l_discount = " & TOTJOBDISC & "," & _
                  " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                  " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                  " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                  " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------------------------
        Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "", "EST NO: " & txtEstimateno & " - DELETE JOB", "", "")
    'NEW LOG AUDIT-----------------------------------------------------------------------

    Call rsRefresh
    rsEsti_HD.Find "id = " & labid.Caption
    cmdJobCancel.Value = True

    Exit Sub

ErrorCode:

    ShowVBError
    Exit Sub
End Sub

Private Sub cmdJobSave_Click()
    Screen.MousePointer = 11

    If cboJcode.Text = "" Then
        MsgBox "Cannot find Job Description... Please repeat choosing Job Description", vbInformation, "Invalid Job Description"
        On Error Resume Next
        cboJcode.SetFocus
        Exit Sub
    End If

    Dim JOBESTIMATENO                                   As String
    Dim JOBLEVEL                                        As String
    Dim JOBLINE_NO                                      As String
    Dim JOBDETCDE                                       As String
    Dim JOBDETDSC                                       As String
    Dim JOBDETUNT                                       As String
    Dim JOBDETVOL                                       As Double
    Dim JOBDETPRC                                       As Double
    Dim JOBDETAMT                                       As Double
    Dim JOBCODE                                         As String
    Dim JOBWCODE                                        As String
    Dim JOBTAXRATE                                      As Double
    Dim JOBDISCRATE                                     As Double
    Dim JOBTAXVAL                                       As Double
    Dim JOBDISVAL                                       As Double
    Dim JOBPOCODE                                       As String
    Dim JOBRep_Or2                                      As String
    Dim JOBDETAIL                                       As String
    Dim JOBDET_AMT                                      As Double
    Dim JOBDIS_VAL                                      As Double
    Dim JOBDISCOUNT_2                                   As Double
    Dim JOBREMARKS                                      As String
    Dim JOBHRS                                          As Double

    JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
    JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0
    JOBESTIMATENO = N2Str2Null(txtEstimateno.Text)
    JOBLEVEL = "'1'"
    JOBLINE_NO = N2Str2Null(Format(txtJobLineNo.Text, "00"))
    JOBDETCDE = N2Str2Null(cboJcode.Text)
    JOBDETDSC = N2Str2Null(Mid(cboJobCode.Text, 1, 250))
    JOBDETUNT = "NULL"
    JOBDETVOL = NumericVal(labDetID.Caption)
    JOBDETPRC = NumericVal(txtJobRate.Text)
    JOBCODE = "NULL"
    JOBWCODE = N2Str2Null(cboJobChargeTo.Text)
    JOBTAXRATE = (VAT_RATE / 100)
    JOBDISCRATE = NumericVal(txtJobDiscount.Text) / 100
    JOBDETAMT = JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE)
    JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
    JOBPOCODE = N2Str2Null(txtJobPostCode.Text)
    JOBRep_Or2 = "NULL"
    JOBDETAIL = N2Str2Null(CheckChar(txtJobDetail.Text))
    JOBDET_AMT = JOBDETPRC
    JOBDIS_VAL = JOBDISVAL * JOBTAXRATE
    JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
    JOBREMARKS = N2Str2Null(CheckChar(txtJobDetail.Text))
    JOBTAXVAL = (JOBDET_AMT - JOBDISCOUNT_2) * JOBTAXRATE
    JOBHRS = NumericVal(txtHRS)
    If NumericVal(txtPartDiscount.Text) > 0 Then
        If NumericVal(txtJobDiscount.Text) > 0 Then
            If JOBDISCOUNT_2 > TOTJOBAMT Then
                MsgBox "Invalid Discount.", vbInformation + vbOKOnly
                On Error Resume Next
                txtJobDiscount.SetFocus
                Exit Sub
            End If
        End If
    End If
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into CSMS_EstDETAILS " & _
            "(det_hrs,TRANSTYPE, estimateno, livil, LINE_NO,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,Rep_Or2,detail,det_amt,dis_val,discount_2, REP_OR)" & _
            " values (" & JOBHRS & ", 'E' " & _
            "," & JOBESTIMATENO & ", " & JOBLEVEL & _
            ", " & JOBLINE_NO & ", " & JOBDETCDE & _
            ", " & JOBDETDSC & ", " & JOBDETUNT & _
            ", " & JOBDETVOL & ", " & JOBDETPRC & _
            ", " & JOBDETAMT & ", " & JOBCODE & _
            ", " & JOBWCODE & ", " & (JOBTAXRATE * 100) & ", " & (JOBDISCRATE * 100) & _
            ", " & JOBTAXVAL & ", " & JOBDISVAL & _
            ", " & JOBPOCODE & ", " & JOBRep_Or2 & _
            ", " & JOBDETAIL & ", " & JOBDET_AMT & _
            ", " & JOBDIS_VAL & ", " & JOBDISCOUNT_2 & _
            ", " & JOBESTIMATENO & " )"
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT----------------------------------------------------------------------
            Call NEW_LogAudit("AA", "JOB ESTIMATE", SQL_STATEMENT, labid, "JOBS", "JOB CODE: " & cboJobCode, "", "")
        'NEW LOG AUDIT----------------------------------------------------------------------

        Call ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CSMS_EstDETAILS set" & _
            " det_hrs = " & JOBHRS & _
            ", REP_OR = " & JOBESTIMATENO & _
            ", estimateno = " & JOBESTIMATENO & "," & _
            " livil = " & JOBLEVEL & "," & _
            " LINE_NO = " & JOBLINE_NO & "," & _
            " detcde = " & JOBDETCDE & "," & _
            " detdsc = " & JOBDETDSC & "," & _
            " detunt = " & JOBDETUNT & "," & _
            " detvol = " & JOBDETVOL & "," & _
            " detprc = " & JOBDETPRC & "," & _
            " detamt = " & JOBDETAMT & "," & _
            " code = " & JOBCODE & "," & _
            " wcode = " & JOBWCODE & "," & _
            " taxrate = " & (JOBTAXRATE * 100) & "," & _
            " discrate = " & (JOBDISCRATE * 100) & "," & _
            " taxval = " & JOBTAXVAL & "," & _
            " disval = " & JOBDISVAL & "," & _
            " pocode = " & JOBPOCODE & "," & _
            " Rep_Or2 = " & JOBRep_Or2 & "," & _
            " detail = " & JOBDETAIL & "," & _
            " det_amt = " & JOBDET_AMT & "," & _
            " dis_val = " & JOBDIS_VAL & "," & _
            " discount_2 = " & JOBDISCOUNT_2 & _
            " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT----------------------------------------------------------------------
            Call NEW_LogAudit("EE", "JOB ESTIMATE", SQL_STATEMENT, labid, "JOBS", "JOB CODE: " & cboJobCode, "", "")
        'NEW LOG AUDIT----------------------------------------------------------------------

        Call ShowSuccessFullyUpdated
    End If
    
    Call FillJobs
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT

    SQL_STATEMENT = "update CSMS_EstHd set" & _
        " labor = " & TOTJOBAMT - TOTJOBTAX & "," & _
        " l_amtvalue = " & TOTJOBAMT & "," & _
        " l_disc = " & TOTJOBDISCVAL & "," & _
        " l_disc2 = " & TOTJOBDISC * (VAT_RATE / 100) & "," & _
        " l_taxval = " & TOTJOBTAX & "," & _
        " l_discount = " & TOTJOBDISC & "," & _
        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
        " wl_amt = " & 0 & "," & _
        " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
        " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT----------------------------------------------------------
        Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "", "EST NO: " & txtEstimateno, "", "")
    'NEW LOG AUDIT----------------------------------------------------------

    Call rsRefresh
    rsEsti_HD.Find "id = " & labid.Caption
    cmdJobCancel.Value = True
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
    Exit Sub
End Sub

Private Sub cmdLast_Click()
    rsEsti_HD.MoveLast
    Call StoreMemVars
End Sub

Private Sub cmdMatCancel_Click()
    Call SendToBack
    cmdCancel.Value = True
    
    Call EnableFrame(True)
End Sub

Private Sub cmdMatDelete_Click()
    If MsgBox("Delete This Materials, Are you Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    SQL_STATEMENT = "delete from CSMS_EstDETAILS where id = " & labDetID.Caption
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT----------------------------------------------------------
        Call NEW_LogAudit("XX", "JOB ESTIMATE", SQL_STATEMENT, labid, "MAT", "MAT CODE: " & cboMatCode, "", labDetID)
    'NEW LOG AUDIT----------------------------------------------------------
    
    Call ShowDeletedMsg

    Dim cnt                                            As Integer
    Dim rsEsti_detDup                                  As New ADODB.Recordset
    Set rsEsti_detDup = gconDMIS.Execute("select id,LINE_NO from CSMS_EstDETAILS where estimateno = " & N2Str2Null(rsEsti_HD!EstimateNo) & " and livil = '3' order by LINE_NO asc")
    If Not rsEsti_detDup.EOF And Not rsEsti_detDup.BOF Then
        cnt = 0
        rsEsti_detDup.MoveFirst
        Do While Not rsEsti_detDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update CSMS_EstDETAILS set LINE_NO = " & N2Str2Null(Format(cnt, "00")) & " where id = " & rsEsti_detDup!ID
            rsEsti_detDup.MoveNext
        Loop
    End If

    Call FillMaterials
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_Esti_Hd set" & _
                  " material = " & TOTMATAMT - TOTMATTAX & "," & _
                  " m_amtvalue = " & TOTMATAMT & "," & _
                  " m_disc = " & TOTMATDISCVAL & "," & _
                  " m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
                  " m_taxval = " & TOTMATTAX & "," & _
                  " m_discount = " & TOTMATDISC & "," & _
                  " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                  " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                  " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                  " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------------------------
        Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "", "EST NO: " & txtEstimateno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------------------------

    Call rsRefresh
    rsEsti_HD.Find "id = " & labid.Caption
    cmdMatCancel.Value = True
End Sub

Private Sub cmdMatSave_Click()
    Screen.MousePointer = 11

    Dim MATESTIMATENO                                   As String
    Dim MATLINE_NO                                      As String
    Dim MATDETCDE                                       As String
    Dim MATDETDSC                                       As String
    Dim MATDETUNT                                       As String
    Dim MATDETVOL                                       As Double
    Dim MATDETPRC                                       As Double
    Dim MATDETAMT                                       As Double
    Dim MatCode                                         As String
    Dim MATWCODE                                        As String
    Dim MATLEVEL                                        As String
    Dim MATTAXRATE                                      As Double
    Dim MATDISCRATE                                     As Double
    Dim MATTAXVAL                                       As Double
    Dim MATDISVAL                                       As Double
    Dim MATPOCODE                                       As String
    Dim MATRep_Or2                                      As String
    Dim MATDETAIL                                       As String
    Dim MATDET_AMT                                      As Double
    Dim MATDIS_VAL                                      As Double
    Dim MATDISCOUNT_2                                   As Double



    
    If RTrim(LTrim(cboMatCode.Text)) = "" Then
        MsgBox "Material code cannot be blank.", vbInformation + vbOKOnly
        cboMatCode.SetFocus
        Exit Sub
    End If

    MATDISVAL = 0: MATTAXVAL = 0: MATDETAMT = 0
    MATDIS_VAL = 0: MATDISCOUNT_2 = 0: MATDISCRATE = 0

    MATESTIMATENO = N2Str2Null(txtEstimateno.Text)
    MATLEVEL = "'3'"
    MATLINE_NO = N2Str2Null(Format(txtMatLineNo.Text, "00"))
    MATDETCDE = N2Str2Null(cboMatCode.Text)
    MATDETDSC = N2Str2Null(cboMaterial.Text)
    MATDETUNT = "NULL"
    MATDETVOL = NumericVal(txtMatQty.Text)
    MATDETPRC = NumericVal(txtMatUnitPrice.Text)
    MATDETAMT = NumericVal(txtMatAmount.Text) / ConvertToBIRDecimalFormat(VAT_RATE)
    MatCode = "NULL"
    MATWCODE = N2Str2Null(cboMatChargeTo.Text)
    MATTAXRATE = (VAT_RATE / 100)
    MATDISCRATE = NumericVal(txtMatDiscount.Text) / 100
    MATDISVAL = (MATDETPRC * MATDISCRATE) - ((MATDETPRC * MATDISCRATE) * MATTAXRATE)
    MATPOCODE = N2Str2Null(txtMatPOCode.Text)
    MATRep_Or2 = "NULL"
    MATDETAIL = "NULL"
    MATDET_AMT = NumericVal(txtMatAmount.Text)
    MATDIS_VAL = MATDISVAL * MATTAXRATE
    MATDISCOUNT_2 = MATDET_AMT * MATDISCRATE
    MATTAXVAL = (MATDET_AMT - MATDISCOUNT_2) * MATTAXRATE
    If MATDISCOUNT_2 > TOTMATAMT Then
        MsgBox "Invalid Discount.", vbInformation + vbOKOnly
        On Error Resume Next
        txtMatDiscount.SetFocus
        Exit Sub
    End If

    If gconDMIS.Execute("Select count(*) from pmis_stockmas where [type] = 'M' and stockno = " & MATDETCDE & "").Fields(0).Value = 0 Then
        MsgBox "Material code does't exist.", vbInformation + vbOKOnly
        cboMatCode.SetFocus
        Exit Sub
    End If
    
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into CSMS_EstDETAILS " & _
            "(TRANSTYPE, estimateno, livil, LINE_NO, detcde, detdsc, detunt, detvol, detprc, detamt, code, wcode, taxrate, discrate, taxval, disval, pocode, Rep_Or2, detail, det_amt, dis_val, discount_2, REP_OR)" & _
            " values ('E',  " & MATESTIMATENO & _
            ", " & MATLEVEL & ", " & MATLINE_NO & "," & _
            " " & MATDETCDE & "," & MATDETDSC & "," & _
            " " & MATDETUNT & ", " & MATDETVOL & "," & _
            " " & MATDETPRC & ", " & MATDETAMT & ", " & MatCode & _
            ", " & MATWCODE & ", " & MATTAXRATE * 100 & ", " & MATDISCRATE * 100 & _
            ", " & MATTAXVAL & ", " & MATDISVAL & ", " & MATPOCODE & _
            ", " & MATRep_Or2 & ", " & MATDETAIL & ", " & MATDET_AMT & _
            ", " & MATDIS_VAL & ", " & MATDISCOUNT_2 & _
            ", " & MATESTIMATENO & ")"
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT---------------------------------------------------------
            Call NEW_LogAudit("AA", "JOB ESTIMATE", SQL_STATEMENT, labid, "MAT", "MAT CODE: " & cboMatCode, "", "")
        'NEW LOG AUDIT---------------------------------------------------------

        Call ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CSMS_EstDETAILS set" & _
            " REP_OR = " & MATESTIMATENO & _
            ", estimateno = " & MATESTIMATENO & "," & _
            " livil = " & MATLEVEL & "," & _
            " LINE_NO = " & MATLINE_NO & "," & _
            " detcde = " & MATDETCDE & "," & _
            " detdsc = " & MATDETDSC & "," & _
            " detunt = " & MATDETUNT & "," & _
            " detvol = " & MATDETVOL & "," & _
            " detprc = " & MATDETPRC & "," & _
            " detamt = " & MATDETAMT & "," & _
            " code = " & MatCode & "," & _
            " wcode = " & MATWCODE & "," & _
            " taxrate = " & MATTAXRATE * 100 & "," & _
            " discrate = " & MATDISCRATE * 100 & "," & _
            " taxval = " & MATTAXVAL & "," & _
            " disval = " & MATDISVAL & "," & _
            " pocode = " & MATPOCODE & "," & _
            " Rep_Or2 = " & MATRep_Or2 & "," & _
            " detail = " & MATDETAIL & "," & _
            " det_amt = " & MATDET_AMT & "," & _
            " dis_val = " & MATDIS_VAL & "," & _
            " discount_2 = " & MATDISCOUNT_2 & _
            " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT---------------------------------------------------------
            Call NEW_LogAudit("EE", "JOB ESTIMATE", SQL_STATEMENT, labid, "MAT", "MAT CODE: " & cboMatCode, "", "")
        'NEW LOG AUDIT---------------------------------------------------------

        Call ShowSuccessFullyUpdated
    End If

    Call FillMaterials
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_EstHd set" & _
        " material = " & TOTMATAMT - TOTMATTAX & "," & _
        " m_amtvalue = " & TOTMATAMT & "," & _
        " m_disc = " & TOTMATDISCVAL & "," & _
        " m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
        " m_taxval = " & TOTMATTAX & "," & _
        " m_discount = " & TOTMATDISC & "," & _
        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
        " wm_amt = " & 0 & "," & _
        " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
        " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT---------------------------------------------------------
        Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "MAT", "EST NO: " & txtEstimateno, "", "")
    'NEW LOG AUDIT---------------------------------------------------------

    Call rsRefresh
    rsEsti_HD.Find "id = " & labid.Caption
    Call StoreMemVars
    cmdMatCancel.Value = True
    Screen.MousePointer = 0
    
    If AddorEdit = "ADD" Then AddMaterials
    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
    Exit Sub
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsEsti_HD.MoveNext
    If rsEsti_HD.EOF Then
        rsEsti_HD.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPartsCancel_Click()
    Call SendToBack
    cmdCancel.Value = True
'    Call FillParts
'    Call FillDetails
    
    Call EnableFrame(True)
End Sub

Private Sub cmdPartsDelete_Click()
    
    If MsgBox("Delete This Parts, Are you Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
        
    SQL_STATEMENT = "delete from CSMS_EstDETAILS where id = " & labDetID.Caption
    gconDMIS.Execute SQL_STATEMENT

    'NEW LOG AUDIT----------------------------------------------------
        Call NEW_LogAudit("XX", "JOB ESTIMATE", SQL_STATEMENT, labid, "PARTS", "PART NO: " & cboPartNo, "", labDetID)
    'NEW LOG AUDIT----------------------------------------------------

    Call ShowDeletedMsg

    Dim cnt                                            As Integer
    Dim rsEsti_detDup                                  As New ADODB.Recordset
    Set rsEsti_detDup = gconDMIS.Execute("select id,LINE_NO from CSMS_EstDETAILS where estimateno = " & N2Str2Null(rsEsti_HD!EstimateNo) & " and livil = '2' order by LINE_NO asc")
    If Not rsEsti_detDup.EOF And Not rsEsti_detDup.BOF Then
        cnt = 0
        rsEsti_detDup.MoveFirst
        Do While Not rsEsti_detDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update CSMS_EstDETAILS set LINE_NO = " & N2Str2Null(Format(cnt, "00")) & " where id = " & rsEsti_detDup!ID
            rsEsti_detDup.MoveNext
        Loop
    End If

    Call FillParts
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_Esti_Hd set" & _
        " parts = " & TOTPARTSAMT - TOTPARTSTAX & "," & _
        " p_amtvalue = " & TOTPARTSAMT & "," & _
        " p_disc = " & TOTPARTSDISCVAL & "," & _
        " p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
        " p_taxval = " & TOTPARTSTAX & "," & _
        " p_discount = " & TOTPARTSDISC & "," & _
        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
        " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
        " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT----------------------------------------------
    Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "", "EST NO: " & txtEstimateno & " - PARTS REMOVE", "", "")
    'NEW LOG AUDIT----------------------------------------------

    Call rsRefresh
    rsEsti_HD.Find "id = " & labid.Caption
    cmdPartsCancel.Value = True
End Sub

Private Sub cmdPartsSave_Click()
    Screen.MousePointer = 11
    On Error GoTo ErrorCode

    Dim PARTSESTIMATENO                                 As String
    Dim PARTSLEVEL                                      As String
    Dim PARTSLINE_NO                                    As String
    Dim PARTSDETCDE                                     As String
    Dim PARTSDETDSC                                     As String
    Dim PARTSDETUNT                                     As String
    Dim PARTSDETVOL                                     As Double
    Dim PARTSDETPRC                                     As Double
    Dim PARTSDETAMT                                     As Double
    Dim PARTSCODE                                       As String
    Dim PARTSWCODE                                      As String
    Dim PARTSTAXRATE                                    As Double
    Dim PARTSDISCRATE                                   As Double
    Dim PARTSTAXVAL                                     As Double
    Dim PARTSDISVAL                                     As Double
    Dim PARTSPOCODE                                     As String
    Dim PARTSRep_Or2                                    As String
    Dim PARTSDETAIL                                     As String
    Dim PARTSDET_AMT                                    As Double
    Dim PARTSDIS_VAL                                    As Double
    Dim PARTSDISCOUNT_2                                 As Double
    Dim PARTSREMARKS                                    As String



    If RTrim(LTrim(cboPartNo.Text)) = "" Then
        MsgBox "Part number cannot be blank.", vbInformation + vbOKOnly
        cboPartNo.SetFocus
        Exit Sub
    End If


    PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
    PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

    PARTSESTIMATENO = N2Str2Null(txtEstimateno.Text)
    PARTSLEVEL = "'2'"
    PARTSLINE_NO = N2Str2Null(Format(txtPartsLineNo.Text, "00"))
    PARTSDETCDE = N2Str2Null(cboPartNo.Text)
    PARTSDETDSC = N2Str2Null(cboDescription.Text)
    PARTSDETUNT = "NULL"
    PARTSDETVOL = NumericVal(txtQTY.Text)
    PARTSDETPRC = NumericVal(txtUnitPrice.Text)
    PARTSDETAMT = NumericVal(txtPartAmount.Text) / ConvertToBIRDecimalFormat(VAT_RATE)
    PARTSCODE = "NULL"
    PARTSWCODE = N2Str2Null(cboChargeTo.Text)
    PARTSTAXRATE = (VAT_RATE / 100)
    PARTSDISCRATE = NumericVal(txtPartDiscount.Text) / 100
    PARTSDISVAL = (PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE)
    PARTSPOCODE = N2Str2Null(txtPartCode.Text)
    PARTSRep_Or2 = "NULL"
    PARTSDETAIL = "NULL"
    PARTSDET_AMT = NumericVal(txtPartAmount.Text)
    PARTSDIS_VAL = PARTSDISVAL * PARTSTAXRATE
    PARTSDISCOUNT_2 = PARTSDET_AMT * PARTSDISCRATE
    'PARTSTAXVAL = (PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE
    PARTSTAXVAL = (PARTSDET_AMT - PARTSDISCOUNT_2) * PARTSTAXRATE
    If PARTSDISCOUNT_2 > TOTPARTSAMT Then
        MsgBox "Invalid Discount.", vbInformation + vbOKOnly
        On Error Resume Next
        txtPartDiscount.SetFocus
        Exit Sub
    End If
    
    If gconDMIS.Execute("Select count(*) from pmis_stockmas where [type] = 'P' and stockno = " & PARTSDETCDE & "").Fields(0).Value = 0 Then
        If gconDMIS.Execute("Select count(*) from pmis_dnpp where partnumber = " & PARTSDETCDE & "").Fields(0).Value = 0 Then
            MsgBox "Part number does't exist.", vbInformation + vbOKOnly
            cboPartNo.SetFocus
            Exit Sub
        End If
    End If
    
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into CSMS_EstDETAILS " & _
            "(TRANSTYPE,estimateno,livil,LINE_NO,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,Rep_Or2,detail,det_amt,dis_val,discount_2, REP_OR)" & _
            " values ('E'," & PARTSESTIMATENO & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
            " " & PARTSDETCDE & "," & PARTSDETDSC & "," & _
            " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
            " " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
            ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
            ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
            ", " & PARTSRep_Or2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
            ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & _
            ", " & PARTSESTIMATENO & ")"
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT---------------------------------------------------------------
            Call NEW_LogAudit("AA", "JOB ESTIMATE", SQL_STATEMENT, labid, "MAT", "PART CODE: " & cboPartNo, "", "")
        'NEW LOG AUDIT---------------------------------------------------------------

        Call ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CSMS_EstDETAILS set" & _
            " REP_OR = " & PARTSESTIMATENO & _
            ", estimateno = " & PARTSESTIMATENO & "," & _
            " livil = " & PARTSLEVEL & "," & _
            " LINE_NO = " & PARTSLINE_NO & "," & _
            " detcde = " & PARTSDETCDE & "," & _
            " detdsc = " & PARTSDETDSC & "," & _
            " detunt = " & PARTSDETUNT & "," & _
            " detvol = " & PARTSDETVOL & "," & _
            " detprc = " & PARTSDETPRC & "," & _
            " detamt = " & PARTSDETAMT & "," & _
            " code = " & PARTSCODE & "," & _
            " wcode = " & PARTSWCODE & "," & _
            " taxrate = " & PARTSTAXRATE * 100 & "," & _
            " discrate = " & PARTSDISCRATE * 100 & "," & _
            " taxval = " & PARTSTAXVAL & "," & _
            " disval = " & PARTSDISVAL & "," & _
            " pocode = " & PARTSPOCODE & "," & _
            " Rep_Or2 = " & PARTSRep_Or2 & "," & _
            " detail = " & PARTSDETAIL & "," & _
            " det_amt = " & PARTSDET_AMT & "," & _
            " dis_val = " & PARTSDIS_VAL & "," & _
            " discount_2 = " & PARTSDISCOUNT_2 & _
            " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT---------------------------------------------------------------
        
        Call NEW_LogAudit("EE", "JOB ESTIMATE", SQL_STATEMENT, labid, "MAT", "PART CODE: " & cboPartNo, "", "")
            'NEW LOG AUDIT---------------------------------------------------------------

        Call ShowSuccessFullyUpdated
    End If

    Call FillParts
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_EstHd set" & _
        " parts = " & TOTPARTSAMT - TOTPARTSTAX & "," & _
        " p_amtvalue = " & TOTPARTSAMT & "," & _
        " p_disc = " & TOTPARTSDISCVAL & "," & _
        " p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
        " p_taxval = " & TOTPARTSTAX & "," & _
        " p_discount = " & TOTPARTSDISC & "," & _
        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
        " wp_amt = " & 0 & "," & _
        " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
        " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT----------------------------------------------------------
        Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, labid, "", "EST NO: " & txtEstimateno, "", "")
    'NEW LOG AUDIT----------------------------------------------------------

    Call rsRefresh
    rsEsti_HD.Find "id = " & labid.Caption
    Call StoreMemVars
    cmdPartsCancel.Value = True
    Screen.MousePointer = 0
    
    If AddorEdit = "ADD" Then AddParts
    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
    Exit Sub
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsEsti_HD.MovePrevious
    If rsEsti_HD.BOF Then
        rsEsti_HD.MoveFirst
        Call ShowFirstRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdSave_Click()
    If txtNiym.Text = "" Then
        MsgBoxXP "CUSMAS must have a name", "Invalid CUSMAS Name", XP_OKOnly, msg_Critical
        On Error Resume Next
        txtNiym.SetFocus
        Exit Sub
    End If
    If cboRecd_by.Text = "" Then
        MsgBoxXP "Service Advisor must not be Empty!", "Invalid SA", XP_OKOnly, msg_Information
        On Error Resume Next
        cboRecd_by.SetFocus
        Exit Sub
    Else
        Set rsEmpNo = New ADODB.Recordset
        Set rsEmpNo = gconDMIS.Execute("select naym from CSMS_vw_EmpNo where naym = '" & cboRecd_by.Text & "'")
        If rsEmpNo.EOF And rsEmpNo.BOF Then
            MsgSpeech "Invalid Service Advisor"
            On Error Resume Next
            cboRecd_by.SetFocus
            Exit Sub
        End If
    End If

    If AddorEdit = "ADD" Then
        Dim rsESTI_HDDup                               As New ADODB.Recordset
        Set rsESTI_HDDup = gconDMIS.Execute("select id from CSMS_Esti_Hd order by id asc")
        If Not rsESTI_HDDup.EOF And Not rsESTI_HDDup.BOF Then
            rsESTI_HDDup.MoveLast
            labid.Caption = NumericVal(rsESTI_HDDup!ID) + 1
        End If
    End If

    Dim VTXTestimateno                                  As String
    Dim VTXTROType                                      As String
    Dim VTXTSvc_No                                      As String
    Dim VTXTAcct_No                                     As String
    Dim VTXTNiym                                        As String
    Dim VTXTPlate_No                                    As String
    Dim VcboModel                                       As String
    Dim VTXTMake                                        As String
    Dim VTXTTerm                                        As String
    Dim VTXTSektion                                     As String
    Dim VTXTKm_rdg                                      As String
    Dim VTXTDte_recd                                    As String
    Dim VTXTCertific8                                   As String
    Dim VTXTDte_comp                                    As String
    Dim VTXTDte_Rel                                     As String
    Dim VtxtAddress                                     As String
    Dim VTXTPart_amt                                    As Double
    Dim VTXTParticipat                                  As String
    Dim vtxtInsuranceName                               As String
    Dim VcboRecd_by                                     As String
    Dim kAdd                                            As Integer
    
    VTXTestimateno = N2Str2Null(txtEstimateno.Text)
    VTXTROType = N2Str2Null(txtROType.Text)
    VTXTSvc_No = N2Str2Null(txtSvc_No.Text)
    VTXTAcct_No = N2Str2Null(txtAcct_No.Text)
    VTXTNiym = N2Str2Null(txtNiym.Text)
    
    For kAdd = 1 To Len(txtAddress.Text)
        If Mid(txtAddress.Text, kAdd, 1) = "-" And Mid(txtAddress.Text, kAdd + 1, 1) = "-" And Mid(txtAddress.Text, kAdd + 1, 1) = "-" Then Exit For
        VtxtAddress = VtxtAddress & Mid(txtAddress.Text, kAdd, 1)
    Next
    VtxtAddress = N2Str2Null(VtxtAddress)
    VTXTPlate_No = N2Str2Null(txtPlate_No.Text)
    VcboModel = N2Str2Null(cboModel.Text)
    VTXTMake = N2Str2Null(txtMake.Text)
    VTXTTerm = N2Str2Null(txtTerm.Text)
    VTXTSektion = N2Str2Null(txtSektion.Text)
    VTXTKm_rdg = N2Str2Null(txtKm_rdg.Text)
    VTXTDte_recd = N2Date2Null(txtDte_recd.Text)
    VTXTCertific8 = N2Str2Null(txtCertific8.Text)
    VTXTDte_comp = N2Date2Null(txtDte_comp.Text)
    VTXTDte_Rel = N2Date2Null(txtDte_Rel.Text)
    VTXTPart_amt = NumericVal(txtPart_amt.Text)
    If chkParticipat.Value = 1 Then
        VTXTParticipat = N2Str2Null(txtParticipat.Text)
        vtxtInsuranceName = N2Str2Null(Text2)
    Else
        VTXTParticipat = N2Str2Null("")
        vtxtInsuranceName = N2Str2Null("")
    End If
    VcboRecd_by = N2Str2Null(SetCodeSA(cboRecd_by.Text))

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into CSMS_ESTHD " & _
            "(TRANSTYPE, estimateno,rotype,svc_no,acct_no,niym,plate_no,model,term,sektion,Recd_by,km_rdg,dte_recd,certific8,dte_comp,dte_rel,part_amt, participat, status, INSCDE)" & _
            " values ('E'," & VTXTestimateno & _
            ", " & VTXTROType & ", " & VTXTSvc_No & _
            ", " & VTXTAcct_No & ", " & VTXTNiym & _
            ", " & VTXTPlate_No & ", " & VcboModel & _
            ", " & VTXTTerm & ", " & VTXTSektion & _
            ", " & VcboRecd_by & ", " & VTXTKm_rdg & _
            ", " & VTXTDte_recd & ", " & VTXTCertific8 & _
            ", " & VTXTDte_comp & ", " & VTXTDte_Rel & _
            ", " & VTXTPart_amt & ", " & VTXTParticipat & ", 'N', " & vtxtInsuranceName & ")"
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------------------
            Call NEW_LogAudit("A", "JOB ESTIMATE", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtEstimateno), "ESTIMATENO", "CSMS_Esti_Hd"), "", "EST NO: " & txtEstimateno, "", "")
        'NEW LOG AUDIT-----------------------------------------------------------------

        Call ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CSMS_ESTHD set" & _
            " SERVICE_ADVISER = " & N2Str2Null(cboRecd_by) & ", " & _
            " estimateno = " & VTXTestimateno & "," & _
            " rotype = " & VTXTROType & "," & _
            " svc_no = " & VTXTSvc_No & "," & _
            " acct_no = " & VTXTAcct_No & "," & _
            " niym = " & VTXTNiym & "," & _
            " plate_no = " & VTXTPlate_No & "," & _
            " model = " & VcboModel & "," & _
            " term = " & VTXTTerm & "," & _
            " sektion = " & VTXTSektion & "," & _
            " recd_by = " & VcboRecd_by & "," & _
            " km_rdg = " & VTXTKm_rdg & "," & _
            " dte_recd = " & VTXTDte_recd & "," & _
            " certific8 = " & VTXTCertific8 & "," & _
            " dte_comp = " & VTXTDte_comp & "," & _
            " dte_rel = " & VTXTDte_Rel & "," & _
            " part_amt = " & VTXTPart_amt & "," & _
            " participat = " & VTXTParticipat & "," & _
            " INSCDE = " & N2Str2Null(Text2) & "," & _
            " status = 'N'" & _
            " where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------------------
            Call NEW_LogAudit("A", "JOB ESTIMATE", SQL_STATEMENT, labid, "", "EST NO: " & txtEstimateno, "", "")
        'NEW LOG AUDIT-----------------------------------------------------------------

        Call ShowSuccessFullyUpdated
    End If

    Call rsRefresh
    rsEsti_HD.Find "id = " & labid.Caption
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdUpload_Click()
    If Function_Access(LOGID, "acess_post", "JOB ESTIMATE") = False Then Exit Sub
    
    If txtROno.Text <> "" Then
        MsgBox "Estimate is already Uploaded", vbInformation, "Info"
        Exit Sub
    ElseIf CheckEstimateStatus(txtEstimateno) = "UPLOADED" Then
        MsgBox "Estimate already uploaded to Repair order. " & vbCrLf & " kindly refresh your Job Estimate Module to display fresh data.", vbInformation, "Info."
        Exit Sub
    End If
        
    Call frmApp.FillEstimateno(txtEstimateno, "ESTIMATE")
    frmApp.Show
End Sub

Private Sub Command3_Click()
    Call FRMx.PassVariable("ESTIMATE MODULE")
    FRMx.Show 1
End Sub

Private Sub Command4_Click()
    If chkParticipat.Value = 0 Then Exit Sub
    
    Call FRMx.PassVariable("ESTIMATE INSURANCE")
    FRMx.Show 1
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
        
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If Picture1.Visible = True Then
                Unload frmALL_AuditInquiry

                frmALL_AuditInquiry.Show
                frmALL_AuditInquiry.ZOrder 0
                frmALL_AuditInquiry.Caption = "Audit Inquiry (JOB ESTIMATE)"
                Call frmALL_AuditInquiry.DisplayHistory(labid, "JOB ESTIMATE")
            End If

            If Me.ActiveControl.Name = "txtEstimateNo" Then
                Dim rsESTI_HDDup                       As New ADODB.Recordset
                Set rsESTI_HDDup = gconDMIS.Execute("select estimateno from Ehhehesti_hd where estimateno = " & N2Str2Null(txtEstimateno.Text))
                If rsESTI_HDDup.EOF And rsESTI_HDDup.BOF Then
                    SendKeys "{TAB}"
                End If
            Else
                If Me.ActiveControl.Name = "txtCertific8" Then
                    If txtPlate_No.Text <> "" Then
                        Me.Enabled = False
                        frmCSMSESTICusveh.Show
                        frmCSMSESTICusveh.ZOrder 0
                    Else
                        MsgBoxXP "Plate Number must be inputed!" & vbCrLf & _
                                 "Please enter 000000 if unknown", "No Plate No.!", XP_OKOnly, msg_Critical
                        On Error Resume Next
                        txtPlate_No.SetFocus
                    End If
                ElseIf Me.ActiveControl.Name = "txtAcct_No" Then
                    If txtAcct_No.Text = "" Then
                        SendToBack
                    Else
                        SendKeys "{TAB}"
                    End If
                ElseIf Me.ActiveControl.Name = "txtParticipat" Then
                    If chkParticipat.Value = 1 And txtParticipat.Text = "" Then
                        SendToBack
                        DoEvents
                        RO_OR_ESTI_OR_PART = "PART"
                    Else
                        SendKeys "{TAB}"
                    End If
                Else
                    If Mid(Me.ActiveControl.Name, 1, 3) = "txt" Or Mid(Me.ActiveControl.Name, 1, 3) = "opt" Or Mid(Me.ActiveControl.Name, 1, 3) = "cbo" Then
                        SendKeys "{TAB}"
                    End If
                End If
            End If
        
        Case vbKeyEscape
            SSTab1.SelectedItem = 0
            Call EnableFrame(True)
            Call SendToBackDisc
            Call SendToBack
        
        Case vbKeyF2
            If Frame2.Enabled = False Then Exit Sub
            If txtROno.Text <> "" Then
                MsgBox "Estimate cannot be Edited When its already uploaded to Repair Order", vbInformation, "Info"
                Exit Sub
            End If
            
            If Picture1.Visible = False Then Exit Sub
            If chkParticipat.Value = 0 Then Exit Sub
            
            Call EnableFrame(False)
            Call StoreParticipationEntry(txtEstimateno)
            PicInsurance.Visible = True
            PicInsurance.ZOrder 0
                    
        Case vbKeyF3
            If Frame2.Enabled = False Then Exit Sub
            If txtROno.Text <> "" Then
                MsgBox "Estimate cannot be Edited When its already uploaded to Repair Order", vbInformation, "Info"
                Exit Sub
            End If
            
            If Picture1.Visible = True Then
                If SSTab1.SelectedItem <> 1 Then
                    SSTab1.SelectedItem = 1
                Else
                    If txtInvoiceNo.Text = "" Then
'                        Call EnableFrame(False)
'                        Call InitJobs
'                        Call AddJobs
                    End If
                End If
            End If
            
        Case vbKeyF4
            If Frame2.Enabled = False Then Exit Sub
            If txtROno.Text <> "" Then
                MsgBox "Estimate cannot be Edited When its already uploaded to Repair Order", vbInformation, "Info"
                Exit Sub
            End If
            
            If Picture1.Visible = True Then
                If SSTab1.SelectedItem <> 2 Then
                    SSTab1.SelectedItem = 2
                Else
                    If txtInvoiceNo.Text = "" Then
'                        Call EnableFrame(False)
'                        Call AddParts
                    End If
                End If
            End If
        
        Case vbKeyF5
            If Frame2.Enabled = False Then Exit Sub
            If txtROno.Text <> "" Then
                MsgBox "Estimate cannot be Edited When its already uploaded to Repair Order", vbInformation, "Info"
                Exit Sub
            End If
            
            If Picture1.Visible = True Then
                If SSTab1.SelectedItem <> 3 Then
                    SSTab1.SelectedItem = 3
                Else
                    If txtInvoiceNo.Text = "" Then
                        Call EnableFrame(False)
                        Call AddMaterials
                        
                        cboMatCode.Text = "MISC":                       Call cbomatcode_LostFocus
                        cboMatCode.Enabled = False:                     cboMaterial.Enabled = False
                        cmdMatDelete.Visible = False
                    End If
                End If
            End If
        
        Case vbKeyF6
            If Frame2.Enabled = False Then Exit Sub
            If txtROno.Text <> "" Then
                MsgBox "Estimate cannot be Edited When its already uploaded to Repair Order", vbInformation, "Info"
                Exit Sub
            End If
            
            If Picture1.Visible = True Then
                If SSTab1.SelectedItem <> 4 Then
                    SSTab1.SelectedItem = 4
                Else
                    If txtInvoiceNo.Text = "" Then
'                        Call EnableFrame(False)
'                        Call AddAccessories
                    End If
                End If
            End If
        
        Case vbKeyF7
            If Frame2.Enabled = False Then Exit Sub
            If txtROno.Text <> "" Then
                MsgBox "Estimate cannot be Edited When its already uploaded to Repair Order", vbInformation, "Info"
                Exit Sub
            End If
            
            If txtInvoiceNo.Text = "" Then
                If SSTab1.SelectedItem <> 0 Then
                    If SSTab1.SelectedItem = 1 Then
                        If lstJObs.ListItems.Count = 0 Then Exit Sub
                    ElseIf SSTab1.SelectedItem = 2 Then
                        If lstParts.ListItems.Count = 0 Then Exit Sub
                    ElseIf SSTab1.SelectedItem = 3 Then
                        If lstMaterials.ListItems.Count = 0 Then Exit Sub
                    Else
                        If lstAccessories.ListItems.Count = 0 Then Exit Sub
                    End If
                    
                    Call EnableFrame(False)
                    Call SendToFrontDisc
                End If
            End If
            
        Case vbKeyF12
            If Frame2.Enabled = False Then Exit Sub
            If txtROno.Text <> "" Then
                MsgBox "Estimate cannot be Edited When its already uploaded to Repair Order", vbInformation, "Info"
                Exit Sub
            End If
            
            If txtPlate_No.Text <> "" Then
                Call frm.SelectSQl("SELECT * FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(txtPlate_No) & "", "ESTIMATE MODULE", GetPlateId(txtPlate_No), txtAcct_No, txtNiym, txtPlate_No)
                frm.Show 1
            Else
                MsgBox "Plate Number must be inputed!" & vbCrLf & _
                    "Please enter 000000 if unknown", vbCritical, "No Plate No."
                Exit Sub
            End If
            
        Case Else
            MoveKeyPress KeyCode
            
    End Select
    
    If fraAddJobs.Enabled = True Then
        If Shift = 2 Then
            If KeyCode = vbKeyJ Then
                optByCode.Value = True: optByCode_Click
            End If
            If KeyCode = vbKeyD Then
                optByDescription.Value = True: optByDescription_Click
            End If
        End If
    End If
End Sub

Function GetPlateId(XPLATENO As String) As Long
    Dim RSTMP                               As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT ID FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(XPLATENO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetPlateId = RSTMP!ID
    End If
    Set RSTMP = Nothing
End Function

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    Set frm = New frmCSMSROCusveh
    Set FRMx = New frmCSMS_MasterSearchCustomer
    Set frmApp = New frmCSMS_UploadEstimate
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False

    RO_OR_ESTI_OR_PART = "ESTI": ESTISHOW = True: SendToBack
    
    DoEvents
    Call initMemvars
    Call InitCbo
    Call InitJobs
    Call InitParts
    Call InitMaterials
    Call InitAccessories
    
    Call rsRefresh
    
    If Not rsEsti_HD.EOF And Not rsEsti_HD.BOF Then rsEsti_HD.MoveLast
    
    DoEvents
    Call StoreMemVars
    
    txtJobLineNo.Text = ""
    Unload frmSplash
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ESTISHOW = False
    Set frmCSMSEstimateEntry = Nothing
End Sub

Private Sub frm_SaveChanges(xPLATE_NO As String, xWARR_CER As String, xMake As String, xMODEL As String, xSERIAL As String, xDESCRIPTION As Variant, FromFrom As String)
    If FromFrom = "ESTIMATE MODULE" Then
        txtPlate_No.Text = xPLATE_NO
        txtCertific8.Text = xWARR_CER
        'txtVIN.Text = xSERIAL
        cboModel.Text = xMODEL
        txtMake.Text = xDESCRIPTION
        
        Unload frm
    End If
End Sub

Private Sub frmApp_PutEstimateNo(ByVal xESTNO As String, FromForm As String)
    If FromForm = "ESTIMATE" Then
        Call rsRefresh
        rsEsti_HD.Find "ID = " & labid & ""
        Call StoreMemVars
        
        Unload frmApp
    End If
End Sub

Private Sub FRMx_SelectionMade(ByVal Xcode As String, xName As String, FromForm As String)
    If FromForm = "ESTIMATE MODULE" Then
        txtAcct_No.Text = Xcode
        txtNiym.Text = xName
        
        Unload FRMx
    ElseIf FromForm = "ESTIMATE INSURANCE" Then
        txtParticipat.Text = Xcode
        Text2.Text = xName
        
        Unload FRMx
    End If
End Sub

Private Sub grdDetails_DblClick()
    Dim XXX As String
    If grdDetails.Row = 1 Then SSTab1.SelectedItem = 1
    If grdDetails.Row = 2 Then SSTab1.SelectedItem = 2
    If grdDetails.Row = 3 Then SSTab1.SelectedItem = 3
    If grdDetails.Row = 4 Then SSTab1.SelectedItem = 4
End Sub

Private Sub lstAccessories_DblClick()
    If txtROno.Text <> "" Then
        MsgBox "Estimate cannot be Edited When its already uploaded to Repair Order", vbInformation, "Info"
        Exit Sub
    ElseIf CheckEstimateStatus(txtEstimateno) = "UPLOADED" Then
        MsgBox "Estimate already uploaded to Repair order. " & vbCrLf & " kindly refresh your Job Estimate Module to display fresh data.", vbInformation, "Info."
        Exit Sub
    End If
    
    If Acnt = 0 Then Exit Sub
    AddorEdit = "EDIT"

    Call SendToBack
    fraAddAccessories.ZOrder 0
    
    If txtInvoiceNo.Text = "" Then
        fraAddAccessories.Enabled = True
    Else
        fraAddAccessories.Enabled = False
    End If
    
    Call EnableFrame(False)
    ShortcutCaption5.Caption = "Edit Accessories"
    Call StoreAccEntry(Trim(Me.lstAccessories.SelectedItem))
End Sub

Private Sub lstJObs_DblClick()
    If txtROno.Text <> "" Then
        MsgBox "Estimate cannot be Edited When its already uploaded to Repair Order", vbInformation, "Info"
        Exit Sub
    ElseIf CheckEstimateStatus(txtEstimateno) = "UPLOADED" Then
        MsgBox "Estimate already uploaded to Repair order. " & vbCrLf & " kindly refresh your Job Estimate Module to display fresh data.", vbInformation, "Info."
        Exit Sub
    End If
    
    
    If ESTIKCNT = 0 Then Exit Sub
    AddorEdit = "EDIT": SendToBack
    fraAddJobs.ZOrder 0

    If txtInvoiceNo.Text = "" Then
        fraAddJobs.Enabled = True
    Else
        fraAddJobs.Enabled = False
    End If
    
    Call EnableFrame(False)
    ShortcutCaption1.Caption = "Edit Jobs"
    Call StoreJobsEntry(Trim(Me.lstJObs.SelectedItem))
End Sub

Private Sub lstMaterials_DblClick()
    Dim nix As String
    
    If txtROno.Text <> "" Then
        MsgBox "Estimate cannot be Edited When its already uploaded to Repair Order", vbInformation, "Info"
        Exit Sub
    ElseIf CheckEstimateStatus(txtEstimateno) = "UPLOADED" Then
        MsgBox "Estimate already uploaded to Repair order. " & vbCrLf & " kindly refresh your Job Estimate Module to display fresh data.", vbInformation, "Info."
        Exit Sub
    End If
    
    If Mcnt = 0 Then Exit Sub
    AddorEdit = "EDIT"
    Call SendToBack
    fraAddMaterials.ZOrder 0
    
    If txtInvoiceNo.Text = "" Then
        fraAddMaterials.Enabled = True
    Else
        fraAddMaterials.Enabled = False
    End If
        
    cboMatCode.Enabled = True:          cboMaterial.Enabled = True
    cmdMatDelete.Visible = True
    Call EnableFrame(False)
    ShortcutCaption3.Caption = "Edit Materials"
    'Call StoreMatEntry(Trim(Me.lstMaterials.SelectedItem))
    
    nix = Trim(Me.lstMaterials.SelectedItem.ListSubItems(8))
    
    
    Call StoreMatEntry(nix)
    
End Sub

Private Sub lstParts_DblClick()
    If txtROno.Text <> "" Then
        MsgBox "Estimate cannot be Edited When its already uploaded to Repair Order", vbInformation, "Info"
        Exit Sub
    ElseIf CheckEstimateStatus(txtEstimateno) = "UPLOADED" Then
        MsgBox "Estimate already uploaded to Repair order. " & vbCrLf & " kindly refresh your Job Estimate Module to display fresh data.", vbInformation, "Info."
        Exit Sub
    End If
    
    If Pcnt = 0 Then Exit Sub
    AddorEdit = "EDIT"
    SendToBack
    fraAddParts.ZOrder 0

    If txtInvoiceNo.Text = "" Then
        fraAddParts.Enabled = True
    Else
        fraAddParts.Enabled = False
    End If
    
    Call EnableFrame(False)
    ShortcutCaption4.Caption = "Edit Parts"
    Call StorePartsEntry(Trim(Me.lstParts.SelectedItem))
End Sub

Private Sub cbomatcode_LostFocus()
    If cboMatCode.Text = "MISC" Then
        cboMatCode.Text = UCase(cboMatCode.Text)
        cboMaterial.Text = SetMatDisc(cboMatCode.Text)
    Else
        cboMatCode.Text = UCase(cboMatCode.Text)
        cboMaterial.Text = SetMatDisc(cboMatCode.Text)
        txtMatUnitPrice.Text = SetMatPrice(cboMatCode.Text)
        txtMatPOCode.Text = SetMatPOCode(cboMatCode.Text)
        txtMatAmount.Text = txtMatUnitPrice.Text
    End If
End Sub

Private Sub optByCode_Click()
'    cboJcode.Enabled = True
'    DoEvents
'    On Error Resume Next
'    cboJcode.SetFocus
'    cboJobCode.Enabled = False
End Sub

Private Sub optByDescription_Click()
'    cboJobCode.Enabled = True
'    DoEvents
'    On Error Resume Next
'    cboJobCode.SetFocus
'    cboJcode.Enabled = False
End Sub

Private Sub TabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

End Sub

Private Sub Picture3_Click()

End Sub

Private Sub Timer1_Timer()
    If lblSTATUS.ForeColor = vbRed Then
        lblSTATUS.ForeColor = vbBlue
    Else
        lblSTATUS.ForeColor = vbRed
    End If
End Sub

Private Sub txtAccAmount_GotFocus()
    txtAccAmount.Text = NumericVal(txtAccQty.Text) * NumericVal(txtAccUnitPrice.Text)
End Sub

Private Sub txtAccAmount_LostFocus()
    txtAccAmount.Text = Format(txtAccAmount.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtAccDiscount_LostFocus()
    If NumericVal(txtAccDiscount.Text) > "100" Then
        MsgBox "Invalid discount", vbInformation + vbOKOnly
        On Error Resume Next
        txtAccDiscount.SetFocus
        Exit Sub
    Else
        txtAccDiscount.Text = Format(txtAccDiscount.Text, "##0.0")
    End If
End Sub

Private Sub txtAccQty_Change()
    If txtAccQty.Text <> "" Then
        txtAccAmount.Text = NumericVal(txtAccQty.Text) * NumericVal(txtAccUnitPrice.Text)
    End If
End Sub

Private Sub txtAccQty_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    Else
    End If
End Sub

Private Sub txtAccQty_LostFocus()
    If txtAccQty.Text <> "" Then
        txtAccAmount.Text = Format(NumericVal(txtAccQty.Text) * NumericVal(txtAccUnitPrice.Text), "#####0.0")
    End If
End Sub

Private Sub txtAcct_No_Change()
    If txtAcct_No.Text <> "" Then
        Set rsCusmas = New ADODB.Recordset
        rsCusmas.Open "select cuscde,cusnam,cusadd from ALL_CUSMAS where cuscde = '" & UCase(txtAcct_No.Text) & "'", gconDMIS
        If Not rsCusmas.EOF And Not rsCusmas.BOF Then
            txtNiym.Text = Null2String(rsCusmas!CUSNAM)
            txtAddress.Text = Null2String(rsCusmas!Cusadd)
        End If
    End If
End Sub

Private Sub txtAcct_No_LostFocus()
    If txtAcct_No.Text <> "" Then
        txtAcct_No.Text = UCase(txtAcct_No.Text)
        Set rsCusmas = New ADODB.Recordset
        rsCusmas.Open "select cuscde,cusnam,cusadd from ALL_CUSMAS where cuscde = '" & txtAcct_No.Text & "'", gconDMIS
        If Not rsCusmas.EOF And Not rsCusmas.BOF Then
            txtNiym.Text = Null2String(rsCusmas!CUSNAM)
            txtAddress.Text = Null2String(rsCusmas!Cusadd)
        Else
            MsgBoxXP "Invalid CUSMAS Code!", "Error", XP_OKOnly, msg_Critical
            txtNiym.Text = "": txtAddress.Text = ""
        End If
    End If
End Sub

Private Sub txtAccUnitPrice_Change()
    If txtAccUnitPrice.Text <> "" Then
        txtAccAmount.Text = NumericVal(txtAccQty.Text) * NumericVal(txtAccUnitPrice.Text)
    End If
End Sub

Private Sub txtAccUnitPrice_LostFocus()
    If txtAccUnitPrice.Text <> "" Then
        txtAccAmount.Text = Format(NumericVal(txtAccQty.Text) * NumericVal(txtAccUnitPrice.Text), MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtCertific8_Click()
'    If txtPlate_No.Text <> "" Then
'        Me.Enabled = False
'        frmCSMSESTICusveh.Show
'        frmCSMSESTICusveh.ZOrder 0
'    Else
'        MsgBoxXP "Plate Number must be inputed!" & vbCrLf & _
'                 "Please enter 000000 if unknown", "No Plate No.!", XP_OKOnly, msg_Critical
'    End If
End Sub

Private Sub txtDiscAmt_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    Else
    End If
End Sub

Private Sub txtDiscAmt_LostFocus()
    txtDiscAmt.Text = Format(txtDiscAmt.Text, "##0.0")
End Sub

Private Sub txtDte_comp_LostFocus()
    If txtDte_comp.Text <> "" Then txtDte_comp.Text = Format(txtDte_comp.Text, "Short Date")
End Sub

Private Sub txtDte_recd_LostFocus()
    If txtDte_recd.Text <> "" Then txtDte_recd.Text = Format(txtDte_recd.Text, "Short Date")
End Sub

Private Sub txtDte_Rel_LostFocus()
    If txtDte_Rel.Text <> "" Then txtDte_Rel.Text = Format(txtDte_Rel.Text, "Short Date")
End Sub

Private Sub txtJobDiscount_LostFocus()
    If NumericVal(txtJobDiscount.Text) > "100" Then
        MsgBox "Invalid discount", vbInformation + vbOKOnly
        On Error Resume Next
        txtJobDiscount.SetFocus
        Exit Sub
    Else
        txtJobDiscount.Text = Format(txtJobDiscount.Text, MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtJobRate_LostFocus()
    txtJobRate.Text = Format(txtJobRate.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtLOAAmount_Change()
    Dim RO_JOB                                      As Double
    Dim RO_PARTS                                    As Double
    Dim RO_MATS                                     As Double
    Dim RO_ACCS                                     As Double

    RO_JOB = TOTJOBAMT + JobInsTotal
    RO_PARTS = TOTPARTSAMT + PartsInsTotal
    RO_MATS = TOTMATAMT + MatInsTotal
    RO_ACCS = TOTACCAMT + AccInsTotal

    If chkAllowManDist.Value = 0 Then
        If NumericVal(txtLOAAmount.Text) > NumericVal(RO_JOB + RO_PARTS + RO_MATS + RO_ACCS) Then
            MsgBox "Warning: LOA Amount should not Exceed Repair Order Total Amount.", vbCritical, "Not Allowed!"
            txtLOAAmount.Text = NumericVal(txtPartTotal.Text)
            Exit Sub
        End If

        If NumericVal(txtLOAAmount.Text) > RO_JOB Then
            txtPartLabor.Text = RO_JOB
            If NumericVal(txtLOAAmount.Text) - RO_JOB > RO_PARTS Then
                txtPartParts.Text = RO_PARTS
                If NumericVal(txtLOAAmount.Text) - RO_JOB - RO_PARTS > RO_MATS Then
                    txtPartMaterials.Text = RO_MATS
                Else
                    txtPartMaterials.Text = NumericVal(txtLOAAmount.Text) - RO_JOB - RO_PARTS
                    If NumericVal(txtLOAAmount.Text) - RO_JOB - RO_PARTS - RO_MATS > RO_ACCS Then
                        txtPartAccessories.Text = NumericVal(txtLOAAmount.Text) - RO_JOB - RO_PARTS - RO_MATS - RO_ACCS
                    Else
                        txtPartAccessories.Text = RO_ACCS
                    End If
                End If
            Else
                txtPartParts.Text = NumericVal(txtLOAAmount.Text) - RO_JOB
            End If
        Else
            txtPartLabor.Text = NumericVal(txtLOAAmount.Text)
            txtPartParts.Text = ZERO
            txtPartMaterials.Text = ZERO
            txtPartAccessories.Text = ZERO
        End If
    End If
End Sub

Private Sub txtLOAAmount_GotFocus()
    txtLOAAmount.Text = NumericVal(txtLOAAmount.Text)
End Sub

Private Sub txtLOAAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtMatAmount_GotFocus()
    txtMatAmount.Text = NumericVal(txtMatQty.Text) * NumericVal(txtMatUnitPrice.Text)
End Sub
Private Sub txtMatAmount_LostFocus()
    txtMatAmount.Text = Format(txtMatAmount.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtMatDiscount_LostFocus()
Dim i As Integer
    If NumericVal(txtMatDiscount.Text) > "100" Then
        MsgBox "Invalid discount", vbInformation + vbOKOnly
        On Error Resume Next
        txtMatDiscount.SetFocus
        Exit Sub
    Else
        txtMatDiscount.Text = Format(txtMatDiscount.Text, "##0.0")
    End If
End Sub

Private Sub txtMatQty_Change()
    If txtMatQty.Text <> "" Then
        txtMatAmount.Text = NumericVal(txtMatQty.Text) * NumericVal(txtMatUnitPrice.Text)
    End If
End Sub

Private Sub txtMatQty_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    Else
    End If
End Sub

Private Sub txtMatQty_LostFocus()
    If txtMatQty.Text <> "" Then
        txtMatAmount.Text = Format(NumericVal(txtMatQty.Text) * NumericVal(txtMatUnitPrice.Text), "#####0.0")
    End If
End Sub

Private Sub txtMatUnitPrice_Change()
    If txtMatUnitPrice.Text <> "" Then
        txtMatAmount.Text = NumericVal(txtMatQty.Text) * NumericVal(txtMatUnitPrice.Text)
    End If
End Sub



Private Sub txtMatUnitPrice_LostFocus()
    If txtMatUnitPrice.Text <> "" Then
        txtMatAmount.Text = Format(NumericVal(txtMatQty.Text) * NumericVal(txtMatUnitPrice.Text), MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtPart_amt_LostFocus()
    txtPart_amt.Text = Format(txtPart_amt.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtPartAccessories_Change()
    If chkAllowManDist.Value = 1 Then
        If NumericVal(txtPartAccessories.Text) > ((TOTACCAMT + AccInsTotal) - N2Str2IntZero(rsEsti_HD!a_discount)) Then
            MsgBox "Warning: System Doesn't allow Participation to Exceed Actual Accessories Amount" & vbCrLf & "                Actual Accessories Amount will be set as default", vbCritical, "Not Allowed!"
            txtPartAccessories.Text = ((TOTACCAMT + AccInsTotal) - N2Str2IntZero(rsEsti_HD!a_discount))
        End If
    End If
    Call SetTotalParticipation
End Sub

Private Sub txtPartAmount_GotFocus()
    txtPartAmount.Text = NumericVal(txtQTY.Text) * NumericVal(txtUnitPrice.Text)
End Sub

Private Sub cboPartno_LostFocus()
    If cboPartNo.Text <> "" Then cboDescription.Text = SetPartDisc(cboPartNo.Text)
    txtUnitPrice.Text = SetPartPrice(cboPartNo.Text)
    txtPartAmount.Text = ToDoubleNumber(NumericVal(txtQTY.Text) * NumericVal(txtUnitPrice.Text))
End Sub

Private Sub txtPartAmount_LostFocus()
    txtPartAmount.Text = Format(txtPartAmount.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtPartDiscount_LostFocus()
    If txtPartDiscount.Text > 100 Then
        MsgBox "Invalid Discount", vbInformation + vbOKOnly
        On Error Resume Next
        txtPartDiscount.SetFocus
        Exit Sub
    Else
        txtPartDiscount.Text = Format(txtPartDiscount.Text, "##0.0")
    End If
End Sub

Private Sub txtParticipat_Change()
    If txtParticipat.Text <> "" And chkParticipat.Value = 1 Then
        txtParticipat.Text = UCase(txtParticipat.Text)
        Set rsCusmas = New ADODB.Recordset
        Set rsCusmas = gconDMIS.Execute("select cuscde,cusnam,cusadd from ALL_CUSMAS where cuscde = '" & txtAcct_No.Text & "'")
        If Not rsCusmas.EOF And Not rsCusmas.BOF Then
            txtNiym.Text = Null2String(rsCusmas!CUSNAM)
            txtAddress.Text = Null2String(rsCusmas!Cusadd)
        End If
        Set rsCusmas = New ADODB.Recordset
        Set rsCusmas = gconDMIS.Execute("select cuscde,cusnam from ALL_CUSMAS where cuscde = '" & txtParticipat.Text & "'")
        If Not rsCusmas.EOF And Not rsCusmas.BOF Then
            txtNiym.Text = txtNiym.Text & "/" & Null2String(rsCusmas!CUSNAM)
        End If
    End If
End Sub

Private Sub txtParticipat_LostFocus()
    If txtParticipat.Text <> "" And chkParticipat.Value = 1 Then
        txtParticipat.Text = UCase(txtParticipat.Text)
        Set rsCusmas = New ADODB.Recordset
        Set rsCusmas = gconDMIS.Execute("select cuscde,cusnam,cusadd from ALL_CUSMAS where cuscde = '" & txtAcct_No.Text & "'")
        If Not rsCusmas.EOF And Not rsCusmas.BOF Then
            txtNiym.Text = Null2String(rsCusmas!CUSNAM)
            txtAddress.Text = Null2String(rsCusmas!Cusadd)
        End If
        Set rsCusmas = New ADODB.Recordset
        Set rsCusmas = gconDMIS.Execute("select cuscde,cusnam from ALL_CUSMAS where cuscde = '" & txtParticipat.Text & "'")
        If Not rsCusmas.EOF And Not rsCusmas.BOF Then
            txtNiym.Text = txtNiym.Text & "/" & Null2String(rsCusmas!CUSNAM)
        End If
    End If
End Sub

Private Sub txtPartLabor_Change()
    If chkAllowManDist.Value = 1 Then
        If NumericVal(txtPartLabor.Text) > ((TOTJOBAMT + JobInsTotal) - N2Str2IntZero(rsEsti_HD!l_discount)) Then
            MsgBox "Warning: System Doesn't allow Participation to Exceed Actual Job Amount" & vbCrLf & "                Actual Job Amount will be set as default", vbCritical, "Not Allowed!"
            txtPartLabor.Text = ((TOTJOBAMT + JobInsTotal) - N2Str2IntZero(rsEsti_HD!l_discount))
        End If
    End If
    Call SetTotalParticipation
End Sub

Private Sub txtPartMaterials_Change()
    If chkAllowManDist.Value = 1 Then
        If NumericVal(txtPartMaterials.Text) > ((TOTMATAMT + MatInsTotal) - (N2Str2Zero(rsEsti_HD!m_discount))) Then
            MsgBox "Warning: System Doesn't allow Participation to Exceed Actual Materials Amount" & vbCrLf & "                Actual Materials Amount will be set as default", vbCritical, "Not Allowed!"
            txtPartMaterials.Text = ((TOTMATAMT + MatInsTotal) - (N2Str2Zero(rsEsti_HD!m_discount)))
        End If
    End If
    Call SetTotalParticipation
End Sub

Private Sub txtPartParts_Change()
    If chkAllowManDist.Value = 1 Then
        If NumericVal(txtPartParts.Text) > ((TOTPARTSAMT + PartsInsTotal) - N2Str2IntZero(rsEsti_HD!p_discount)) Then
            MsgBox "Warning: System Doesn't allow Participation to Exceed Actual Parts Amount" & vbCrLf & "                Actual Parts Amount will be set as default", vbCritical, "Not Allowed!"
            txtPartParts.Text = ((TOTPARTSAMT + PartsInsTotal) - N2Str2IntZero(rsEsti_HD!p_discount))
        End If
    End If
    Call SetTotalParticipation
End Sub

Private Sub txtQTY_Change()
    If txtQTY.Text <> "" Then txtPartAmount.Text = NumericVal(txtQTY.Text) * NumericVal(txtUnitPrice.Text)
End Sub

Private Sub txtQTY_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    Else
    End If
End Sub

Private Sub txtQty_LostFocus()
    txtQTY.Text = Format(txtQTY.Text, "########0.0")
End Sub

Private Sub txtUnitPrice_Change()
    If txtUnitPrice.Text <> "" Then
        txtPartAmount.Text = NumericVal(txtQTY.Text) * NumericVal(txtUnitPrice.Text)
    End If
End Sub

Private Sub txtUnitPrice_LostFocus()
    txtUnitPrice.Text = Format(txtUnitPrice.Text, MAXIMUM_DIGIT)
End Sub

Sub EnableFrame(COND As Boolean)
    Frame2.Enabled = COND
    Picture1.Enabled = COND
End Sub
Function getro(XXX As String) As String
    Dim rsREPOR             As ADODB.Recordset
    
    Set rsREPOR = New ADODB.Recordset
    Set rsREPOR = gconDMIS.Execute("Select rep_or from CSMS_REPOR where EstimateNo = '" & XXX & "'")
    If Not (rsREPOR.EOF And rsREPOR.BOF) Then
        getro = Null2String(rsREPOR!REP_OR)
    Else
        getro = ""
    End If
Set rsREPOR = Nothing
End Function

Sub getinvoiceandpayterm()
Dim rsREPOR             As ADODB.Recordset

Set rsREPOR = New ADODB.Recordset
Set rsREPOR = gconDMIS.Execute("Select term,invoice from CSMS_REPOR where EstimateNo = '" & txtEstimateno.Text & "'")
If Not (rsREPOR.EOF And rsREPOR.BOF) Then
    txtInvoiceNo.Text = Null2String(rsREPOR!INVOICE)
    txtTerm.Text = Null2String(rsREPOR!TERM)
Else
    txtInvoiceNo.Text = ""
    txtTerm.Text = ""
End If
Set rsREPOR = Nothing
End Sub





















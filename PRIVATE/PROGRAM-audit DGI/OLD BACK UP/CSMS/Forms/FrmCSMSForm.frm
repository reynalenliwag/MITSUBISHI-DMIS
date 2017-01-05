VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form FrmCSMSComplaintsForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complaint Form"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCSMSForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   6390
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   60
      ScaleHeight     =   5865
      ScaleWidth      =   6255
      TabIndex        =   27
      Top             =   30
      Width           =   6285
      Begin VB.TextBox cbocustomer 
         Height          =   345
         Left            =   1620
         TabIndex        =   45
         Top             =   2730
         Width           =   4425
      End
      Begin VB.TextBox txtPrepared 
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
         Left            =   1620
         TabIndex        =   3
         Top             =   1590
         Width           =   4035
      End
      Begin VB.TextBox txtComplaint 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         TabIndex        =   11
         Top             =   5070
         Width           =   2865
      End
      Begin VB.TextBox txtItem 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         TabIndex        =   10
         Top             =   4665
         Width           =   2865
      End
      Begin VB.TextBox txtMileage 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         TabIndex        =   9
         Top             =   4275
         Width           =   1335
      End
      Begin VB.TextBox txtVin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         TabIndex        =   8
         Top             =   3900
         Width           =   2865
      End
      Begin VB.TextBox txtveh 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3135
         Width           =   2865
      End
      Begin VB.TextBox txtDealer 
         Height          =   330
         Left            =   1620
         TabIndex        =   5
         Top             =   2370
         Width           =   1665
      End
      Begin VB.TextBox txtAttention 
         Height          =   345
         Left            =   1620
         TabIndex        =   0
         Top             =   420
         Width           =   4425
      End
      Begin VB.TextBox txtFax 
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
         Left            =   1620
         TabIndex        =   2
         Top             =   1200
         Width           =   3045
      End
      Begin VB.TextBox txtTel 
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
         Left            =   1620
         TabIndex        =   1
         Top             =   810
         Width           =   3045
      End
      Begin MSComCtl2.DTPicker DTPDateRequested 
         Height          =   315
         Left            =   1620
         TabIndex        =   4
         Top             =   2010
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56360961
         CurrentDate     =   39248
      End
      Begin MSComCtl2.DTPicker DTPAcquisitionDate 
         Height          =   315
         Left            =   1620
         TabIndex        =   7
         Top             =   3525
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56360961
         CurrentDate     =   39248
      End
      Begin VB.Label labid 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   6030
         TabIndex        =   46
         Top             =   1020
         Visible         =   0   'False
         Width           =   345
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   9225
         _Version        =   655364
         _ExtentX        =   16272
         _ExtentY        =   556
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   705
         TabIndex        =   41
         Top             =   2850
         Width           =   840
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complaints(s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   40
         Top             =   5130
         Width           =   1185
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Installed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   405
         TabIndex        =   39
         Top             =   4770
         Width           =   1140
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Mileage "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   38
         Top             =   4350
         Width           =   1395
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V.I.N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1170
         TabIndex        =   37
         Top             =   3990
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acquisition Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   36
         Top             =   3585
         Width           =   1380
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Left            =   1035
         TabIndex        =   35
         Top             =   3240
         Width           =   510
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DealerShip"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   615
         TabIndex        =   34
         Top             =   2430
         Width           =   930
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Requested"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   33
         Top             =   2085
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attention"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   780
         TabIndex        =   32
         Top             =   540
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   915
         TabIndex        =   31
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel.No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1020
         TabIndex        =   30
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label13 
         Caption         =   "Note: Kindly attach xerox copy of OR/CR of the vehicle on this Form"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   5490
         Width           =   5175
      End
   End
   Begin Crystal.CrystalReport rptComplaintReport 
      Left            =   90
      Top             =   6030
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox PicControl 
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
      Left            =   570
      ScaleHeight     =   885
      ScaleWidth      =   5835
      TabIndex        =   20
      Top             =   5910
      Width           =   5835
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
         Height          =   765
         Left            =   5040
         MouseIcon       =   "FrmCSMSForm.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMSForm.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Height          =   765
         Left            =   4320
         MouseIcon       =   "FrmCSMSForm.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMSForm.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Height          =   765
         Left            =   3600
         MouseIcon       =   "FrmCSMSForm.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMSForm.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Height          =   765
         Left            =   2880
         MouseIcon       =   "FrmCSMSForm.frx":1E6F
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMSForm.frx":1FC1
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Height          =   765
         Left            =   2160
         MouseIcon       =   "FrmCSMSForm.frx":231D
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMSForm.frx":246F
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Height          =   765
         Left            =   1440
         MouseIcon       =   "FrmCSMSForm.frx":2782
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMSForm.frx":28D4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   720
         MouseIcon       =   "FrmCSMSForm.frx":2BCE
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMSForm.frx":2D20
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   0
         MouseIcon       =   "FrmCSMSForm.frx":3078
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMSForm.frx":31CA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox picSave 
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
      Left            =   660
      ScaleHeight     =   885
      ScaleWidth      =   5790
      TabIndex        =   21
      Top             =   5970
      Width           =   5790
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
         Height          =   765
         Left            =   4935
         MouseIcon       =   "FrmCSMSForm.frx":3529
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMSForm.frx":367B
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancel"
         Top             =   30
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
         Height          =   765
         Left            =   4215
         MouseIcon       =   "FrmCSMSForm.frx":39B9
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMSForm.frx":3B0B
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.PictureBox PicSearch 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6795
      Left            =   60
      ScaleHeight     =   6795
      ScaleWidth      =   6285
      TabIndex        =   24
      Top             =   30
      Visible         =   0   'False
      Width           =   6285
      Begin XtremeReportControl.ReportControl ReportControl 
         Height          =   5865
         Left            =   90
         TabIndex        =   29
         Top             =   810
         Width           =   6075
         _Version        =   655364
         _ExtentX        =   10716
         _ExtentY        =   10345
         _StockProps     =   64
      End
      Begin VB.CommandButton Command 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5700
         TabIndex        =   44
         Top             =   390
         Width           =   495
      End
      Begin VB.TextBox txtkeyword 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   25
         Top             =   390
         Width           =   5565
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
         Height          =   345
         Left            =   0
         TabIndex        =   26
         Top             =   -30
         Width           =   6675
         _Version        =   655364
         _ExtentX        =   11774
         _ExtentY        =   609
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
      End
   End
End
Attribute VB_Name = "FrmCSMSComplaintsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thecode                                             As String
Dim UPDATE_MODE                                         As Boolean
Dim theID                                               As String
Dim ADD_OR_EDIT                                         As String
Dim rsComplain                                          As ADODB.Recordset

Sub displayVehicle()
    Dim RS                                             As New ADODB.Recordset    '

    Set RS = gconDMIS.Execute("SELECT model,vin from CSMS_cusveh Where cuscde='" & thecode & "'")
    With RS
        If Not .EOF And Not .BOF Then
            txtveh.Text = Null2String(!Model)
            txtVin.Text = Null2String(!VIN)
        End If
    End With
    Set RS = Nothing
End Sub

Sub rsRefresh()
    Set rsComplain = New ADODB.Recordset
    rsComplain.Open "SELECT * From CSMS_ComplaintsForm  order by id", gconDMIS, adOpenKeyset, adLockOptimistic
End Sub

Sub StoreMemvars()
    If Not (rsComplain.BOF And rsComplain.EOF) Then
        labid.Caption = Null2String(rsComplain!ID)
        theID = Null2String(rsComplain!ID)
        txtAttention.Text = Null2String(rsComplain!Attention)
        cbocustomer = Null2String(rsComplain!CustomerName)
        DTPDateRequested = Null2String(rsComplain!DateRequested)
        txtTel.Text = Null2String(rsComplain!TelNo)
        txtFax = Null2String(rsComplain!Telefax)
        DTPAcquisitionDate = Null2String(rsComplain!AcquisitionDate)
        txtDealer.Text = Null2String(rsComplain!DealerShip)
        txtVin.Text = Null2String(rsComplain!VIN)
        txtveh.Text = Null2String(rsComplain!Model)
        txtMileage.Text = Null2String(rsComplain!CurrentMileage)
        txtItem.Text = Null2String(rsComplain!ItemInstalled)
        txtComplaint.Text = Null2String(rsComplain!Complaint)
        txtPrepared.Text = Null2String(rsComplain!PreparedBy)
    Else
        Call ShowNoRecord
        Call cmdAdd_Click
    End If
End Sub

Sub InitMemVars()
    theID = ""
    txtAttention.Text = ""
    cbocustomer = ""
    DTPDateRequested = Date
    txtTel.Text = ""
    txtFax = ""
    DTPAcquisitionDate = Date
    txtDealer.Text = ""
    txtVin.Text = ""
    txtveh.Text = "'"
    txtMileage.Text = ""
    txtItem.Text = ""
    txtComplaint.Text = ""
    txtPrepared.Text = ""
End Sub


Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "COMPLAINTS FORM") = False Then Exit Sub
    ADD_OR_EDIT = "ADD"
    
    Call InitMemVars
    PicControl.Visible = False
    picSave.Visible = True
    picMain.Enabled = True
        
    On Error Resume Next
    txtAttention.SetFocus
End Sub

Private Sub cmdCancel_Click()
    picSave.Visible = False
    PicControl.Visible = True
    picMain.Enabled = False
    
    Call StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "COMPLAINTS FORM") = False Then Exit Sub

    If MsgBox("delete this record, Are You Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub

    gconDMIS.Execute ("DELETE FROM CSMS_complaintsForm Where ID = '" & labid & "'")
    Call ShowDeletedMsg
    Call rsRefresh
    Call StoreMemvars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "COMPLAINTS FORM") = False Then Exit Sub
    ADD_OR_EDIT = "EDIT"
    
    picSave.Visible = True
    PicControl.Visible = False
    picMain.Enabled = True
    
    On Error Resume Next
    txtAttention.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Call FillConcern
    PicSearch.Visible = True
    PicSearch.ZOrder 0
    On Error Resume Next
    txtkeyword.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsComplain.MoveNext
    If rsComplain.EOF Then
        rsComplain.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    rsComplain.MovePrevious
    If rsComplain.BOF Then
        rsComplain.MoveFirst
        Call ShowLastRecordMsg
    End If
    Call StoreMemvars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "COMPLAINTS FORM") = False Then Exit Sub
    
    Screen.MousePointer = 11
    PrintSQLReport rptComplaintReport, CSMS_REPORT_PATH & "CSMS_complaintsReport.rpt", "{CSMS_complaintsform.id} = " & labid.Caption & "", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    Dim SQL                                             As String
    Dim TheAttention                                    As String
    Dim Thetelefax                                      As String
    Dim Thetelno                                        As String
    Dim TheDealer                                       As String
    Dim theCustomer                                     As String
    Dim themodel                                        As String
    Dim TheVin                                          As String
    Dim theMilleage                                     As String
    Dim theItem                                         As String
    Dim theComplaints                                   As String
    Dim thePreparedBy                                   As String
    Dim theAcquisitionDate                              As String
    Dim ThedateRequested                                As String
    Dim RSTMP                                           As New ADODB.Recordset
    
    If Len(txtAttention) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Attention!"
        On Error Resume Next
        txtAttention.SetFocus
        Exit Sub
    End If

'    If Len(txtFax) = 0 Then
'        ShowIsRequiredMsg "Missing Parameters...Fax No!"
'        On Error Resume Next
'        txtFax.SetFocus
'        Exit Sub
'    End If

    If Len(txtTel) = 0 Then
        ShowIsRequiredMsg "Missing Parameters...tel No!!"
        On Error Resume Next
        txtTel.SetFocus
        Exit Sub
    End If

    If Len(txtDealer) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Dealer Ship!"
        On Error Resume Next
        txtDealer.SetFocus
        Exit Sub
    End If

    If Len(cbocustomer) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Customer!"
        On Error Resume Next
        cbocustomer.SetFocus
        Exit Sub
    End If

    If Len(txtveh) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Vehicle Type!"
        On Error Resume Next
        txtveh.SetFocus
        Exit Sub
    End If

    If Len(txtVin) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Vin "
        On Error Resume Next
        txtVin.SetFocus
        Exit Sub
    End If

    If Len(txtItem) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Item Installed"
        On Error Resume Next
        txtItem.SetFocus
        Exit Sub
    End If

    If Len(txtComplaint) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Complaints Field!"
        On Error Resume Next
        txtComplaint.SetFocus
        Exit Sub
    End If

    If Len(txtPrepared) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Prepared By!"
        On Error Resume Next
        txtPrepared.SetFocus
        Exit Sub
    End If

    If Len(txtMileage) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Current Milleage!"
        On Error Resume Next
        txtMileage.SetFocus
        Exit Sub
    End If

    TheAttention = N2Str2Null(Trim(txtAttention.Text))
    Thetelefax = N2Str2Null(Trim(txtFax.Text))
    Thetelno = N2Str2Null(Trim(txtTel.Text))
    ThedateRequested = N2Str2Null(DTPDateRequested)
    theAcquisitionDate = N2Str2Null(DTPAcquisitionDate)
    TheDealer = N2Str2Null(Trim(txtDealer.Text))
    theCustomer = N2Str2Null(cbocustomer)
    themodel = N2Str2Null(Trim(txtveh.Text))
    TheVin = N2Str2Null(Trim(txtVin.Text))
    theMilleage = N2Str2Null(Trim(txtMileage.Text))
    theItem = N2Str2Null(Trim(txtItem.Text))
    theComplaints = N2Str2Null(Trim(txtComplaint.Text))
    thePreparedBy = N2Str2Null(Trim(txtPrepared.Text))
    
    If ADD_OR_EDIT = "ADD" Then
        SQL = "INSERT INTO CSMS_ComplaintsForm " & _
            " VALUES(" & TheAttention & _
            ", " & Thetelefax & _
            ", " & Thetelno & _
            ", " & ThedateRequested & _
            ", " & theCustomer & _
            ", " & TheDealer & _
            ", " & themodel & _
            ", " & theAcquisitionDate & _
            ", " & TheVin & _
            ", " & theMilleage & _
            ", " & theItem & _
            ", " & theComplaints & _
            ", " & thePreparedBy & ")"

        gconDMIS.Execute (SQL)
        Call ShowSuccessFullyAdded
        Set RSTMP = gconDMIS.Execute("SELECT MAX(ID) AS ID FROM CSMS_ComplaintsForm")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            labid.Caption = RSTMP!ID
        End If
        Set RSTMP = Nothing
    Else
        SQL = "UPDATE CSMS_ComplaintsForm set " & _
            " attention = " & TheAttention & _
            ", telefax = " & Thetelefax & _
            ", telno = " & Thetelno & _
            ", dateRequested = " & ThedateRequested & _
            ", customername = " & theCustomer & _
            ", DealerShip = " & TheDealer & _
            ", Model = " & themodel & _
            ", Acquisitiondate = " & theAcquisitionDate & _
            ", Vin = " & TheVin & _
            ", CurrentMileage = " & theMilleage & _
            ", ItemInstalled = " & theItem & _
            ", Complaint = " & theComplaints & _
            ", PreparedBy = " & thePreparedBy & _
            " WHERE ID = " & theID & ""
        
        gconDMIS.Execute (SQL)
        Call ShowSuccessFullyUpdated
    End If
    
    Call cmdCancel_Click
    Call rsRefresh
    rsComplain.Find "ID = " & labid & ""

    Exit Sub
Errorcode:
    Call ShowVBError
    Exit Sub
End Sub

Private Sub Command_Click()
    PicSearch.Visible = False
    PicSearch.ZOrder 1
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
        
    With ReportControl
        .Columns.Add 0, "Attention", 100, True
        .Columns.Add 1, "Customer Name", 100, True
        .Columns.Add 2, "Vin no", 100, True
        .Columns.Add 3, "Model", 100, True
        .Columns.Add 4, "Compliant", 100, True
        .Columns.Add 6, "ID", 0, False
        
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.GroupRowTextBold = True
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots
        .SetCustomDraw xtpCustomBeforeDrawRow
    End With
    
    Call rsRefresh
    Call StoreMemvars
End Sub

Sub FillConcern()
    Dim RecSet                                         As New ADODB.Recordset
    Dim fld                                            As Field
    Dim j                                              As Long
    Dim REC                                            As XtremeReportControl.ReportRecord
    Set RecSet = gconDMIS.Execute("SELECT ATTENTION, CUSTOMERNAME, VIN, MODEL, COMPLAINT, ID FROM CSMS_ComplaintsForm order by id desc")
    
    ReportControl.Records.DeleteAll
    While Not RecSet.EOF
        j = j + 1
        Set REC = ReportControl.Records.Add
        For Each fld In RecSet.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        RecSet.MoveNext
    Wend
    
    ReportControl.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RecSet = Nothing
End Sub

Private Sub txtkeyword_Change()
    ReportControl.FilterText = txtkeyword
    ReportControl.Populate
End Sub

Private Sub txtMileage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

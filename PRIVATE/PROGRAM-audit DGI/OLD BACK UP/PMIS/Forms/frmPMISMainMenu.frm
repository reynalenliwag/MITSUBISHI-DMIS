VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmPMISMainMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PMIS Main Menu"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "frmPMISMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmPMISMainMenu.frx":01CA
   ScaleHeight     =   7035
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11235
      _Version        =   655364
      _ExtentX        =   19817
      _ExtentY        =   12938
      _StockProps     =   64
      Appearance      =   9
      Color           =   4
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   6
      Item(0).Caption =   "Main Modules"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "Picture1"
      Item(1).Caption =   "Tables"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "Picture2"
      Item(2).Caption =   "Inquiry"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "Picture6"
      Item(3).Caption =   "File Maintenance"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "Picture3"
      Item(4).Caption =   "Reports"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "Picture4"
      Item(5).Caption =   "Other Setups"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "Picture5"
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6705
         Left            =   -70045
         ScaleHeight     =   6705
         ScaleWidth      =   11115
         TabIndex        =   106
         Top             =   540
         Visible         =   0   'False
         Width           =   11115
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   6
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":06E5
            Style           =   1  'Graphical
            TabIndex        =   119
            Tag             =   "1364"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   7
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":0D9D
            Style           =   1  'Graphical
            TabIndex        =   118
            Tag             =   "1365"
            Top             =   1207
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   9
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":136E
            Style           =   1  'Graphical
            TabIndex        =   117
            Tag             =   "1373"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   10
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":1877
            Style           =   1  'Graphical
            TabIndex        =   116
            Tag             =   "1374"
            Top             =   1215
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   11
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":1EF1
            Style           =   1  'Graphical
            TabIndex        =   115
            Tag             =   "1376"
            Top             =   2880
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   12
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":253F
            Style           =   1  'Graphical
            TabIndex        =   114
            Tag             =   "1377"
            Top             =   3735
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   13
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":2BD9
            Style           =   1  'Graphical
            TabIndex        =   113
            Tag             =   "1375"
            Top             =   2070
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   14
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":3361
            Style           =   1  'Graphical
            TabIndex        =   112
            Tag             =   "1372"
            Top             =   5445
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   16
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":39CF
            Style           =   1  'Graphical
            TabIndex        =   111
            Tag             =   "1378"
            Top             =   4590
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   18
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":3F87
            Style           =   1  'Graphical
            TabIndex        =   110
            Tag             =   "1371"
            Top             =   4595
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   20
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":45FD
            Style           =   1  'Graphical
            TabIndex        =   109
            Tag             =   "1369"
            Top             =   2901
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   21
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":4D5A
            Style           =   1  'Graphical
            TabIndex        =   108
            Tag             =   "1370"
            Top             =   3748
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   22
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":54A1
            Style           =   1  'Graphical
            TabIndex        =   107
            Tag             =   "1366"
            Top             =   2054
            Width           =   615
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "BROWSE ERROR FILES"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6345
            TabIndex        =   132
            Top             =   4770
            Width           =   3225
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "PO TRANSACTIONS"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   131
            Top             =   3105
            Width           =   3225
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "MRR TRANSACTIONS"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   130
            Top             =   3960
            Width           =   3225
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "ISSUANCES TRANSACTIONS"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   129
            Top             =   4815
            Width           =   3765
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSACTION DETAILS"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1170
            TabIndex        =   128
            Top             =   5625
            Width           =   3225
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "CHECK PREVIOUS BALANCE"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6345
            TabIndex        =   127
            Top             =   540
            Width           =   4260
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "INVENTORY RANKING INQUIRY"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6345
            TabIndex        =   126
            Top             =   1395
            Width           =   4080
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "DEALER SRP / DNP LISTING"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6345
            TabIndex        =   125
            Top             =   2295
            Width           =   4260
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "DEALER / DISTRIBUTOR DNP COMPARISON"
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   6345
            TabIndex        =   124
            Top             =   2925
            Width           =   3480
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "DEALER / DISTRIBUTOR SRP COMPARISON"
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   6345
            TabIndex        =   123
            Top             =   3780
            Width           =   3570
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "PARTS COMPUTERIZED STOCK CARDS"
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   1215
            TabIndex        =   122
            Top             =   2115
            Width           =   2865
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "COUNTER INQUIRY"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   121
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "PARTS AVAILABILITY INQUIRY"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   120
            Top             =   540
            Width           =   3990
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6795
         Left            =   -70000
         ScaleHeight     =   6795
         ScaleWidth      =   11115
         TabIndex        =   6
         Top             =   555
         Visible         =   0   'False
         Width           =   11115
         Begin VB.Label Label64 
            BackStyle       =   0  'Transparent
            Caption         =   "UNDER CONSTRUCTION..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   3330
            TabIndex        =   10
            Top             =   2340
            Width           =   4665
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6705
         Left            =   -70000
         ScaleHeight     =   6705
         ScaleWidth      =   11070
         TabIndex        =   5
         Top             =   570
         Visible         =   0   'False
         Width           =   11070
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   65
            Left            =   3780
            Picture         =   "frmPMISMainMenu.frx":5C4F
            Style           =   1  'Graphical
            TabIndex        =   94
            Tag             =   "1397"
            Top             =   5715
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   64
            Left            =   3780
            Picture         =   "frmPMISMainMenu.frx":631F
            Style           =   1  'Graphical
            TabIndex        =   93
            Tag             =   "1396"
            Top             =   5026
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   63
            Left            =   165
            Picture         =   "frmPMISMainMenu.frx":6A7B
            Style           =   1  'Graphical
            TabIndex        =   65
            Tag             =   "1379"
            Top             =   225
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   62
            Left            =   165
            Picture         =   "frmPMISMainMenu.frx":71F2
            Style           =   1  'Graphical
            TabIndex        =   64
            Tag             =   "1380"
            Top             =   911
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   61
            Left            =   165
            Picture         =   "frmPMISMainMenu.frx":78C7
            Style           =   1  'Graphical
            TabIndex        =   63
            Tag             =   "1381"
            Top             =   1597
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   60
            Left            =   165
            Picture         =   "frmPMISMainMenu.frx":7F0D
            Style           =   1  'Graphical
            TabIndex        =   62
            Tag             =   "1383"
            Top             =   2969
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   59
            Left            =   165
            Picture         =   "frmPMISMainMenu.frx":8615
            Style           =   1  'Graphical
            TabIndex        =   61
            Tag             =   "1382"
            Top             =   2283
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   58
            Left            =   165
            Picture         =   "frmPMISMainMenu.frx":8D15
            Style           =   1  'Graphical
            TabIndex        =   60
            Tag             =   "1384"
            Top             =   3655
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   57
            Left            =   165
            Picture         =   "frmPMISMainMenu.frx":93F1
            Style           =   1  'Graphical
            TabIndex        =   59
            Tag             =   "1387"
            Top             =   5027
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   56
            Left            =   165
            Picture         =   "frmPMISMainMenu.frx":9AE6
            Style           =   1  'Graphical
            TabIndex        =   58
            Tag             =   "1385"
            Top             =   4341
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   55
            Left            =   7455
            Picture         =   "frmPMISMainMenu.frx":A24B
            Style           =   1  'Graphical
            TabIndex        =   57
            Tag             =   "1399"
            Top             =   912
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   54
            Left            =   7455
            Picture         =   "frmPMISMainMenu.frx":A971
            Style           =   1  'Graphical
            TabIndex        =   56
            Tag             =   "1398"
            Top             =   225
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   53
            Left            =   7455
            Picture         =   "frmPMISMainMenu.frx":B083
            Style           =   1  'Graphical
            TabIndex        =   55
            Tag             =   "1400"
            Top             =   1599
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   52
            Left            =   7455
            Picture         =   "frmPMISMainMenu.frx":B859
            Style           =   1  'Graphical
            TabIndex        =   54
            Tag             =   "1402"
            Top             =   2973
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   51
            Left            =   7455
            Picture         =   "frmPMISMainMenu.frx":BF5C
            Style           =   1  'Graphical
            TabIndex        =   53
            Tag             =   "1401"
            Top             =   2286
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   50
            Left            =   150
            Picture         =   "frmPMISMainMenu.frx":C654
            Style           =   1  'Graphical
            TabIndex        =   52
            Tag             =   "1388"
            Top             =   5715
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   49
            Left            =   7455
            Picture         =   "frmPMISMainMenu.frx":CD57
            Style           =   1  'Graphical
            TabIndex        =   51
            Tag             =   "1403"
            Top             =   3660
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   48
            Left            =   3780
            Picture         =   "frmPMISMainMenu.frx":D3C7
            Style           =   1  'Graphical
            TabIndex        =   50
            Tag             =   "1395"
            Top             =   4338
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   47
            Left            =   3780
            Picture         =   "frmPMISMainMenu.frx":DAD5
            Style           =   1  'Graphical
            TabIndex        =   49
            Tag             =   "1389"
            Top             =   210
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   46
            Left            =   3780
            Picture         =   "frmPMISMainMenu.frx":E275
            Style           =   1  'Graphical
            TabIndex        =   48
            Tag             =   "1390"
            Top             =   898
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   45
            Left            =   3780
            Picture         =   "frmPMISMainMenu.frx":E949
            Style           =   1  'Graphical
            TabIndex        =   47
            Tag             =   "1392"
            Top             =   2274
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   44
            Left            =   3780
            Picture         =   "frmPMISMainMenu.frx":F06E
            Style           =   1  'Graphical
            TabIndex        =   46
            Tag             =   "1391"
            Top             =   1586
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   43
            Left            =   3780
            Picture         =   "frmPMISMainMenu.frx":F73E
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "1393"
            Top             =   2962
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   42
            Left            =   7470
            Picture         =   "frmPMISMainMenu.frx":FD93
            Style           =   1  'Graphical
            TabIndex        =   44
            Tag             =   "1404"
            Top             =   4347
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   41
            Left            =   3780
            Picture         =   "frmPMISMainMenu.frx":10528
            Style           =   1  'Graphical
            TabIndex        =   43
            Tag             =   "1394"
            Top             =   3650
            Width           =   615
         End
         Begin VB.Label Label67 
            BackColor       =   &H00FFFFFF&
            Caption         =   "INVENTORY ADJUSTMENT REPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   4500
            TabIndex        =   96
            Top             =   5850
            Width           =   2985
         End
         Begin VB.Label Label66 
            BackColor       =   &H00FFFFFF&
            Caption         =   "STOCKS BELOW SAFETY STOCK LEVEL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   540
            Left            =   4500
            TabIndex        =   95
            Top             =   225
            Width           =   2850
         End
         Begin VB.Label Label65 
            BackColor       =   &H00FFFFFF&
            Caption         =   "STOCK STATUS REPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   870
            TabIndex        =   88
            Top             =   4485
            Width           =   2115
         End
         Begin VB.Label Label63 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ISSUANCES FOR THE MONTH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   915
            TabIndex        =   87
            Top             =   3810
            Width           =   2550
         End
         Begin VB.Label Label62 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TRANSACTION LISTING ISSUANCE REPORT "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   915
            TabIndex        =   86
            Top             =   2355
            Width           =   2535
         End
         Begin VB.Label Label61 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TRANSACTION LISTING RECEIPTS REPORT "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   915
            TabIndex        =   85
            Top             =   1680
            Width           =   1965
         End
         Begin VB.Label Label60 
            BackColor       =   &H00FFFFFF&
            Caption         =   "RECEIPTS FOR THE MONTH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   915
            TabIndex        =   84
            Top             =   3135
            Width           =   2565
         End
         Begin VB.Label Label59 
            BackColor       =   &H00FFFFFF&
            Caption         =   "RIV REPORT FOR WORK-IN PROGRESS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   900
            TabIndex        =   83
            Top             =   960
            Width           =   2595
         End
         Begin VB.Label Label58 
            BackColor       =   &H00FFFFFF&
            Caption         =   "DAILY SALES REPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   900
            TabIndex        =   82
            Top             =   390
            Width           =   2325
         End
         Begin VB.Label Label57 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TOTAL RETAIL SALES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   4485
            TabIndex        =   81
            Top             =   3135
            Width           =   2205
         End
         Begin VB.Label Label56 
            BackColor       =   &H00FFFFFF&
            Caption         =   "UNPOSTED ISSUANCES TRANSACTION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   4485
            TabIndex        =   80
            Top             =   2325
            Width           =   2685
         End
         Begin VB.Label Label55 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SLOW MOVING PARTS FOR DISPOSAL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   4485
            TabIndex        =   79
            Top             =   960
            Width           =   2670
         End
         Begin VB.Label Label54 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TOTAL PURCHASES REPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   4500
            TabIndex        =   78
            Top             =   5220
            Width           =   2580
         End
         Begin VB.Label Label53 
            BackColor       =   &H00FFFFFF&
            Caption         =   "UNPOSTED RECEIPTS TRANSACTION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   480
            Left            =   4485
            TabIndex        =   77
            Top             =   1635
            Width           =   1995
         End
         Begin VB.Label Label52 
            BackColor       =   &H00FFFFFF&
            Caption         =   "INVENTORY RANKING REPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   915
            TabIndex        =   76
            Top             =   5175
            Width           =   2625
         End
         Begin VB.Label Label51 
            BackColor       =   &H00FFFFFF&
            Caption         =   "INVENTORY GROSS RETURN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   8190
            TabIndex        =   75
            Top             =   1080
            Width           =   2475
         End
         Begin VB.Label Label50 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BIR YEAR REPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   8190
            TabIndex        =   74
            Top             =   4545
            Width           =   1785
         End
         Begin VB.Label Label49 
            BackColor       =   &H00FFFFFF&
            Caption         =   "PARTS MOVING AVERAGE DEMAND"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   8190
            TabIndex        =   73
            Top             =   270
            Width           =   2370
         End
         Begin VB.Label Label48 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BEGINNING INVENTORY REPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   4500
            TabIndex        =   72
            Top             =   4500
            Width           =   2865
         End
         Begin VB.Label Label47 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TOTAL COST OF SALES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4500
            TabIndex        =   71
            Top             =   3825
            Width           =   2355
         End
         Begin VB.Label Label44 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ORDERED PARTS REPORT BY CATEGORY"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   8190
            TabIndex        =   70
            Top             =   2340
            Width           =   2430
         End
         Begin VB.Label Label43 
            BackColor       =   &H00FFFFFF&
            Caption         =   "PARTS BACK-ORDER REPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   8190
            TabIndex        =   69
            Top             =   3135
            Width           =   2610
         End
         Begin VB.Label Label41 
            BackColor       =   &H00FFFFFF&
            Caption         =   "MOVEMENT CATEGORY REPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   900
            TabIndex        =   68
            Top             =   5895
            Width           =   2835
         End
         Begin VB.Label Label37 
            BackColor       =   &H00FFFFFF&
            Caption         =   "EXCEL REPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   8190
            TabIndex        =   67
            Top             =   3825
            Width           =   1515
         End
         Begin VB.Label Label30 
            BackColor       =   &H00FFFFFF&
            Caption         =   "FILL RATE REPORTS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   8190
            TabIndex        =   66
            Top             =   1755
            Width           =   2565
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6840
         Left            =   -70015
         ScaleHeight     =   6840
         ScaleWidth      =   11205
         TabIndex        =   4
         Top             =   555
         Visible         =   0   'False
         Width           =   11205
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   39
            Left            =   405
            Picture         =   "frmPMISMainMenu.frx":10C15
            Style           =   1  'Graphical
            TabIndex        =   42
            Tag             =   "1405"
            Top             =   330
            Width           =   900
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   38
            Left            =   405
            Picture         =   "frmPMISMainMenu.frx":1160C
            Style           =   1  'Graphical
            TabIndex        =   41
            Tag             =   "1406"
            Top             =   1395
            Width           =   900
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   36
            Left            =   405
            Picture         =   "frmPMISMainMenu.frx":11E94
            Style           =   1  'Graphical
            TabIndex        =   40
            Tag             =   "1407"
            Top             =   2460
            Width           =   900
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "PASSWORD MAINTENANCE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1530
            TabIndex        =   9
            Top             =   2760
            Width           =   6015
         End
         Begin VB.Label label 
            BackStyle       =   0  'Transparent
            Caption         =   "COMPANY PROFILE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   1530
            TabIndex        =   8
            Top             =   600
            Width           =   3195
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "USER MODULES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1530
            TabIndex        =   7
            Top             =   1680
            Width           =   3675
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6705
         Left            =   -70030
         ScaleHeight     =   6705
         ScaleWidth      =   11115
         TabIndex        =   3
         Top             =   555
         Visible         =   0   'False
         Width           =   11115
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   8
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":127B8
            Style           =   1  'Graphical
            TabIndex        =   92
            Tag             =   "1288"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   17
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":12E85
            Style           =   1  'Graphical
            TabIndex        =   91
            Tag             =   "1287"
            Top             =   1207
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   31
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":134BC
            Style           =   1  'Graphical
            TabIndex        =   90
            Tag             =   "1296"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   32
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":13BE2
            Style           =   1  'Graphical
            TabIndex        =   89
            Tag             =   "1297"
            Top             =   1215
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   35
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":14334
            Style           =   1  'Graphical
            TabIndex        =   26
            Tag             =   "1299"
            Top             =   2880
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   34
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":14A2A
            Style           =   1  'Graphical
            TabIndex        =   25
            Tag             =   "1301"
            Top             =   3735
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   33
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":150A9
            Style           =   1  'Graphical
            TabIndex        =   24
            Tag             =   "1413"
            Top             =   2070
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   30
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":1576C
            Style           =   1  'Graphical
            TabIndex        =   23
            Tag             =   "1295"
            Top             =   5445
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   29
            Left            =   5490
            Picture         =   "frmPMISMainMenu.frx":15E9D
            Style           =   1  'Graphical
            TabIndex        =   22
            Tag             =   "1302"
            Top             =   4590
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   28
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":163CE
            Style           =   1  'Graphical
            TabIndex        =   21
            Tag             =   "1294"
            Top             =   4595
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   27
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":16B2A
            Style           =   1  'Graphical
            TabIndex        =   20
            Tag             =   "1290"
            Top             =   2901
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   26
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":1723A
            Style           =   1  'Graphical
            TabIndex        =   19
            Tag             =   "1291"
            Top             =   3748
            Width           =   615
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   24
            Left            =   360
            Picture         =   "frmPMISMainMenu.frx":17950
            Style           =   1  'Graphical
            TabIndex        =   18
            Tag             =   "1289"
            Top             =   2054
            Width           =   615
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6345
            TabIndex        =   39
            Top             =   4770
            Width           =   3225
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "HARI PARTS MASTER FILE"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   38
            Top             =   3105
            Width           =   3225
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "SUPPLIER MASTER FILE"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   37
            Top             =   3960
            Width           =   3225
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "PARTS MASTER FILE"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   36
            Top             =   4815
            Width           =   3225
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "MATERIALS MASTER FILE"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   35
            Top             =   5670
            Width           =   3225
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "SALESMAN MASTER FILE"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6345
            TabIndex        =   34
            Top             =   540
            Width           =   3225
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "COUNTER MASTER FILE"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6345
            TabIndex        =   33
            Top             =   1395
            Width           =   3225
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "ITEM  PARTS"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6345
            TabIndex        =   32
            Top             =   2295
            Width           =   3225
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PHYSICAL INVENTORY MENU"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6345
            TabIndex        =   31
            Top             =   3060
            Width           =   3465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CREATE INVENTORY DATABASE"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6345
            TabIndex        =   30
            Top             =   3915
            Width           =   3870
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "CUSTOMER MASTER FILE"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   29
            Top             =   2295
            Width           =   3225
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "UPDATE LOCATION"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   28
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "SYSTEM CONFIGURATION"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1215
            TabIndex        =   27
            Top             =   585
            Width           =   3135
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6705
         Left            =   0
         ScaleHeight     =   6705
         ScaleWidth      =   11070
         TabIndex        =   2
         Top             =   540
         Width           =   11070
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   23
            Left            =   240
            Picture         =   "frmPMISMainMenu.frx":17FB7
            Style           =   1  'Graphical
            TabIndex        =   133
            Tag             =   "1467"
            Top             =   360
            Width           =   900
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   0
            Left            =   240
            Picture         =   "frmPMISMainMenu.frx":18724
            Style           =   1  'Graphical
            TabIndex        =   105
            Tag             =   "1304"
            Top             =   1380
            Width           =   900
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   2
            Left            =   225
            Picture         =   "frmPMISMainMenu.frx":18E91
            Style           =   1  'Graphical
            TabIndex        =   104
            Tag             =   "1305"
            Top             =   2340
            Width           =   900
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   3
            Left            =   210
            Picture         =   "frmPMISMainMenu.frx":195B7
            Style           =   1  'Graphical
            TabIndex        =   103
            Tag             =   "1307"
            Top             =   4305
            Width           =   900
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   4
            Left            =   225
            Picture         =   "frmPMISMainMenu.frx":19E20
            Style           =   1  'Graphical
            TabIndex        =   102
            Tag             =   "1306"
            Top             =   3330
            Width           =   900
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   5
            Left            =   4395
            Picture         =   "frmPMISMainMenu.frx":1A69C
            Style           =   1  'Graphical
            TabIndex        =   101
            Tag             =   "1308"
            Top             =   4305
            Width           =   900
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   19
            Left            =   4395
            Picture         =   "frmPMISMainMenu.frx":1AF10
            Style           =   1  'Graphical
            TabIndex        =   100
            Tag             =   "1319"
            Top             =   1320
            Width           =   900
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   15
            Left            =   4395
            Picture         =   "frmPMISMainMenu.frx":1B780
            Style           =   1  'Graphical
            TabIndex        =   99
            Tag             =   "1318"
            Top             =   345
            Width           =   900
         End
         Begin VB.CommandButton cmdAction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   1
            Left            =   4395
            Picture         =   "frmPMISMainMenu.frx":1BF60
            Style           =   1  'Graphical
            TabIndex        =   98
            Tag             =   "1299"
            Top             =   3315
            Width           =   900
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "PARTS REQUISITION ISSUANCE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   1365
            TabIndex        =   134
            Top             =   570
            Width           =   2325
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "INVENTORY RECONCILIATION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5505
            TabIndex        =   97
            Top             =   3645
            Width           =   2955
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "PURCHASE RECEIVING AND STORING"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   5490
            TabIndex        =   17
            Top             =   1500
            Width           =   2370
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "PURCHASE ORDER DATA ENTRY"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   5490
            TabIndex        =   16
            Top             =   645
            Width           =   2775
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "PARTS ISSUANCE        (OVER THE COUNTER)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   1350
            TabIndex        =   15
            Top             =   1560
            Width           =   2325
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "ADVANCE BILL DATA ENTRY"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   1320
            TabIndex        =   14
            Top             =   4620
            Width           =   2655
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "PARTS ISSUANCE        (SERVICE ISSUANCE)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   1320
            TabIndex        =   13
            Top             =   2535
            Width           =   1965
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "DR OUT ISSUANCE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Top             =   3660
            Width           =   1725
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "INVENTORY ADJUSTMENT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5475
            TabIndex        =   11
            Top             =   4650
            Width           =   2955
         End
      End
   End
   Begin VB.Label Label6 
      Caption         =   "FORCE CANCEL OF NON-VAT O.R."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   900
      TabIndex        =   1
      Top             =   4650
      Width           =   4965
   End
End
Attribute VB_Name = "frmPMISMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAction_Click(Index As Integer)
   Select Case cmdAction(Index).Tag
         '***************************************************************************
            ''FILES''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case FILES_SYSTEMCONFIGURATION
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES SYSTEM CONFIGURATION") = False Then Exit Sub
            End If
            'frmPMISSignatories.Show
        Case FILES_UPDATELOCATION
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES UPDATE LOCATION") = False Then Exit Sub
            End If
            frmPMISUpdateLocation.Show
        Case FILES_CUSTOMERMASTERFILE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES MASTER CUSTOMER") = False Then Exit Sub
            End If
            frmALLCustomer.Show
        Case FILES_MASTER_HARIPARTS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES MASTER HARIPARTS") = False Then Exit Sub
            End If
            frmPMISDNPPEntry.Show
        Case FILES_MASTER_SUPPLIER
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES MASTER SUPPLIER") = False Then Exit Sub
            End If
            frmAMISMASTERFILEVendor.Show
        Case FILES_MASTER_PARTS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES MASTER PARTS") = False Then Exit Sub
            End If
            frmPMISParts.Show
        Case FILES_MASTER_MATERIALS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES MASTER MATERIALS") = False Then Exit Sub
            End If
            frmCSMSMaterials.Show
        Case FILES_MASTER_SALESMAN
            frmPMISSalesMan.Show
        Case FILES_MASTER_COUNTER
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES MASTER COUNTER") = False Then Exit Sub
            End If
            frmPMISCounter.Show
        Case FILES_ITEMPARTS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES ITEM PARTS") = False Then Exit Sub
            End If
            frmPMISPartsEntry.Show
        Case FILES_PHYSICALINVENTORY_PHYSICALINVENTORYMENU
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES PHYSICALINVENTORY") = False Then Exit Sub
            End If
            On Error Resume Next
            'frmPMISINVMenu.Show
            frmPMISINVMenuNew.Show
        Case FILES_PHYSICALINVENTORY_CREATEINVENTORYDATABASE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES CREATE INVENTORY DATABASE") = False Then Exit Sub
            End If
            frmPMISCreateINVDATA.Show
        Case FILES_LOCATION
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "FILES LOCATION") = False Then Exit Sub
            End If
            frmPMISLocation.Show
            '***************************************************************************
            ''TRANSACTIONS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case TRANSACTION_PARTSREQUISITION
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION PARTS ISSUANCE OVER THE COUNTER") = False Then Exit Sub
            End If
            frmPMISPrisForms.Show
            
        Case TRANS_PARTSISSUANCEOVERTHECOUNTER, TOOL_CASHISSUANCE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION PARTS ISSUANCE OVER THE COUNTER") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISCustomerOrder
            COUNTERTYPE = "CSH"
            frmPMISCustomerOrder.txtTranType.Text = "CSH"
            frmPMISCustomerOrder.Show
        Case TRANS_CHARGECOUNTERISSUANCE, TOOL_CHARGEISSUANCE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION CHARGE COUNTER ISSUANCE") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISCustomerOrder
            COUNTERTYPE = "CHG"
            frmPMISCustomerOrder.txtTranType.Text = "CHG"
            frmPMISCustomerOrder.Show
        Case TRANS_PARTSISSUANCESERVICEISSUANCE, TOOL_RIVISSUANCE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION PARTS ISSUANCE SERVICE ISSUANCE") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISCustomerOrder
            COUNTERTYPE = "RIV"
            frmPMISCustomerOrder.txtTranType.Text = "RIV"
            frmPMISCustomerOrder.Show
        Case TRANS_DROUTISSUANCE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION DR OUT ISSUANCE") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISCustomerOrder
            COUNTERTYPE = "DR"
            frmPMISCustomerOrder.txtTranType.Text = "DR"
            frmPMISCustomerOrder.Show
        Case TRANS_ADVANCEBILLDATAENTRY
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION ADVANCE BILL DATA ENTRY") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISCustomerOrder
            COUNTERTYPE = "ADB"
            frmPMISCustomerOrder.txtTranType.Text = "ADB"
            frmPMISCustomerOrder.Show
        Case TRANS_INVENTORYADJUSTMENT
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION INVENTORY ADJUSTMENT") = False Then Exit Sub
            End If
            frmPMISInventoryAdjustment.Show
        Case TRANS_ORDERPROCESSING_SERVICE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION ORDER PROCESSING SERVICE") = False Then Exit Sub
            End If
            ORDERTYPE = "S"
           frmMain.LoadPOProc ("Service")
        Case TRANS_ORDERPROCESSING_EMERGENCY
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION ORDER PROCESSING EMERGENCY") = False Then Exit Sub
            End If
            ORDERTYPE = "E"
            frmMain.LoadPOProc ("Emergency")
        Case TRANS_ORDERPROCESSING_COLLISSION
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION ORDER PROCESSING COLLISSION") = False Then Exit Sub
            End If
            ORDERTYPE = "C"
            frmMain.LoadPOProc ("Collision")
        Case TRANS_ORDERPROCESSING_REPLENISHMENT
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION ORDER PROCESSING REPLENISHMENT") = False Then Exit Sub
            End If
            ORDERTYPE = "R"
            frmMain.LoadPOProc ("Replenishment")
            
        Case TRANS_ORDERPROCESSING_OVERTHECOUNTER
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION ORDER PROCESSING OVER THE COUNTER") = False Then Exit Sub
            End If
            ORDERTYPE = "O"
           frmMain.LoadPOProc ("Over The Counter")
        Case TRANS_ORDERPROCESSING_PROMOTIONALSALESCAMPAIGN
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION ORDER PROCESSING PROMOTIONAL SALES CAMPAIGN") = False Then Exit Sub
            End If
            ORDERTYPE = "P"
            frmMain.LoadPOProc ("Promotional Sales Campaign")
        Case TRANS_ORDERPROCESSING_FLEETACCOUNTS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION ORDER PROCESSING FLEET ACCOUNTS") = False Then Exit Sub
            End If
            ORDERTYPE = "F"
            frmMain.LoadPOProc ("Fleet Accounts")
        Case TRANS_ORDERPROCESSING_JTLANCER
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION ORDER PROCESSING JT LANCER") = False Then Exit Sub
            End If
            ORDERTYPE = "J"
            frmMain.LoadPOProc ("JT Lancer")
        Case TRANS_ORDERPROCESSING_TOOLS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION ORDER PROCESSING TOOLS") = False Then Exit Sub
            End If
            ORDERTYPE = "T"
            frmMain.LoadPOProc ("Tools")
        Case TRANS_PURCHASEPROCESSING_PURCHASEORDERDATAENTRY
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION PURCHASE PROCESSING PO DATA ENTRY") = False Then Exit Sub
            End If
            frmPMISPurchase.Show
        Case TRANS_PURCHASEPROCESSING_RECEIVINGSTORING, TOOL_RECEIVINGDATAENTRY

            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION PURCHASE PROCESSING RECEIVING STORING") = False Then Exit Sub
            End If
            frmPMISReceiving2.Show
        Case TRANS_TRAHIST_CASHCOUNTERISSUANCE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION HISTORY CASH COUNTER ISSUANCE") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISCustomerOrderHist
            COUNTERTYPE = "CSH"
            frmPMISCustomerOrderHist.txtTranType.Text = "CSH"
            frmPMISCustomerOrderHist.Show
        Case TRANS_TRAHIST_CHARGECOUNTERISSUANCE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION HISTORY CHARGE COUNTER ISSUANCE") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISCustomerOrderHist
            COUNTERTYPE = "CHG"
            frmPMISCustomerOrderHist.txtTranType.Text = "CHG"
            frmPMISCustomerOrderHist.Show
        Case TRANS_TRAHIST_REQUISTIONISSUANCEVOUCHER
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION HISTORY REQUISTION ISSUANCE VOUCHER") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISCustomerOrderHist
            COUNTERTYPE = "RIV"
            frmPMISCustomerOrderHist.txtTranType.Text = "RIV"
            frmPMISCustomerOrderHist.Show
        Case TRANS_TRAHIST_DROUTISSUANCE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION HISTORY DR OUT ISSUANCE") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISCustomerOrderHist
            COUNTERTYPE = "DR"
            frmPMISCustomerOrderHist.txtTranType.Text = "DR"
            frmPMISCustomerOrderHist.Show
        Case TRANS_TRAHIST_ADVANCEBILLDATAENTRY
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION HISTORY ADVANCE BILL DATA ENTRY") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISCustomerOrderHist
            COUNTERTYPE = "ADB"
            frmPMISCustomerOrderHist.txtTranType.Text = "ADB"
            frmPMISCustomerOrderHist.Show
        Case TRANS_TRAHIST_RECEIVINGSTORING
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "TRANSACTION HISTORY RECEIVING STORING") = False Then Exit Sub
            End If
            On Error Resume Next
            frmPMISReceivingHist.Show
           
            '***************************************************************************
            ''REPORTS'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case REPORT_DAILYSALESREPORT
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS DAILY SALES") = False Then Exit Sub
            End If
            frmPMISDailySales.Show
        Case REPORT_RIVREPORTFORWORKINPROGRESS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS RIV FOR WORKINPROGRESS") = False Then Exit Sub
            End If
            ISSREPTYPE = "RIV_INPROCESS"
            frmPMISIssuances.Show
        Case REPORT_DEALERINTERNAL_TRALSTNG_RECEIPTS, TOOL_RECEIPTSTRANSACTIONLISTING
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS INTERNAL TRASACTION LISTNG RECEIPTS") = False Then Exit Sub
            End If
            frmPMISRCRange.Show
        Case REPORT_DEALERINTERNAL_TRALSTNG_ISSUANCE, TOOL_ISSUANCETRANSACTIONLISTING
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS INTERNAL TRASACTION LISTNG ISSUANCE") = False Then Exit Sub
            End If
            frmPMISISRange.Show
        Case REPORT_DEALERINTERNAL_MEREPORT_RECEIPTSFORTHEMONTH, TOOL_RECEIPTSFORTHEMONTH
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS INTERNAL RECEIPTS FOR THE MONTH") = False Then Exit Sub
            End If
            frmPMISReceipts.Show
        Case REPORT_DEALERINTERNAL_MEREPORT_ISSUANCESFORTHEMONTH, TOOL_ISSUANCESFORTHEMONTH
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS INTERNAL ISSUANCES FOR THE MONTH") = False Then Exit Sub
            End If
            ISSREPTYPE = "ISS_FORTHEMONTH"
            frmPMISIssuances.Show
        Case REPORT_DEALERINTERNAL_MEREPORT_STOCKSTATUSREPORT
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS INTERNAL STOCK STATUS REPORT") = False Then Exit Sub
            End If
            frmPMISPrintStockStat.Show
        Case REPORT_DEALERINTERNAL_INVREPORT_INVENTORYRANKING
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS INTERNAL INVENTORY RANKING") = False Then Exit Sub
            End If
            frmPMISPrintRankfle.Show
        Case REPORT_DEALERINTERNAL_INVREPORT_MOVEMENTCATEGORY
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS INTERNAL MOVEMENT CATEGORY") = False Then Exit Sub
            End If
            frmPMISMoveCat.Show
        Case REPORT_DEALERINTERNAL_INVREPORT_STOCKSBELOWSAFETYSTOCKLEVEL
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS INTERNAL STOCKS BELOW SAFETY STOCK LEVEL") = False Then Exit Sub
            End If
            frmPMISPrintBelowSafetyStock.Show
        Case REPORT_DEALERINTERNAL_INVREPORT_SLOWMOVINGPARTSFORDISPOSAL
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS INTERNAL SLOW MOVING PARTS FOR DISPOSAL") = False Then Exit Sub
            End If
            frmPMISSlowMoving.Show
        Case REPORT_DEALERINTERNAL_UNPOSTED_UNPOSTEDRECEIPTSTRANSACTION
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS INTERNAL UNPOSTED RECEIPTS TRANSACTION") = False Then Exit Sub
            End If
            frmPMISUnPostedRCRange.Show
        Case REPORT_DEALERINTERNAL_UNPOSTED_UNPOSTEDISSUANCESTRANSACTION
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS INTERNAL UNPOSTED ISSUANCES TRANSACTION") = False Then Exit Sub
            End If
            frmPMISUnPostedRange.Show
        Case REPORT_PARTSRUNDOWN_TOTALRETAILSALES
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS TOTAL RETAIL SALES") = False Then Exit Sub
            End If
            PRR_REPORT = "RETAIL SALES"
            frmPMISPRRMonthlyReports.Show
        Case REPORT_PARTSRUNDOWN_TOTALCOSTOFSALES
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS TOTAL COST OF SALES") = False Then Exit Sub
            End If
            PRR_REPORT = "COST OF SALES"
            frmPMISPRRMonthlyReports.Show
            'MsgBox "Reports not Available... this module will be customized " & vbCrLf & _
                    "depending on dealers process flow and implementation procedure...", vbInformation, "Parts Rundown Report..."
        Case REPORT_PARTSRUNDOWN_BEGINNINGINVENTORYREPORT
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS PARTS RUNDOWN BEGINNING INVENTORY REPORT") = False Then Exit Sub
            End If
            PRR_REPORT = "BEGINNING INVENTORY"
            frmPMISPRRMonthlyReports.Show
            'MsgBox "Reports not Available... this module will be customized " & vbCrLf & _
                   "depending on dealers process flow and implementation procedure...", vbInformation, "Parts Rundown Report..."
        Case REPORT_PARTSRUNDOWN_TOTALPURCHASESREPORT
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS PARTS RUNDOWN TOTAL PURCHASES") = False Then Exit Sub
            End If
            PRR_REPORT = "TOTAL PURCHASES"
            frmPMISPRRMonthlyReports.Show
            'MsgBox "Reports not Available... this module will be customized " & vbCrLf & _
                   "depending on dealers process flow and implementation procedure...", vbInformation, "Parts Rundown Report..."
        Case REPORT_PARTSRUNDOWN_INVENTORYADJUSTMENTS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS PARTS RUNDOWN INVENTORY ADJUSTMENTS") = False Then Exit Sub
            End If
            PRR_REPORT = "INVENTORY ADJUSTMENTS"
            frmPMISPRRMonthlyReports.Show
            'MsgBox "Reports not Available... this module will be customized " & vbCrLf & _
                   "depending on dealers process flow and implementation procedure...", vbInformation, "Parts Rundown Report..."
        Case REPORT_PARTSRUNDOWN_PARTSMOVINGAVERAGEDEMAND
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS PARTS RUNDOWN INVENTORY ADJUSTMENTS") = False Then Exit Sub
            End If
            PRR_REPORT = "PARTS MAD"
            frmPMISPRRMonthlyReports.Show
            'MsgBox "Reports not Available... this module will be customized " & vbCrLf & _
                   "depending on dealers process flow and implementation procedure...", vbInformation, "Parts Rundown Report..."
        Case REPORT_PARTSRUNDOWN_INVENTORYGROSSRETURN
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS PARTS RUNDOWN INVENTORY GROSS RETURN") = False Then Exit Sub
            End If
            MsgBox "Reports not Available... this module will be customized " & vbCrLf & _
                   "depending on dealers process flow and implementation procedure...", vbInformation, "Parts Rundown Report..."
        Case REPORT_PARTSRUNDOWN_FILLRATEREPORTS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS PARTS RUNDOWN FILL RATE") = False Then Exit Sub
            End If
            MsgBox "Reports not Available... this module will be customized " & vbCrLf & _
                   "depending on dealers process flow and implementation procedure...", vbInformation, "Parts Rundown Report..."

        Case REPORT_PARTSRUNDOWN_ORDEREDPARTSREPORTBYCATEGORY
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS PARTS RUNDOWN ORDERED PARTS REPORT BY CATEGORY") = False Then Exit Sub
            End If
            MsgBox "Reports not Available... this module will be customized " & vbCrLf & _
                   "depending on dealers process flow and implementation procedure...", vbInformation, "Parts Rundown Report..."
        Case REPORT_PARTSRUNDOWN_PARTSBACKORDERREPORT
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS PARTS RUNDOWN PARTS BACK ORDER") = False Then Exit Sub
            End If
            MsgBox "Reports not Available... this module will be customized " & vbCrLf & _
                   "depending on dealers process flow and implementation procedure...", vbInformation, "Parts Rundown Report..."
        Case REPORT_PARTSRUNDOWN_EXCELREPORT
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORTS PARTS RUNDOWN EXCEL") = False Then Exit Sub
            End If
            FrmPMISRunDown.Show 1
        Case REPORT_GOV_BIRYEARREPORT
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "REPORT GOV BIR YEAR REPORT") = False Then Exit Sub
            End If
            BIR_YearEnd = "PARTS"
            frmPMISBIR_YearEnd.Show

            '***************************************************************************
            ''MAINTENANCE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case MAINTENANCE_COMPANYPROFILE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "MAINTENANCE COMPANY PROFILE") = False Then Exit Sub
            End If
        Case MAINTENANCE_USERMODULES
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "MAINTENANCE USER MODULES") = False Then Exit Sub
            End If
            frmRAM_User.Show
        Case MAINTENANCE_PASSWORDMAINTENANCE
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "MAINTENANCE PASSWORD") = False Then Exit Sub
            End If
            'frmAccMaintenance.Show
        Case MAINTENANCE_TRANSFERDATANOW
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "MAINTENANCE TRANSFER DATA NOW") = False Then Exit Sub
            End If
        Case MAINTENANCE_EXPORTDATANOW
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "MAINTENANCE EXPORT DATA NOW") = False Then Exit Sub
            End If
            '***************************************************************************
            ''INQUIRY''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case INQUIRY_PARTSAVAILABILITYINQUIRY, TOOL_PARTSSRPLOOKUP
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY PARTS AVAILABILITY") = False Then Exit Sub
            End If
            frmPMISPartsInquiry.Show
        Case INQUIRY_COUNTERINQUIRY, TOOL_PARTSCOUNTERINQUIRY
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY COUNTER INQUIRY") = False Then Exit Sub
            End If
            frmPMISCounterInquiry.Show
        Case INQUIRY_PARTSLEDGER_PARTSCOMPUTERIZEDSTOCKCARDS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY PARTS COMPUTERIZED STOCKCARDS") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISQuery
            PARTSQUERY = 1
            frmPMISQuery.Show
        Case INQUIRY_LEDGERBYTRANSACTIONS_POTRANSACTIONS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY PO TRANSACTIONS") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISQuery
            PARTSQUERY = 3
            frmPMISQuery.Show
        Case INQUIRY_LEDGERBYTRANSACTIONS_MRRTRANSACTIONS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY MRR TRANSACTIONS") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISQuery
            PARTSQUERY = 4
            frmPMISQuery.Show
        Case INQUIRY_LEDGERBYTRANSACTIONS_ID_ISSUANCESTRANS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY ISSUANCES TRANSACTIONS") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISQuery
            PARTSQUERY = 5
            frmPMISQuery.Show
        Case INQUIRY_LEDGERBYTRANSACTIONS_TRANSACTIONDETAILS
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY TRANSACTION DETAILS") = False Then Exit Sub
            End If
            On Error Resume Next
            Unload frmPMISQuery
            PARTSQUERY = 7
            frmPMISQuery.Show
        Case INQUIRY_INVBALANCE_CHECKPREVIOUSBAL, TOOL_PARTSPREVIOUSBALANCEINQUIRY
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY CHECK PREVIOUS BALANCE") = False Then Exit Sub
            End If
            frmPMISCheckPrevBal.Show
        Case INQUIRY_INVBALANCE_INVENTORYRANKINGINQUIRY
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY INVENTORY RANKING INQUIRY") = False Then Exit Sub
            End If
            frmPMISRankingInquiry.Show
        Case INQUIRY_DNPSRPCOMPARISON_DEALERSRPDNPLISTING
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY DEALER SRP DNP LISTING") = False Then Exit Sub
            End If
            frmMain.ShowDNPSRPListing
        Case INQUIRY_DNPSRPCOMPARISON_DEALERDISTRIBUTORDNPCOMPARISON
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY DEALER DISTRIBUTOR DNP COMPARISON") = False Then Exit Sub
            End If
            frmPMISPartsDNPComparison.Show
        Case INQUIRY_DNPSRPCOMPARISON_DEALERDISTRIBUTORSRPCOMPARISON
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY DEALER DISTRIBUTOR SRP COMPARISON") = False Then Exit Sub
            End If
            frmPMISPartsSRPComparison.Show
        Case INQUIRY_OTHERQUERIES_BROWSEERRORFILES
            If ApplySecurityValidation = True Then
                If Module_Access(LOGID, "INQUIRY BROWSE ERROR FILES") = False Then Exit Sub
            End If
            frmPMISErrorQuery.Show
    End Select
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    TabControl1.SelectedItem = 0
End Sub

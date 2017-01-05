VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCSMS_CQI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QUALITY INFORMATION (PWA)"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_CQI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   12345
   Begin VB.PictureBox picAdds 
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
      Height          =   960
      Left            =   360
      ScaleHeight     =   960
      ScaleWidth      =   12015
      TabIndex        =   119
      Top             =   7770
      Width           =   12015
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   11220
         MouseIcon       =   "frmCSMS_CQI.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Exit Window"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   10530
         MouseIcon       =   "frmCSMS_CQI.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Print this Record"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost"
         Height          =   795
         Left            =   9840
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmCSMS_CQI.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Unpost this Transaction"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post"
         Height          =   795
         Left            =   9150
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmCSMS_CQI.frx":1E89
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":1FDB
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Post this Transaction"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   900
         MouseIcon       =   "frmCSMS_CQI.frx":2300
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":2452
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Delete Selected Record"
         Top             =   45
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdCancelCO 
         Caption         =   "Cancel"
         Height          =   795
         Left            =   8460
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmCSMS_CQI.frx":277D
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":28CF
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Cancel this Transaction"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   7770
         MouseIcon       =   "frmCSMS_CQI.frx":2C09
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":2D5B
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Edit Selected Record"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   7080
         MouseIcon       =   "frmCSMS_CQI.frx":30B7
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":3209
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Add Record"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   795
         Left            =   6360
         MouseIcon       =   "frmCSMS_CQI.frx":351C
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":366E
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Move to Last Record"
         Top             =   45
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   795
         Left            =   5640
         MouseIcon       =   "frmCSMS_CQI.frx":39BE
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":3B10
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Move to First Record"
         Top             =   45
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   4950
         MouseIcon       =   "frmCSMS_CQI.frx":3E6E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":3FC0
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Find a Record"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   4260
         MouseIcon       =   "frmCSMS_CQI.frx":42BA
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":440C
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Move to Next Record"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   3570
         MouseIcon       =   "frmCSMS_CQI.frx":4764
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":48B6
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Move to Previous Record"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2880
         MouseIcon       =   "frmCSMS_CQI.frx":4C15
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":4D67
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Refresh active transaction"
         Top             =   45
         Width           =   705
      End
      Begin VB.Label lblSTATUS 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   166
         Top             =   120
         Width           =   3435
      End
   End
   Begin VB.PictureBox Picture31 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   5175
      Left            =   12420
      ScaleHeight     =   5145
      ScaleWidth      =   3315
      TabIndex        =   260
      Top             =   150
      Width           =   3345
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
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
         Height          =   285
         Index           =   10
         Left            =   30
         ScaleHeight     =   255
         ScaleWidth      =   5265
         TabIndex        =   281
         Top             =   4620
         Width           =   5295
         Begin VB.Label lblcap 
            BackColor       =   &H00FCE2CF&
            Caption         =   " ACCESSORIES"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   26
            Left            =   0
            TabIndex        =   282
            Top             =   0
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
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
         Height          =   255
         Index           =   11
         Left            =   30
         ScaleHeight     =   225
         ScaleWidth      =   5265
         TabIndex        =   279
         Top             =   4347
         Width           =   5295
         Begin VB.Label lblcap 
            BackColor       =   &H00FCE2CF&
            Caption         =   " OCCURENCE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   27
            Left            =   0
            MouseIcon       =   "frmCSMS_CQI.frx":52E2
            MousePointer    =   99  'Custom
            TabIndex        =   280
            Top             =   0
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
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
         Height          =   555
         Index           =   7
         Left            =   30
         ScaleHeight     =   525
         ScaleWidth      =   5265
         TabIndex        =   277
         Top             =   3774
         Width           =   5295
         Begin VB.Label lblcap 
            BackColor       =   &H00FCE2CF&
            Caption         =   " ACTION"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   28
            Left            =   0
            MouseIcon       =   "frmCSMS_CQI.frx":5434
            MousePointer    =   99  'Custom
            TabIndex        =   278
            Top             =   0
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
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
         Height          =   525
         Index           =   8
         Left            =   30
         ScaleHeight     =   495
         ScaleWidth      =   5265
         TabIndex        =   275
         Top             =   3231
         Width           =   5295
         Begin VB.Label lblcap 
            BackColor       =   &H00FCE2CF&
            Caption         =   " LOCATION"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   29
            Left            =   0
            MouseIcon       =   "frmCSMS_CQI.frx":5586
            MousePointer    =   99  'Custom
            TabIndex        =   276
            Top             =   15
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
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
         Height          =   555
         Index           =   9
         Left            =   30
         ScaleHeight     =   525
         ScaleWidth      =   5265
         TabIndex        =   273
         Top             =   2658
         Width           =   5295
         Begin VB.Label lblcap 
            BackColor       =   &H00FCE2CF&
            Caption         =   " ROAD"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   31
            Left            =   0
            TabIndex        =   274
            Top             =   0
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
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
         Height          =   555
         Index           =   2
         Left            =   30
         ScaleHeight     =   525
         ScaleWidth      =   5265
         TabIndex        =   271
         Top             =   2085
         Width           =   5295
         Begin VB.Label lblcap 
            BackColor       =   &H00FCE2CF&
            Caption         =   " A/T"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   30
            Left            =   0
            TabIndex        =   272
            Top             =   15
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
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
         Height          =   555
         Index           =   3
         Left            =   30
         ScaleHeight     =   525
         ScaleWidth      =   5265
         TabIndex        =   269
         Top             =   1512
         Width           =   5295
         Begin VB.Label lblcap 
            BackColor       =   &H00FCE2CF&
            Caption         =   " M/T"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   35
            Left            =   0
            TabIndex        =   270
            Top             =   0
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
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
         Height          =   285
         Index           =   4
         Left            =   30
         ScaleHeight     =   255
         ScaleWidth      =   5265
         TabIndex        =   267
         Top             =   1209
         Width           =   5295
         Begin VB.Label lblcap 
            BackColor       =   &H00FCE2CF&
            Caption         =   " SHIFT POSITION"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   34
            Left            =   0
            TabIndex        =   268
            Top             =   0
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
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
         Height          =   285
         Index           =   5
         Left            =   30
         ScaleHeight     =   255
         ScaleWidth      =   5265
         TabIndex        =   265
         Top             =   906
         Width           =   5295
         Begin VB.Label lblcap 
            BackColor       =   &H00FCE2CF&
            Caption         =   " SHIFTING"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   33
            Left            =   0
            TabIndex        =   266
            Top             =   15
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
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
         Height          =   285
         Index           =   6
         Left            =   30
         ScaleHeight     =   255
         ScaleWidth      =   5265
         TabIndex        =   263
         Top             =   603
         Width           =   5295
         Begin VB.Label lblcap 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCE2CF&
            Caption         =   " WEATHER"
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
            Height          =   270
            Index           =   32
            Left            =   0
            TabIndex        =   264
            Top             =   0
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
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
         Height          =   285
         Index           =   0
         Left            =   30
         ScaleHeight     =   255
         ScaleWidth      =   5265
         TabIndex        =   261
         Top             =   300
         Width           =   5295
         Begin VB.Label lblcap 
            BackColor       =   &H00FCE2CF&
            Caption         =   " ENGINE TEMP"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   36
            Left            =   0
            TabIndex        =   262
            Top             =   0
            Width           =   1260
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   283
         Top             =   0
         Width           =   6375
         _Version        =   655364
         _ExtentX        =   11245
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "CONDITION"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   16711680
         ForeColor       =   16777215
      End
   End
   Begin VB.PictureBox picPRINT 
      Appearance      =   0  'Flat
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
      Height          =   2265
      Left            =   12420
      ScaleHeight     =   2235
      ScaleWidth      =   2055
      TabIndex        =   158
      Top             =   6090
      Visible         =   0   'False
      Width           =   2085
      Begin VB.OptionButton Option5 
         Caption         =   "VIEW REPORT"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   1770
         Width           =   1875
      End
      Begin VB.OptionButton Option4 
         Caption         =   "EXCEL REPORT"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   1440
         Value           =   -1  'True
         Width           =   1875
      End
      Begin wizButton.cmd cmd4 
         Height          =   345
         Left            =   60
         TabIndex        =   159
         Top             =   300
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   609
         TX              =   "ACL"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCSMS_CQI.frx":56D8
      End
      Begin wizButton.cmd cmd5 
         Height          =   345
         Left            =   60
         TabIndex        =   161
         Top             =   690
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   609
         TX              =   "CLAIM"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCSMS_CQI.frx":56F4
      End
      Begin wizButton.cmd cmd6 
         Height          =   225
         Left            =   1830
         TabIndex        =   165
         Top             =   0
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   397
         TX              =   "x"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCSMS_CQI.frx":5710
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption11 
         Height          =   225
         Index           =   2
         Left            =   0
         TabIndex        =   164
         Top             =   1110
         Width           =   2835
         _Version        =   655364
         _ExtentX        =   5001
         _ExtentY        =   397
         _StockProps     =   14
         Caption         =   "PRINT TYPE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   16711680
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption11 
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   160
         Top             =   0
         Width           =   2835
         _Version        =   655364
         _ExtentX        =   5001
         _ExtentY        =   397
         _StockProps     =   14
         Caption         =   "PRINT OPTION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   16711680
      End
   End
   Begin VB.PictureBox picMENU 
      Appearance      =   0  'Flat
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
      Height          =   1515
      Left            =   60
      ScaleHeight     =   1485
      ScaleWidth      =   2205
      TabIndex        =   153
      Top             =   6240
      Width           =   2235
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1770
         Top             =   30
      End
      Begin VB.Label Label3 
         Caption         =   "SHIFT F1 - VIEW AUDIT TRAIL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   30
         TabIndex        =   327
         Top             =   1110
         Width           =   2295
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption11 
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   157
         Top             =   0
         Width           =   2835
         _Version        =   655364
         _ExtentX        =   5001
         _ExtentY        =   397
         _StockProps     =   14
         Caption         =   "SHORTCUTS"
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
         ForeColor       =   8388608
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F12 - DISAPPROVED QIR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   309
         Top             =   900
         Width           =   1890
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F9 - APPROVED QIR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   308
         Top             =   660
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F3 - ADD PARTS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   60
         TabIndex        =   156
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F4 - TO ADD JOBS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1770
         TabIndex        =   155
         Top             =   570
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F4 - ADD SUBLET/JOBS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   154
         Top             =   450
         Width           =   1755
      End
   End
   Begin VB.PictureBox picSaves 
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
      Height          =   885
      Left            =   10890
      ScaleHeight     =   885
      ScaleWidth      =   1590
      TabIndex        =   120
      Top             =   7770
      Width           =   1590
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   690
         MouseIcon       =   "frmCSMS_CQI.frx":572C
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":587E
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Cancel"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   0
         MouseIcon       =   "frmCSMS_CQI.frx":5BBC
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":5D0E
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Save this Record"
         Top             =   45
         Width           =   705
      End
   End
   Begin VB.PictureBox picSEARCH 
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
      Height          =   6165
      Left            =   60
      ScaleHeight     =   6135
      ScaleWidth      =   2205
      TabIndex        =   116
      Top             =   60
      Width           =   2235
      Begin VB.ComboBox cboSearchBy 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCSMS_CQI.frx":605E
         Left            =   30
         List            =   "frmCSMS_CQI.frx":606E
         Style           =   2  'Dropdown List
         TabIndex        =   339
         Top             =   270
         Width           =   2145
      End
      Begin VB.TextBox txtSEARCH 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   118
         Top             =   630
         Width           =   2145
      End
      Begin MSComctlLib.ListView lsvDLR 
         Height          =   5085
         Left            =   30
         TabIndex        =   117
         Top             =   990
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   8969
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
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "QIR NO."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   123
         Top             =   0
         Width           =   2205
         _Version        =   655364
         _ExtentX        =   3889
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "SEARCH"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   8388608
      End
   End
   Begin TabDlg.SSTab TabControl1 
      Height          =   7695
      Left            =   2310
      TabIndex        =   167
      Top             =   60
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   13573
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "INFORMATION"
      TabPicture(0)   =   "frmCSMS_CQI.frx":60BB
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picHD"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picDET"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "NOTES"
      TabPicture(1)   =   "frmCSMS_CQI.frx":60D7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picNOTES"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CONDITIONS"
      TabPicture(2)   =   "frmCSMS_CQI.frx":60F3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picCONDITIONS"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.PictureBox picDET 
         Appearance      =   0  'Flat
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
         Height          =   2145
         Left            =   0
         ScaleHeight     =   2115
         ScaleWidth      =   9975
         TabIndex        =   210
         Top             =   5550
         Width           =   10005
         Begin VB.PictureBox picPARTS 
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
            Height          =   2205
            Left            =   30
            ScaleHeight     =   2205
            ScaleWidth      =   9945
            TabIndex        =   211
            Top             =   30
            Width           =   9945
            Begin VB.TextBox txtTCOST 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   360
               Left            =   5160
               TabIndex        =   213
               Text            =   "0.00"
               Top             =   1680
               Width           =   1665
            End
            Begin VB.TextBox txtLTS 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   360
               Left            =   8310
               TabIndex        =   212
               Text            =   "0.00"
               Top             =   1680
               Width           =   885
            End
            Begin MSComctlLib.ListView lsvPARTS 
               Height          =   1665
               Left            =   0
               TabIndex        =   214
               Top             =   0
               Width           =   9915
               _ExtentX        =   17489
               _ExtentY        =   2937
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
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "PART No"
                  Object.Width           =   3175
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "DESCRIPTION"
                  Object.Width           =   4410
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Text            =   "QTY"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "COST"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   4
                  Text            =   "OP CODE"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "LTS"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "TYPE"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "ID"
                  Object.Width           =   0
               EndProperty
            End
            Begin VB.Label lblcap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "TOTAL LTS"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   93
               Left            =   7380
               TabIndex        =   338
               Top             =   1740
               Width           =   780
            End
            Begin VB.Label lblFLATRATE 
               Alignment       =   2  'Center
               Caption         =   "247.5"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Left            =   1200
               TabIndex        =   337
               Top             =   1740
               Width           =   675
            End
            Begin VB.Label lblcap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "TOTAL PARTS"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   86
               Left            =   4035
               TabIndex        =   317
               Top             =   1740
               Width           =   1005
            End
            Begin VB.Label lblcap 
               AutoSize        =   -1  'True
               Caption         =   " (Labor Cost) :                   * Total LTS"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   80
               Left            =   30
               TabIndex        =   307
               Top             =   1740
               Width           =   2730
            End
         End
      End
      Begin VB.PictureBox picHD 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   5235
         Left            =   0
         ScaleHeight     =   5205
         ScaleWidth      =   9975
         TabIndex        =   176
         Top             =   300
         Width           =   10005
         Begin VB.TextBox txtCASUAL 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1710
            TabIndex        =   19
            ToolTipText     =   "Press Enter to search for causal partno"
            Top             =   3720
            Width           =   2025
         End
         Begin VB.TextBox txtFLATRATE 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   5100
            TabIndex        =   335
            Text            =   "247.5"
            Top             =   2430
            Width           =   945
         End
         Begin VB.TextBox txtCDESC 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            MaxLength       =   250
            TabIndex        =   25
            Top             =   4440
            Width           =   1245
         End
         Begin VB.TextBox txtNDESC 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            MaxLength       =   250
            TabIndex        =   24
            Top             =   4110
            Width           =   1245
         End
         Begin VB.ComboBox cboCCODE 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1710
            TabIndex        =   23
            Text            =   "cboCCODE"
            Top             =   4440
            Width           =   825
         End
         Begin VB.ComboBox cboNCODE 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1710
            TabIndex        =   22
            Text            =   "cboNCODE"
            Top             =   4110
            Width           =   825
         End
         Begin VB.ComboBox cboCType 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmCSMS_CQI.frx":610F
            Left            =   4560
            List            =   "frmCSMS_CQI.frx":6122
            Style           =   2  'Dropdown List
            TabIndex        =   330
            Top             =   270
            Width           =   1425
         End
         Begin VB.TextBox txtPREVACL 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            MaxLength       =   4
            TabIndex        =   30
            Top             =   4800
            Width           =   2115
         End
         Begin VB.TextBox txtPREVRO 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5220
            MaxLength       =   4
            TabIndex        =   31
            Top             =   4770
            Width           =   1965
         End
         Begin VB.TextBox txtSubletType 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4500
            TabIndex        =   21
            Top             =   3750
            Width           =   1545
         End
         Begin wizButton.cmd cmd8 
            Height          =   285
            Left            =   9510
            TabIndex        =   325
            ToolTipText     =   "search plate no"
            Top             =   2310
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   503
            TX              =   ". . ."
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "frmCSMS_CQI.frx":6135
         End
         Begin VB.TextBox txtModel 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   4140
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   3000
            Width           =   1905
         End
         Begin VB.PictureBox Picture2 
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
            Height          =   405
            Left            =   1710
            ScaleHeight     =   405
            ScaleWidth      =   1605
            TabIndex        =   231
            Top             =   270
            Width           =   1605
            Begin VB.OptionButton optPWAREQ2 
               Caption         =   "NO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   780
               Style           =   1  'Graphical
               TabIndex        =   1
               Top             =   0
               Width           =   765
            End
            Begin VB.OptionButton optPWAREQ1 
               Caption         =   "YES"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   0
               Top             =   0
               Value           =   -1  'True
               Width           =   765
            End
         End
         Begin VB.TextBox txtPLATENO 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7590
            TabIndex        =   37
            Top             =   2310
            Width           =   1845
         End
         Begin VB.TextBox TXTGRAND 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   7590
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   4080
            Width           =   2295
         End
         Begin VB.TextBox TXTSUBREP 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   7590
            TabIndex        =   41
            Text            =   "0.00"
            Top             =   3780
            Width           =   2295
         End
         Begin VB.TextBox txtLABORCOST 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   7590
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "0.00"
            Top             =   3450
            Width           =   2295
         End
         Begin VB.TextBox txtINV 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   7590
            TabIndex        =   43
            Top             =   4410
            Width           =   2295
         End
         Begin VB.TextBox txtPCODE 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5220
            MaxLength       =   4
            TabIndex        =   28
            Top             =   4110
            Width           =   825
         End
         Begin VB.TextBox txtSCODE 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5220
            MaxLength       =   4
            TabIndex        =   29
            Top             =   4440
            Width           =   825
         End
         Begin VB.TextBox txtCCODE 
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            MaxLength       =   4
            TabIndex        =   27
            Top             =   4290
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtCASUAL1 
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   20
            Top             =   3570
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.PictureBox Picture7 
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
            Height          =   435
            Left            =   1710
            ScaleHeight     =   435
            ScaleWidth      =   2325
            TabIndex        =   180
            Top             =   3270
            Width           =   2325
            Begin VB.OptionButton optSENT1 
               Caption         =   "YES"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   0
               Width           =   1125
            End
            Begin VB.OptionButton optSENT2 
               Caption         =   "NO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   1110
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   0
               Width           =   1125
            End
         End
         Begin VB.PictureBox Picture6 
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
            Height          =   435
            Left            =   1710
            ScaleHeight     =   435
            ScaleWidth      =   2295
            TabIndex        =   179
            Top             =   2820
            Width           =   2295
            Begin VB.OptionButton optATT1 
               Caption         =   "PHOTO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   0
               Width           =   1125
            End
            Begin VB.OptionButton optATT2 
               Caption         =   "SAMPLE PART"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   1110
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   0
               Width           =   1125
            End
         End
         Begin VB.PictureBox Picture4 
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
            Height          =   405
            Left            =   1710
            ScaleHeight     =   405
            ScaleWidth      =   2325
            TabIndex        =   178
            Top             =   2400
            Width           =   2325
            Begin VB.OptionButton optTRAN1 
               Caption         =   "MANUAL"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   0
               Width           =   1125
            End
            Begin VB.OptionButton optTRAN2 
               Caption         =   "AUTO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   1110
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   0
               Width           =   1125
            End
         End
         Begin VB.ComboBox cboTECH 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7590
            TabIndex        =   39
            Text            =   "Combo1"
            Top             =   3060
            Width           =   2310
         End
         Begin VB.ComboBox cboSA 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7590
            TabIndex        =   38
            Text            =   "Combo1"
            Top             =   2670
            Width           =   2310
         End
         Begin VB.TextBox txtKM 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7590
            TabIndex        =   36
            Top             =   1980
            Width           =   1845
         End
         Begin VB.TextBox txtPWAT 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4560
            TabIndex        =   4
            Top             =   690
            Width           =   1425
         End
         Begin VB.TextBox txtCUST 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            TabIndex        =   8
            Top             =   1380
            Width           =   3915
         End
         Begin VB.TextBox txtDLR 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7590
            TabIndex        =   5
            Text            =   " "
            Top             =   270
            Width           =   2295
         End
         Begin VB.TextBox txtDCODE 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7590
            TabIndex        =   32
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txtENGINE 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4020
            MaxLength       =   25
            TabIndex        =   10
            Top             =   1740
            Width           =   1995
         End
         Begin VB.TextBox txtAXLE 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            TabIndex        =   11
            Top             =   2070
            Width           =   4305
         End
         Begin VB.TextBox txtPWA 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   690
            Width           =   1545
         End
         Begin VB.TextBox txtVIN 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1650
            MaxLength       =   25
            TabIndex        =   9
            Top             =   1740
            Width           =   2265
         End
         Begin wizButton.cmd cmd1 
            Height          =   315
            Left            =   9510
            TabIndex        =   177
            ToolTipText     =   "search a reapir order"
            Top             =   1650
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            TX              =   "..."
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "frmCSMS_CQI.frx":6151
         End
         Begin MSComCtl2.DTPicker dptDELDATTE 
            Height          =   315
            Left            =   4320
            TabIndex        =   6
            Top             =   1020
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   54394881
            CurrentDate     =   39562
         End
         Begin MSComCtl2.DTPicker dptINSD 
            Height          =   315
            Left            =   7590
            TabIndex        =   33
            Top             =   930
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   54394881
            CurrentDate     =   39562
         End
         Begin MSComCtl2.DTPicker dptREPD 
            Height          =   315
            Left            =   7590
            TabIndex        =   34
            Top             =   1290
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   54394881
            CurrentDate     =   39562
         End
         Begin wizButton.cmd cmd2 
            Height          =   285
            Left            =   5670
            TabIndex        =   181
            Top             =   1380
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   503
            TX              =   "..."
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "frmCSMS_CQI.frx":616D
         End
         Begin MSComCtl2.DTPicker dtpTranDate 
            Height          =   315
            Left            =   1710
            TabIndex        =   7
            Top             =   1020
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   54394881
            CurrentDate     =   39562
         End
         Begin VB.TextBox txtRO 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7590
            MaxLength       =   10
            TabIndex        =   35
            Top             =   1650
            Width           =   1845
         End
         Begin VB.TextBox txtNCODE 
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            MaxLength       =   4
            TabIndex        =   26
            Top             =   4020
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "FLAT RATE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   92
            Left            =   4140
            TabIndex        =   336
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "PREV. RO NO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   90
            Left            =   4215
            TabIndex        =   2
            Top             =   4860
            Width           =   975
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "PREV. ACL NO. OF RESUBMIT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Index           =   89
            Left            =   240
            TabIndex        =   329
            Top             =   4800
            Width           =   1440
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "SUBLET TYPE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   88
            Left            =   3840
            TabIndex        =   328
            Top             =   3720
            Width           =   630
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "TRAN. DATE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   87
            Left            =   780
            TabIndex        =   326
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CLAIM TYPE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   85
            Left            =   3615
            TabIndex        =   315
            Top             =   390
            Width           =   870
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "MODEL:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   84
            Left            =   4170
            TabIndex        =   314
            Top             =   2820
            Width           =   570
         End
         Begin VB.Label lblTRANNO 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   7740
            TabIndex        =   313
            Top             =   0
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PWA REQUEST"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   37
            Left            =   600
            TabIndex        =   230
            Top             =   390
            Width           =   1080
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "INVOICE #"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   52
            Left            =   6285
            TabIndex        =   209
            Top             =   4500
            Width           =   1290
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "LABOR COST"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   51
            Left            =   5820
            TabIndex        =   208
            Top             =   3480
            Width           =   1755
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   " SUBLET COST"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   50
            Left            =   5595
            TabIndex        =   207
            Top             =   3810
            Width           =   1980
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "GRAND TOTAL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   49
            Left            =   6330
            TabIndex        =   206
            Top             =   4140
            Width           =   1245
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "SUBLET CODE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   48
            Left            =   4185
            TabIndex        =   205
            Top             =   4500
            Width           =   1005
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "CAUSAL PART NO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   47
            Left            =   390
            TabIndex        =   204
            Top             =   3810
            Width           =   1290
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "NATURE CODE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   46
            Left            =   615
            TabIndex        =   203
            Top             =   4230
            Width           =   1065
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "CAUSE CODE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   45
            Left            =   720
            TabIndex        =   202
            Top             =   4530
            Width           =   960
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "PAINT CODE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   44
            Left            =   4275
            TabIndex        =   201
            Top             =   4170
            Width           =   915
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "ATTACHMENTS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   585
            TabIndex        =   200
            Top             =   2850
            Width           =   1095
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "SENT REGISTRATION CARD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   810
            Index           =   18
            Left            =   465
            TabIndex        =   199
            Top             =   3150
            Width           =   1215
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "PLATE NO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   6735
            TabIndex        =   198
            Top             =   2370
            Width           =   840
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "MILEAGE (KM)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   6345
            TabIndex        =   197
            Top             =   2010
            Width           =   1230
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "DLR QIR REF. NO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   5775
            TabIndex        =   196
            Top             =   285
            Width           =   1800
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "TECHNICIAN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   6465
            TabIndex        =   195
            Top             =   3150
            Width           =   1110
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "SA NAME"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   5940
            TabIndex        =   194
            Top             =   2760
            Width           =   1635
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "REPAIR ORDER"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   6255
            TabIndex        =   193
            Top             =   1710
            Width           =   1320
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "REPAIR DATE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   6405
            TabIndex        =   192
            Top             =   1380
            Width           =   1170
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "INSP. DATE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   6570
            TabIndex        =   191
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "DEALER CODE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   6330
            TabIndex        =   190
            Top             =   660
            Width           =   1245
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "VIN /ENGINE#"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   570
            TabIndex        =   189
            Top             =   1770
            Width           =   1035
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "TM/AXLE NO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   780
            TabIndex        =   188
            Top             =   2130
            Width           =   900
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "TRANSMISSION TYPE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Index           =   4
            Left            =   -15
            TabIndex        =   187
            Top             =   2460
            Width           =   1695
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "CUSTOMER"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   855
            TabIndex        =   186
            Top             =   1440
            Width           =   825
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "DEL. DATE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   3480
            TabIndex        =   185
            Top             =   1110
            Width           =   765
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "PWA TYPE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   3720
            TabIndex        =   184
            Top             =   735
            Width           =   750
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "PWA NO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   1065
            TabIndex        =   183
            Top             =   735
            Width           =   615
         End
         Begin VB.Label LABID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   8280
            TabIndex        =   182
            Top             =   0
            Width           =   1575
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption13 
            Height          =   285
            Left            =   0
            TabIndex        =   340
            Top             =   -30
            Width           =   9945
            _Version        =   655364
            _ExtentX        =   17542
            _ExtentY        =   503
            _StockProps     =   14
            Caption         =   "INFORMATION"
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
            ForeColor       =   8388608
         End
      End
      Begin VB.PictureBox picCONDITIONS 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   7395
         Left            =   -75000
         ScaleHeight     =   7365
         ScaleWidth      =   9975
         TabIndex        =   215
         Top             =   300
         Width           =   10005
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   4965
            Left            =   0
            ScaleHeight     =   4935
            ScaleWidth      =   5355
            TabIndex        =   284
            Top             =   0
            Width           =   5385
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
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
               Height          =   285
               Index           =   1
               Left            =   30
               ScaleHeight     =   255
               ScaleWidth      =   5265
               TabIndex        =   305
               Top             =   4530
               Width           =   5295
               Begin VB.CheckBox optACC 
                  Caption         =   "A/C ON"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   90
                  Top             =   0
                  Width           =   1185
               End
               Begin VB.CheckBox optACC 
                  Caption         =   "HEAVY ELECTRIC LOAD"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   2550
                  TabIndex        =   91
                  Top             =   0
                  Width           =   2505
               End
               Begin VB.Label lblcap 
                  BackColor       =   &H00FCE2CF&
                  Caption         =   " ACC."
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   69
                  Left            =   0
                  TabIndex        =   306
                  Top             =   0
                  Width           =   1380
               End
            End
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
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
               Height          =   285
               Index           =   12
               Left            =   30
               ScaleHeight     =   255
               ScaleWidth      =   5265
               TabIndex        =   303
               Top             =   300
               Width           =   5295
               Begin VB.CheckBox optENG 
                  Caption         =   "HOT"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   51
                  Top             =   -30
                  Width           =   1155
               End
               Begin VB.CheckBox optENG 
                  Caption         =   "COLD"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   1
                  Left            =   2550
                  TabIndex        =   52
                  Top             =   -30
                  Width           =   1155
               End
               Begin VB.CheckBox optENG 
                  Caption         =   "ALL TEMP"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   2
                  Left            =   3810
                  TabIndex        =   53
                  Top             =   -30
                  Width           =   1155
               End
               Begin VB.Label lblcap 
                  BackColor       =   &H00FCE2CF&
                  Caption         =   " ENGINE TEMP"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   70
                  Left            =   0
                  TabIndex        =   304
                  Top             =   0
                  Width           =   1260
               End
            End
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
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
               Height          =   285
               Index           =   13
               Left            =   30
               ScaleHeight     =   255
               ScaleWidth      =   5265
               TabIndex        =   301
               Top             =   600
               Width           =   5295
               Begin VB.CheckBox optWEA 
                  Caption         =   "WARM"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   54
                  Top             =   0
                  Width           =   1155
               End
               Begin VB.CheckBox optWEA 
                  Caption         =   "COLD"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   2550
                  TabIndex        =   55
                  Top             =   0
                  Width           =   1155
               End
               Begin VB.CheckBox optWEA 
                  Caption         =   "ALL TEMP"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   2
                  Left            =   3810
                  TabIndex        =   56
                  Top             =   0
                  Width           =   1155
               End
               Begin VB.Label lblcap 
                  BackColor       =   &H00FCE2CF&
                  Caption         =   " WEATHER"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   71
                  Left            =   0
                  TabIndex        =   302
                  Top             =   0
                  Width           =   1260
               End
            End
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
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
               Height          =   285
               Index           =   14
               Left            =   30
               ScaleHeight     =   255
               ScaleWidth      =   5265
               TabIndex        =   299
               Top             =   900
               Width           =   5295
               Begin VB.CheckBox optSHIF 
                  Caption         =   "NORMAL"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   57
                  Top             =   0
                  Width           =   1245
               End
               Begin VB.CheckBox optSHIF 
                  Caption         =   "FAST"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   2550
                  TabIndex        =   58
                  Top             =   0
                  Width           =   1245
               End
               Begin VB.Label lblcap 
                  BackColor       =   &H00FCE2CF&
                  Caption         =   " SHIFTING"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   72
                  Left            =   0
                  TabIndex        =   300
                  Top             =   0
                  Width           =   1260
               End
            End
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
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
               Height          =   285
               Index           =   15
               Left            =   30
               ScaleHeight     =   255
               ScaleWidth      =   5265
               TabIndex        =   297
               Top             =   1200
               Width           =   5295
               Begin VB.CheckBox optPOS 
                  Caption         =   "4WD"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   59
                  Top             =   0
                  Width           =   1125
               End
               Begin VB.CheckBox optPOS 
                  Caption         =   "2WD"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   2550
                  TabIndex        =   60
                  Top             =   0
                  Width           =   1125
               End
               Begin VB.Label lblcap 
                  BackColor       =   &H00FCE2CF&
                  Caption         =   " SHIFT POS."
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   73
                  Left            =   0
                  TabIndex        =   298
                  Top             =   0
                  Width           =   1470
               End
            End
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
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
               Height          =   525
               Index           =   16
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   5265
               TabIndex        =   295
               Top             =   1500
               Width           =   5295
               Begin VB.CheckBox optMT 
                  Caption         =   "1ST"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   61
                  Top             =   0
                  Width           =   945
               End
               Begin VB.CheckBox optMT 
                  Caption         =   "2ND"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   2550
                  TabIndex        =   62
                  Top             =   0
                  Width           =   945
               End
               Begin VB.CheckBox optMT 
                  Caption         =   "3RD"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   2
                  Left            =   3810
                  TabIndex        =   63
                  Top             =   0
                  Width           =   675
               End
               Begin VB.CheckBox optMT 
                  Caption         =   "4TH"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   3
                  Left            =   4560
                  TabIndex        =   64
                  Top             =   0
                  Width           =   945
               End
               Begin VB.CheckBox optMT 
                  Caption         =   "5TH"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   4
                  Left            =   1290
                  TabIndex        =   65
                  Top             =   240
                  Width           =   945
               End
               Begin VB.CheckBox optMT 
                  Caption         =   "NUETRAL"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   5
                  Left            =   2550
                  TabIndex        =   66
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.CheckBox optMT 
                  Caption         =   "REVERSE"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   6
                  Left            =   3810
                  TabIndex        =   67
                  Top             =   240
                  Width           =   1365
               End
               Begin VB.Label lblcap 
                  BackColor       =   &H00FCE2CF&
                  Caption         =   " M/T"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   810
                  Index           =   74
                  Left            =   0
                  TabIndex        =   296
                  Top             =   0
                  Width           =   1260
               End
            End
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
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
               Height          =   525
               Index           =   17
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   5265
               TabIndex        =   293
               Top             =   2040
               Width           =   5295
               Begin VB.CheckBox optAT 
                  Caption         =   "D"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   3
                  Left            =   4560
                  TabIndex        =   71
                  Top             =   0
                  Width           =   1035
               End
               Begin VB.CheckBox optAT 
                  Caption         =   "1"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   68
                  Top             =   0
                  Width           =   1035
               End
               Begin VB.CheckBox optAT 
                  Caption         =   "2"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   2550
                  TabIndex        =   69
                  Top             =   0
                  Width           =   1035
               End
               Begin VB.CheckBox optAT 
                  Caption         =   "3"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   2
                  Left            =   3810
                  TabIndex        =   70
                  Top             =   0
                  Width           =   555
               End
               Begin VB.CheckBox optAT 
                  Caption         =   "N"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   4
                  Left            =   1290
                  TabIndex        =   72
                  Top             =   240
                  Width           =   1035
               End
               Begin VB.CheckBox optAT 
                  Caption         =   "R"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   5
                  Left            =   2550
                  TabIndex        =   73
                  Top             =   240
                  Width           =   1035
               End
               Begin VB.CheckBox optAT 
                  Caption         =   "OVERDRIVE"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   6
                  Left            =   3810
                  TabIndex        =   74
                  Top             =   240
                  Width           =   1425
               End
               Begin VB.Label lblcap 
                  BackColor       =   &H00FCE2CF&
                  Caption         =   " A/T"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   720
                  Index           =   75
                  Left            =   0
                  TabIndex        =   294
                  Top             =   0
                  Width           =   1260
               End
            End
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
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
               Height          =   555
               Index           =   18
               Left            =   30
               ScaleHeight     =   525
               ScaleWidth      =   5265
               TabIndex        =   291
               Top             =   2580
               Width           =   5295
               Begin VB.CheckBox optROAD 
                  Caption         =   "PAVED"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   75
                  Top             =   0
                  Width           =   1035
               End
               Begin VB.CheckBox optROAD 
                  Caption         =   "INPAVED"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   2550
                  TabIndex        =   76
                  Top             =   0
                  Width           =   1125
               End
               Begin VB.CheckBox optROAD 
                  Caption         =   "ROCKY"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   2
                  Left            =   3810
                  TabIndex        =   77
                  Top             =   0
                  Width           =   1035
               End
               Begin VB.CheckBox optROAD 
                  Caption         =   "MUDDY"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   3
                  Left            =   1290
                  TabIndex        =   78
                  Top             =   240
                  Width           =   1035
               End
               Begin VB.Label lblcap 
                  BackColor       =   &H00FCE2CF&
                  Caption         =   " ROAD"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   780
                  Index           =   76
                  Left            =   0
                  TabIndex        =   292
                  Top             =   0
                  Width           =   1260
               End
            End
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
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
               Height          =   525
               Index           =   19
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   5265
               TabIndex        =   289
               Top             =   3150
               Width           =   5295
               Begin VB.CheckBox optLOC 
                  Caption         =   "HIGHWAY"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   79
                  Top             =   0
                  Width           =   1245
               End
               Begin VB.CheckBox optLOC 
                  Caption         =   "UPHILL"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   2550
                  TabIndex        =   80
                  Top             =   0
                  Width           =   1065
               End
               Begin VB.CheckBox optLOC 
                  Caption         =   "DOWNHILL"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   2
                  Left            =   3810
                  TabIndex        =   81
                  Top             =   0
                  Width           =   1305
               End
               Begin VB.CheckBox optLOC 
                  Caption         =   "STOP AND GO TRAFFIC"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   3
                  Left            =   1290
                  TabIndex        =   82
                  Top             =   240
                  Width           =   2535
               End
               Begin VB.Label lblcap 
                  BackColor       =   &H00FCE2CF&
                  Caption         =   " LOCATION"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   660
                  Index           =   77
                  Left            =   0
                  TabIndex        =   290
                  Top             =   0
                  Width           =   1260
               End
            End
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
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
               Height          =   525
               Index           =   20
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   5265
               TabIndex        =   287
               Top             =   3690
               Width           =   5295
               Begin VB.CheckBox optACT 
                  Caption         =   "CRANKING"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   83
                  Top             =   0
                  Width           =   1185
               End
               Begin VB.CheckBox optACT 
                  Caption         =   "IDLING"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   2550
                  TabIndex        =   84
                  Top             =   0
                  Width           =   1065
               End
               Begin VB.CheckBox optACT 
                  Caption         =   "CRUISING"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   2
                  Left            =   3810
                  TabIndex        =   85
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.CheckBox optACT 
                  Caption         =   "ACCELERATING"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   3
                  Left            =   1290
                  TabIndex        =   86
                  Top             =   240
                  Width           =   1785
               End
               Begin VB.CheckBox optACT 
                  Caption         =   "DECELERATING"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   4
                  Left            =   3120
                  TabIndex        =   87
                  Top             =   240
                  Width           =   1905
               End
               Begin VB.Label lblcap 
                  BackColor       =   &H00FCE2CF&
                  Caption         =   " ACTION"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   540
                  Index           =   78
                  Left            =   0
                  TabIndex        =   288
                  Top             =   0
                  Width           =   1260
               End
            End
            Begin VB.PictureBox Picture5 
               Appearance      =   0  'Flat
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
               Height          =   285
               Index           =   21
               Left            =   30
               ScaleHeight     =   255
               ScaleWidth      =   5265
               TabIndex        =   285
               Top             =   4230
               Width           =   5295
               Begin VB.CheckBox optOCC 
                  Caption         =   "CONSISTENT"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   88
                  Top             =   0
                  Width           =   1455
               End
               Begin VB.CheckBox optOCC 
                  Caption         =   "INTERMITTENT"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   3120
                  TabIndex        =   89
                  Top             =   0
                  Width           =   1875
               End
               Begin VB.Label lblcap 
                  BackColor       =   &H00FCE2CF&
                  Caption         =   " OCCURENCE"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   79
                  Left            =   0
                  TabIndex        =   286
                  Top             =   0
                  Width           =   1260
               End
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
               Height          =   285
               Left            =   -60
               TabIndex        =   342
               Top             =   -30
               Width           =   9945
               _Version        =   655364
               _ExtentX        =   17542
               _ExtentY        =   503
               _StockProps     =   14
               Caption         =   "CONDITION"
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
               ForeColor       =   8388608
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
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
            Height          =   4965
            Left            =   5400
            ScaleHeight     =   4935
            ScaleWidth      =   4515
            TabIndex        =   223
            Top             =   0
            Width           =   4545
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
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
               Height          =   705
               Left            =   30
               ScaleHeight     =   675
               ScaleWidth      =   4425
               TabIndex        =   316
               Top             =   810
               Width           =   4455
               Begin VB.OptionButton optDEL 
                  Caption         =   "OTHERS"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   2790
                  MouseIcon       =   "frmCSMS_CQI.frx":6189
                  MousePointer    =   99  'Custom
                  TabIndex        =   97
                  Top             =   390
                  Width           =   1305
               End
               Begin VB.OptionButton optDEL 
                  Caption         =   "GAS STATION"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   2790
                  MouseIcon       =   "frmCSMS_CQI.frx":62DB
                  MousePointer    =   99  'Custom
                  TabIndex        =   95
                  Top             =   30
                  Width           =   1605
               End
               Begin VB.OptionButton optDEL 
                  Caption         =   "3-STAR SHOP"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   690
                  MouseIcon       =   "frmCSMS_CQI.frx":642D
                  MousePointer    =   99  'Custom
                  TabIndex        =   96
                  Top             =   360
                  Width           =   1665
               End
               Begin VB.OptionButton optDEL 
                  Caption         =   "DEALER"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   690
                  MouseIcon       =   "frmCSMS_CQI.frx":657F
                  MousePointer    =   99  'Custom
                  TabIndex        =   94
                  Top             =   30
                  Value           =   -1  'True
                  Width           =   1275
               End
            End
            Begin VB.PictureBox Picture8 
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
               Height          =   645
               Left            =   750
               ScaleHeight     =   645
               ScaleWidth      =   3615
               TabIndex        =   224
               Top             =   780
               Width           =   3615
            End
            Begin VB.TextBox txtEVERY 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   750
               TabIndex        =   92
               Top             =   360
               Width           =   645
            End
            Begin VB.TextBox txtKMS 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   3060
               TabIndex        =   93
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox txtSPE 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2340
               Left            =   30
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   98
               Top             =   2520
               Width           =   4395
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption14 
               Height          =   285
               Left            =   0
               TabIndex        =   343
               Top             =   -30
               Width           =   9945
               _Version        =   655364
               _ExtentX        =   17542
               _ExtentY        =   503
               _StockProps     =   14
               Caption         =   "VEHICLE MAINTENANCE"
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
               ForeColor       =   8388608
            End
            Begin VB.Label lblcap 
               Caption         =   "EVERY"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   25
               Left            =   180
               TabIndex        =   228
               Top             =   465
               Width           =   540
            End
            Begin VB.Label lblcap 
               Caption         =   "MONTH/S, EVERY "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   39
               Left            =   1680
               TabIndex        =   227
               Top             =   450
               Width           =   1305
            End
            Begin VB.Label lblcap 
               Caption         =   "KMS"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   41
               Left            =   4020
               TabIndex        =   226
               Top             =   450
               Width           =   465
            End
            Begin VB.Label lblcap 
               Caption         =   "OTHERS (PLS. SPECIFY)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   38
               Left            =   60
               TabIndex        =   225
               Top             =   2280
               Width           =   2280
            End
         End
         Begin VB.PictureBox picAUTHOR 
            Appearance      =   0  'Flat
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
            Height          =   2025
            Left            =   0
            ScaleHeight     =   1995
            ScaleWidth      =   9945
            TabIndex        =   216
            Top             =   5400
            Width           =   9975
            Begin VB.CommandButton cmdDEF 
               Caption         =   "DEFAULT SIGNATORIES"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Left            =   150
               TabIndex        =   334
               Top             =   1380
               Width           =   1695
            End
            Begin VB.TextBox txtREQ 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   120
               MaxLength       =   50
               TabIndex        =   99
               Top             =   660
               Width           =   3045
            End
            Begin VB.TextBox txtAPP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   6750
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   101
               Top             =   690
               Width           =   3045
            End
            Begin VB.TextBox txtCHECK 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   3510
               MaxLength       =   50
               TabIndex        =   100
               Top             =   690
               Width           =   3045
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
               Height          =   285
               Left            =   0
               TabIndex        =   344
               Top             =   0
               Width           =   9945
               _Version        =   655364
               _ExtentX        =   17542
               _ExtentY        =   503
               _StockProps     =   14
               Caption         =   "SIGNATORIES"
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
            Begin VB.Label lblcap 
               AutoSize        =   -1  'True
               Caption         =   "APPROVED BY"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   56
               Left            =   7680
               TabIndex        =   222
               Top             =   450
               Width           =   1020
            End
            Begin VB.Label lblcap 
               AutoSize        =   -1  'True
               Caption         =   "CHECKED BY"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   55
               Left            =   4380
               TabIndex        =   221
               Top             =   450
               Width           =   915
            End
            Begin VB.Label lblcap 
               AutoSize        =   -1  'True
               Caption         =   "REQUESTED BY"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   54
               Left            =   840
               TabIndex        =   220
               Top             =   420
               Width           =   1110
            End
            Begin VB.Label lblcap 
               AutoSize        =   -1  'True
               Caption         =   "HARI SERVICE DEPARTMENT"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   40
               Left            =   7080
               TabIndex        =   219
               Top             =   1050
               Width           =   2085
            End
            Begin VB.Label lblcap 
               AutoSize        =   -1  'True
               Caption         =   "DEALER SERVICE MANAGER"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   57
               Left            =   3810
               TabIndex        =   218
               Top             =   1050
               Width           =   2025
            End
            Begin VB.Label lblcap 
               AutoSize        =   -1  'True
               Caption         =   "DEALER WARRANTY ADMINISTRATOR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   58
               Left            =   30
               TabIndex        =   217
               Top             =   1020
               Width           =   2775
            End
         End
         Begin VB.Label lblDETCODE 
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
            Height          =   315
            Left            =   4590
            TabIndex        =   319
            Top             =   5010
            Visible         =   0   'False
            Width           =   3075
         End
         Begin VB.Label lblLINENO 
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
            Height          =   345
            Left            =   1200
            TabIndex        =   318
            Top             =   5010
            Visible         =   0   'False
            Width           =   2895
         End
      End
      Begin VB.PictureBox picNOTES 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   7395
         Left            =   -75000
         ScaleHeight     =   7365
         ScaleWidth      =   9975
         TabIndex        =   168
         Top             =   300
         Width           =   10005
         Begin VB.TextBox txtDESC 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1890
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   1320
            Width           =   7965
         End
         Begin VB.TextBox txtSUBJ 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1890
            MaxLength       =   250
            TabIndex        =   44
            Top             =   330
            Width           =   7995
         End
         Begin VB.TextBox txtHIST 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   1890
            MaxLength       =   250
            TabIndex        =   45
            Top             =   750
            Width           =   7965
         End
         Begin VB.TextBox txtCOMM 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1290
            Left            =   1890
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            Top             =   5850
            Width           =   7965
         End
         Begin VB.TextBox txtANA 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
            Left            =   1890
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            Top             =   2010
            Width           =   7995
         End
         Begin VB.TextBox txtREC 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   1890
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   4680
            Width           =   7995
         End
         Begin VB.TextBox txtCORR 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
            Left            =   1890
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            Top             =   3360
            Width           =   7995
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption6 
            Height          =   285
            Left            =   0
            TabIndex        =   341
            Top             =   -30
            Width           =   9945
            _Version        =   655364
            _ExtentX        =   17542
            _ExtentY        =   503
            _StockProps     =   14
            Caption         =   "NOTES"
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
            ForeColor       =   8388608
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "SUBJECT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   1185
            TabIndex        =   175
            Top             =   480
            Width           =   645
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "HISTORY"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   1170
            TabIndex        =   174
            Top             =   840
            Width           =   660
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   810
            TabIndex        =   173
            Top             =   1440
            Width           =   1020
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "OTHER COMMENTS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   43
            Left            =   435
            TabIndex        =   172
            Top             =   5820
            Width           =   1395
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            Caption         =   "RECOMMENDATION"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   390
            TabIndex        =   171
            Top             =   4680
            Width           =   1440
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "CORRECTIVE ACTION"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1065
            Index           =   23
            Left            =   855
            TabIndex        =   170
            Top             =   3360
            Width           =   945
         End
         Begin VB.Label lblcap 
            Alignment       =   1  'Right Justify
            Caption         =   "ANALYSIS (CAUSE OF THE PROBLEM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   720
            Index           =   22
            Left            =   450
            TabIndex        =   169
            Top             =   2040
            Width           =   1350
         End
      End
   End
   Begin VB.PictureBox picAddJob 
      Appearance      =   0  'Flat
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
      Height          =   2025
      Left            =   2603
      ScaleHeight     =   1995
      ScaleWidth      =   7155
      TabIndex        =   128
      Top             =   3480
      Visible         =   0   'False
      Width           =   7185
      Begin VB.CommandButton cmdDeleteJob 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   120
         MouseIcon       =   "frmCSMS_CQI.frx":66D1
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":6823
         Style           =   1  'Graphical
         TabIndex        =   232
         ToolTipText     =   "Delete Selected Record"
         Top             =   1020
         Width           =   705
      End
      Begin VB.CommandButton cmsCloseAJ 
         Caption         =   "&Cancel"
         Height          =   885
         Left            =   6390
         MouseIcon       =   "frmCSMS_CQI.frx":6B4E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":6CA0
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Cancel Entry"
         Top             =   1020
         Width           =   645
      End
      Begin VB.TextBox txtJobDesc 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   134
         Top             =   570
         Width           =   5205
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   675
         Left            =   9600
         MouseIcon       =   "frmCSMS_CQI.frx":6FDE
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":7130
         Style           =   1  'Graphical
         TabIndex        =   131
         ToolTipText     =   "Save Entry"
         Top             =   1020
         Width           =   645
      End
      Begin VB.CommandButton cmdCloseTerm 
         Caption         =   "&Cancel"
         Height          =   675
         Index           =   0
         Left            =   10230
         MouseIcon       =   "frmCSMS_CQI.frx":7480
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":75D2
         Style           =   1  'Graphical
         TabIndex        =   130
         ToolTipText     =   "Cancel Entry"
         Top             =   1020
         Width           =   645
      End
      Begin VB.TextBox txtJobCost 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5430
         TabIndex        =   135
         Top             =   570
         Width           =   1605
      End
      Begin VB.CommandButton cmdSaveJobs 
         Caption         =   "&Save"
         Height          =   885
         Left            =   5760
         MouseIcon       =   "frmCSMS_CQI.frx":7910
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":7A62
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Save Entry"
         Top             =   1020
         Width           =   645
      End
      Begin VB.Label lblLITEMNO 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5160
         TabIndex        =   138
         Top             =   0
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "COST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   67
         Left            =   6630
         TabIndex        =   133
         Top             =   300
         Width           =   405
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "JOB DESCRIPTION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   66
         Left            =   180
         TabIndex        =   132
         Top             =   330
         Width           =   1350
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption9 
         Height          =   255
         Left            =   0
         TabIndex        =   129
         Top             =   0
         Width           =   10005
         _Version        =   655364
         _ExtentX        =   17648
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "ADD/EDIT JOBS"
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
         ForeColor       =   8388608
      End
   End
   Begin VB.PictureBox picPWA 
      Appearance      =   0  'Flat
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
      Height          =   2070
      Left            =   4140
      ScaleHeight     =   2040
      ScaleWidth      =   4035
      TabIndex        =   233
      Top             =   3277
      Visible         =   0   'False
      Width           =   4065
      Begin VB.TextBox txtAPP1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   236
         Top             =   750
         Width           =   2745
      End
      Begin VB.CommandButton cmdCancelPWA 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   3240
         MouseIcon       =   "frmCSMS_CQI.frx":7DB2
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":7F04
         Style           =   1  'Graphical
         TabIndex        =   238
         ToolTipText     =   "Cancel"
         Top             =   1170
         Width           =   705
      End
      Begin VB.CommandButton cmdSavePWA 
         Caption         =   "&Save"
         Height          =   795
         Left            =   2550
         MouseIcon       =   "frmCSMS_CQI.frx":8242
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":8394
         Style           =   1  'Graphical
         TabIndex        =   237
         ToolTipText     =   "Save this Record"
         Top             =   1170
         Width           =   705
      End
      Begin VB.TextBox txtPWAno1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   235
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label lblcap 
         Alignment       =   1  'Right Justify
         Caption         =   "APPROVED BY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   42
         Left            =   345
         TabIndex        =   259
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "PWA NO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   53
         Left            =   540
         TabIndex        =   239
         Top             =   405
         Width           =   615
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption12 
         Height          =   255
         Left            =   0
         TabIndex        =   234
         Top             =   0
         Width           =   4815
         _Version        =   655364
         _ExtentX        =   8493
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "ENTER PWA NO"
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
         ForeColor       =   8388608
      End
   End
   Begin VB.PictureBox picAddPart 
      Appearance      =   0  'Flat
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
      Height          =   1935
      Left            =   690
      ScaleHeight     =   1905
      ScaleWidth      =   10905
      TabIndex        =   240
      Top             =   3345
      Visible         =   0   'False
      Width           =   10935
      Begin VB.TextBox cboPARTS 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         TabIndex        =   242
         ToolTipText     =   "Press Enter key to search Part Number"
         Top             =   510
         Width           =   1935
      End
      Begin VB.OptionButton optMAN 
         Caption         =   "Manual Entry of Part No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   870
         TabIndex        =   125
         Top             =   900
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6270
         TabIndex        =   245
         Top             =   510
         Width           =   825
      End
      Begin VB.TextBox txtPartCost 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7170
         TabIndex        =   246
         Top             =   510
         Width           =   1395
      End
      Begin VB.TextBox txtOPCODE 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8640
         TabIndex        =   247
         Top             =   510
         Width           =   1305
      End
      Begin VB.TextBox txtALTS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9990
         TabIndex        =   248
         Top             =   510
         Width           =   825
      End
      Begin VB.CommandButton cmdCloseAP 
         Caption         =   "&Cancel"
         Height          =   825
         Left            =   10170
         MouseIcon       =   "frmCSMS_CQI.frx":86E4
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":8836
         Style           =   1  'Graphical
         TabIndex        =   241
         ToolTipText     =   "Cancel Entry"
         Top             =   960
         Width           =   645
      End
      Begin VB.TextBox txtPartDesc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2100
         TabIndex        =   244
         Top             =   510
         Width           =   4125
      End
      Begin VB.CommandButton cmdDeletePart 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   90
         MouseIcon       =   "frmCSMS_CQI.frx":8B74
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":8CC6
         Style           =   1  'Graphical
         TabIndex        =   249
         ToolTipText     =   "Delete Selected Record"
         Top             =   960
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdSavePart 
         Caption         =   "&Save"
         Height          =   825
         Left            =   9540
         MouseIcon       =   "frmCSMS_CQI.frx":8FF1
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":9143
         Style           =   1  'Graphical
         TabIndex        =   250
         ToolTipText     =   "Save Entry"
         Top             =   960
         Width           =   645
      End
      Begin VB.TextBox txtPartNO 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5880
         TabIndex        =   243
         Top             =   1110
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label txtPartID 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   331
         Top             =   1080
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblPITEMNO 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   8940
         TabIndex        =   251
         Top             =   0
         Visible         =   0   'False
         Width           =   1905
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption8 
         Height          =   225
         Left            =   0
         TabIndex        =   258
         Top             =   0
         Width           =   10905
         _Version        =   655364
         _ExtentX        =   19235
         _ExtentY        =   397
         _StockProps     =   14
         Caption         =   "ADD/EDIT PARTS"
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
         ForeColor       =   8388608
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "PART NUMBER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   59
         Left            =   60
         TabIndex        =   257
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "PART NAME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   60
         Left            =   2160
         TabIndex        =   256
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   61
         Left            =   6510
         TabIndex        =   255
         Top             =   270
         Width           =   300
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "OP CODE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   62
         Left            =   8940
         TabIndex        =   254
         Top             =   270
         Width           =   675
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "COST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   63
         Left            =   7650
         TabIndex        =   253
         Top             =   270
         Width           =   405
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "LTS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   64
         Left            =   10290
         TabIndex        =   252
         Top             =   270
         Width           =   255
      End
   End
   Begin VB.PictureBox picAddSub 
      Appearance      =   0  'Flat
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
      Height          =   1965
      Left            =   1575
      ScaleHeight     =   1935
      ScaleWidth      =   10005
      TabIndex        =   139
      Top             =   3330
      Visible         =   0   'False
      Width           =   10035
      Begin VB.TextBox txtSUBQTY 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6210
         TabIndex        =   142
         Top             =   570
         Width           =   945
      End
      Begin VB.TextBox txtSLTS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8850
         TabIndex        =   144
         Top             =   570
         Width           =   1065
      End
      Begin VB.TextBox txtSOPCODE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7170
         TabIndex        =   143
         Top             =   570
         Width           =   1605
      End
      Begin VB.TextBox txtSubCode 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   140
         Top             =   570
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeleteSublet 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   120
         MouseIcon       =   "frmCSMS_CQI.frx":9493
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":95E5
         Style           =   1  'Graphical
         TabIndex        =   145
         ToolTipText     =   "Delete Selected Record"
         Top             =   990
         Width           =   705
      End
      Begin VB.TextBox txtSCOST 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4470
         TabIndex        =   148
         Text            =   "0.00"
         Top             =   1380
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox txtSubletDesc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1650
         TabIndex        =   141
         Top             =   570
         Width           =   4545
      End
      Begin VB.CommandButton cmdCloseAS 
         Caption         =   "&Cancel"
         Height          =   855
         Left            =   9300
         MouseIcon       =   "frmCSMS_CQI.frx":9910
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":9A62
         Style           =   1  'Graphical
         TabIndex        =   147
         ToolTipText     =   "Cancel Entry"
         Top             =   990
         Width           =   645
      End
      Begin VB.CommandButton cmdSaveSub 
         Caption         =   "&Save"
         Height          =   855
         Left            =   8670
         MouseIcon       =   "frmCSMS_CQI.frx":9DA0
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_CQI.frx":9EF2
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Save Entry"
         Top             =   990
         Width           =   645
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   91
         Left            =   6480
         TabIndex        =   332
         Top             =   360
         Width           =   300
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "LTS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   83
         Left            =   9210
         TabIndex        =   312
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "OP CODE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   82
         Left            =   7560
         TabIndex        =   311
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "CODE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   81
         Left            =   150
         TabIndex        =   310
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblSITEMNO 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   8010
         TabIndex        =   149
         Top             =   0
         Visible         =   0   'False
         Width           =   1905
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption10 
         Height          =   255
         Left            =   0
         TabIndex        =   152
         Top             =   0
         Width           =   10785
         _Version        =   655364
         _ExtentX        =   19024
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "ADD/EDIT SUBLET JOBS"
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
         ForeColor       =   8388608
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "SUBLET DESCRIPTION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   68
         Left            =   1710
         TabIndex        =   151
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         Caption         =   "COST"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   65
         Left            =   5670
         TabIndex        =   150
         Top             =   1110
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.PictureBox picSeachPlateNo 
      Appearance      =   0  'Flat
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
      Height          =   3645
      Left            =   1980
      ScaleHeight     =   3615
      ScaleWidth      =   8355
      TabIndex        =   320
      Top             =   2250
      Visible         =   0   'False
      Width           =   8385
      Begin XtremeReportControl.ReportControl rptPlate 
         Height          =   2805
         Left            =   90
         TabIndex        =   322
         Top             =   690
         Width           =   8175
         _Version        =   655364
         _ExtentX        =   14420
         _ExtentY        =   4948
         _StockProps     =   64
      End
      Begin VB.TextBox txtSeachP 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   321
         Top             =   300
         Width           =   8205
      End
      Begin wizButton.cmd cmd7 
         Height          =   225
         Left            =   8100
         TabIndex        =   324
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   397
         TX              =   "X"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCSMS_CQI.frx":A242
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   225
         Index           =   1
         Left            =   -30
         TabIndex        =   323
         Top             =   0
         Width           =   8385
         _Version        =   655364
         _ExtentX        =   14790
         _ExtentY        =   397
         _StockProps     =   14
         Caption         =   " SEARCH VEHICLE"
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
         ForeColor       =   8388608
      End
   End
   Begin VB.PictureBox picSEARCHI 
      Appearance      =   0  'Flat
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
      Height          =   4635
      Left            =   840
      ScaleHeight     =   4605
      ScaleWidth      =   10665
      TabIndex        =   124
      Top             =   1995
      Visible         =   0   'False
      Width           =   10695
      Begin XtremeReportControl.ReportControl rptLIST 
         Height          =   3645
         Left            =   60
         TabIndex        =   229
         Top             =   720
         Width           =   10515
         _Version        =   655364
         _ExtentX        =   18547
         _ExtentY        =   6429
         _StockProps     =   64
      End
      Begin VB.TextBox txtSEARCHI 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   333
         Top             =   330
         Width           =   10515
      End
      Begin wizButton.cmd cmd3 
         Height          =   225
         Left            =   10410
         TabIndex        =   127
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   397
         TX              =   "X"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCSMS_CQI.frx":A25E
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption7 
         Height          =   225
         Left            =   0
         TabIndex        =   126
         Top             =   0
         Width           =   10725
         _Version        =   655364
         _ExtentX        =   18918
         _ExtentY        =   397
         _StockProps     =   14
         Caption         =   "SEARCH INFORMATION"
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
         ForeColor       =   8388608
      End
   End
End
Attribute VB_Name = "frmCSMS_CQI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS                                                  As ADODB.Recordset
Dim ADD_OR_EDIT                                         As String
Dim ADD_OR_EDIT_ITEM                                    As String
Dim AUDIT_SQL                                           As String
Dim rsPartMas                                           As ADODB.Recordset
Dim xlApp                                               As Excel.Application
Dim xlBook                                              As Excel.Workbook
Dim xlSheet                                             As Excel.Worksheet
Dim PUBLIC_LTS                                          As Double
Dim xBEFORE_SAVE                                        As String
Dim WithEvents frm                                      As frmCSMS_MasterSearch
Attribute frm.VB_VarHelpID = -1

Function FindCodeDesc(vTABLE As String, vFIELD As String, vKFIELD As String, VCODE As String) As String
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT " & vFIELD & " as DESCR FROM " & vTABLE & " WHERE " & vKFIELD & " = '" & VCODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindCodeDesc = Null2String(RSTMP!DESCR)
    Else
        FindCodeDesc = ""
    End If

    Set RSTMP = Nothing
End Function

Function GenerateNewTranno()
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT TRANNO FROM CSMS_CQIR WHERE MONTH(TRANDATE) = " & Month(Date) & " AND YEAR(TRANDATE) = " & Year(Date) & " ORDER BY TRANNO DESC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        RSTMP.MoveFirst
        GenerateNewTranno = Format(RSTMP!TRANNO + 1, "00000")
    Else
        GenerateNewTranno = Format(1, "000")
    End If

    Set RSTMP = Nothing
End Function

Function GenerateDLRCODE()
    GenerateDLRCODE = COMPANY_CODE & Right(Year(Date), 2) & Format(Month(Date), "00")
End Function

Function GetBeforeValue(vID As Long, vTABLE As String) As String
    Dim rsKUTO                                         As New ADODB.Recordset
    Dim vResult                                        As String
    Set rsKUTO = gconDMIS.Execute("SELECT * FROM " & vTABLE & " WHERE ID = " & vID & "")
    MsgBox rsKUTO.GetString(adClipString, , "-")
    '    If Not (rsKUTO.BOF And rsKUTO.EOF) Then
    '        Do While Not rsKUTO.EOF
    '        RO-1000 ADDED EDITED R0-1001
    '    Else
    '
    '    End If

    Set rsKUTO = Nothing
End Function

Function GENERATEREFNO() As String
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT DLR_CQIR_REFERENCENO FROM CSMS_CQIR WHERE MONTH(TRANDATE) = " & Month(Date) & " AND YEAR(TRANDATE) = " & Year(Date) & " ORDER BY RIGHT(DLR_CQIR_REFERENCENO,3) DESC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        RSTMP.MoveFirst
        GENERATEREFNO = Right(RSTMP!DLR_CQIR_REFERENCENO, 3) + 1
    Else
        GENERATEREFNO = "001"
    End If
    Set RSTMP = Nothing
End Function

Function FINDTECHNICIAN(VTECHCODE As String) As String
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_VW_TECHNICIAN WHERE TECHNICIAN = '" & VTECHCODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FINDTECHNICIAN = Null2String(RSTMP!TECH_NAME)
    Else
        Set RSTMP = New ADODB.Recordset
        Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_CONTRACTOR WHERE CODE = '" & VTECHCODE & "'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            FINDTECHNICIAN = Null2String(RSTMP!CompanyName)
        Else
            FINDTECHNICIAN = ""
        End If
    End If
    Set RSTMP = Nothing
End Function

Function FindSAName(SCODE As String) As String
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_VW_EMPNO WHERE CODE = '" & SCODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindSAName = Null2String(RSTMP!NAYM)
    End If
    Set RSTMP = Nothing
End Function

Function SetPartIDSTOCKNO(DDD As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKNO from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(DDD) & "", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartIDSTOCKNO = Null2String(rsPartMas!ID)
    End If
    Set rsPartMas = Nothing
End Function

Function SetSTOCKDESC2(ppp As String)
    If ppp <> "" Then
        Set rsPartMas = New ADODB.Recordset
        rsPartMas.Open "Select id,STOCKDESC from PMIS_PARTMAS where id = " & ppp, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            SetSTOCKDESC2 = Null2String(rsPartMas!STOCKDESC)

        End If
    End If
    Set rsPartMas = Nothing
End Function

Public Function flex_FillReportView(RS As ADODB.Recordset, grd As XtremeReportControl.ReportControl, Optional ByVal WithSN As Boolean = False)
    Dim fld                                            As ADODB.FIELD
    Dim J                                              As Long
    Dim REC                                            As XtremeReportControl.ReportRecord

    grd.Records.DeleteAll

    While Not RS.EOF
        J = J + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem J
        End If
        For Each fld In RS.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        RS.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RS = Nothing
End Function

Sub DisplayWarrantyRO()
    Screen.MousePointer = 11
    Call ReportControlAddColumnHeader(rptLIST, "RO NO, FULL NAME, JOB DESCRIPTION, , , LINE NO, JOB CODE")
    Call ReportControlPaintManager(rptLIST)
    Call ResizeColumnHeader(rptLIST, "8, 20, 25, 0, 0, 0, 0")
    Call flex_FillReportView(gconDMIS.Execute("SELECT dbo.CSMS_Repor.REP_OR, dbo.CSMS_Repor.NIYM, dbo.CSMS_Ro_Det.DETDSC, dbo.CSMS_Repor.ID, " & _
                                            " dbo.CSMS_Ro_Det.ID AS Expr13, CSMS_Ro_Det.LINE_NO, dbo.CSMS_Ro_Det.DETCDE " & _
                                            " FROM dbo.CSMS_Repor INNER JOIN " & _
                                            " dbo.CSMS_Ro_Det ON dbo.CSMS_Repor.REP_OR = dbo.CSMS_Ro_Det.REP_OR " & _
                                            " WHERE dbo.CSMS_Repor.TRANSTYPE = 'R' AND (dbo.CSMS_Ro_Det.LIVIL = '1') AND (dbo.CSMS_Ro_Det.WCODE = 'W')"), rptLIST)
    Screen.MousePointer = 0
End Sub

Sub DisplayVEHICLES()
    Screen.MousePointer = 11
    Call ReportControlAddColumnHeader(rptPlate, "PLATE NO, CONDUCITON NO., VIN NO., ENGINE NO. , MODEL DESCRIPTION, ")
    Call ReportControlPaintManager(rptPlate)
    Call ResizeColumnHeader(rptPlate, "10, 20, 20, 20, 0")
    Call flex_FillReportView(gconDMIS.Execute("SELECT PLATE_NO, VCOND_NO, VIN, ENGINE, DESCRIPTION, ID FROM CSMS_CUSVEH"), rptPlate)
    Screen.MousePointer = 0
End Sub

Sub NextPage10()
    Dim RSTMP                                          As New ADODB.Recordset


    Dim xCNT                                           As Integer
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "QIR_BLANK.xlt")
    Set xlSheet = xlBook.Worksheets(1)

    xCNT = lsvPARTS.ListItems.Count

    If optPWAREQ1.Value = True Then
        xlSheet.Cells(7, "D") = "X"
        xlSheet.Cells(7, "G") = ""
    Else
        xlSheet.Cells(7, "G") = "X"
        xlSheet.Cells(7, "D") = ""
    End If

    xlSheet.Cells(5, "AA") = txtDLR
    xlSheet.Cells(9, "B") = txtPWA
    xlSheet.Cells(9, "Q") = txtPWAT
    xlSheet.Cells(10, "B") = dptDELDATTE.Value
    xlSheet.Cells(11, "B") = txtCUST.Text
    xlSheet.Cells(12, "B") = Mid(txtVIN, 1, 1)
    xlSheet.Cells(12, "C") = Mid(txtVIN, 2, 1)
    xlSheet.Cells(12, "D") = Mid(txtVIN, 3, 1)
    xlSheet.Cells(12, "E") = Mid(txtVIN, 4, 1)
    xlSheet.Cells(12, "F") = Mid(txtVIN, 5, 1)
    xlSheet.Cells(12, "G") = Mid(txtVIN, 6, 1)
    xlSheet.Cells(12, "H") = Mid(txtVIN, 7, 1)
    xlSheet.Cells(12, "I") = Mid(txtVIN, 8, 1)
    xlSheet.Cells(12, "J") = Mid(txtVIN, 9, 1)
    xlSheet.Cells(12, "K") = Mid(txtVIN, 10, 1)
    xlSheet.Cells(12, "L") = Mid(txtVIN, 11, 1)
    xlSheet.Cells(12, "M") = Mid(txtVIN, 12, 1)
    xlSheet.Cells(12, "N") = Mid(txtVIN, 13, 1)
    xlSheet.Cells(12, "O") = Mid(txtVIN, 14, 1)
    xlSheet.Cells(12, "P") = Mid(txtVIN, 15, 1)
    xlSheet.Cells(12, "Q") = Mid(txtVIN, 16, 1)
    xlSheet.Cells(12, "R") = Mid(txtVIN, 17, 1)

    xlSheet.Cells(13, "B") = Mid(txtENGINE, 1, 1)
    xlSheet.Cells(13, "C") = Mid(txtENGINE, 2, 1)
    xlSheet.Cells(13, "D") = Mid(txtENGINE, 3, 1)
    xlSheet.Cells(13, "E") = Mid(txtENGINE, 4, 1)
    xlSheet.Cells(13, "F") = Mid(txtENGINE, 5, 1)
    xlSheet.Cells(13, "G") = Mid(txtENGINE, 6, 1)
    xlSheet.Cells(13, "H") = Mid(txtENGINE, 7, 1)
    xlSheet.Cells(13, "I") = Mid(txtENGINE, 8, 1)
    xlSheet.Cells(13, "J") = Mid(txtENGINE, 9, 1)
    xlSheet.Cells(13, "K") = Mid(txtENGINE, 10, 1)
    xlSheet.Cells(13, "L") = Mid(txtENGINE, 11, 1)
    xlSheet.Cells(13, "M") = Mid(txtENGINE, 12, 1)

    xlSheet.Cells(14, "B") = Mid(txtAXLE, 1, 1)
    xlSheet.Cells(14, "C") = Mid(txtAXLE, 2, 1)
    xlSheet.Cells(14, "D") = Mid(txtAXLE, 3, 1)
    xlSheet.Cells(14, "E") = Mid(txtAXLE, 4, 1)
    xlSheet.Cells(14, "F") = Mid(txtAXLE, 5, 1)
    xlSheet.Cells(14, "G") = Mid(txtAXLE, 6, 1)
    xlSheet.Cells(14, "H") = Mid(txtAXLE, 7, 1)
    xlSheet.Cells(14, "J") = Mid(txtAXLE, 8, 1)
    xlSheet.Cells(14, "I") = Mid(txtAXLE, 9, 1)
    xlSheet.Cells(14, "K") = Mid(txtAXLE, 10, 1)
    xlSheet.Cells(14, "L") = Mid(txtAXLE, 11, 1)
    xlSheet.Cells(14, "M") = Mid(txtAXLE, 12, 1)

    If optTRAN1.Value = True Then
        xlSheet.Cells(15, "F") = "X": xlSheet.Cells(15, "L") = ""
    Else
        xlSheet.Cells(15, "F") = "": xlSheet.Cells(15, "L") = "X"
    End If

    If optATT1.Value = True Then
        xlSheet.Cells(16, "F") = "X": xlSheet.Cells(16, "L") = ""
    Else
        xlSheet.Cells(16, "F") = "": xlSheet.Cells(16, "L") = "X"
    End If

    xlSheet.Cells(17, "M") = txtModel.Text
    xlSheet.Cells(9, "AF") = txtDCODE.Text
    xlSheet.Cells(10, "AF") = dptINSD.Value
    xlSheet.Cells(11, "AF") = dptREPD.Value
    xlSheet.Cells(12, "AF") = txtRO.Text
    xlSheet.Cells(13, "AF") = txtKM.Text
    xlSheet.Cells(14, "AF") = txtPLATENO.Text

    If optSENT1.Value = True Then
        xlSheet.Cells(15, "AI") = "X": xlSheet.Cells(15, "AM") = ""
    Else
        xlSheet.Cells(15, "AI") = "": xlSheet.Cells(15, "AM") = "X"
    End If

    xlSheet.Cells(16, "AF") = cboSA
    xlSheet.Cells(17, "AF") = cboTECH

    xlSheet.Cells(19, "B") = Replace(txtSUBJ.Text, vbCrLf, " ")
    xlSheet.Cells(21, "B") = Replace(txtHIST, vbCrLf, " ")
    xlSheet.Cells(24, "B") = Replace(txtDESC.Text, vbCrLf, " ")
    xlSheet.Cells(26, "B") = Replace(txtANA.Text, vbCrLf, " ")
    xlSheet.Cells(38, "B") = Replace(txtCORR.Text, vbCrLf, " ")
    xlSheet.Cells(42, "B") = Replace(txtREC.Text, vbCrLf, " ")

    xlSheet.Cells(46, "A") = txtCASUAL.Text
    xlSheet.Cells(45, "I") = cboNCODE.Text
    xlSheet.Cells(46, "I") = cboCCODE.Text
    xlSheet.Cells(45, "R") = txtPCODE.Text
    xlSheet.Cells(46, "R") = txtSCODE.Text

    Dim INDEX                                          As Integer
    Dim VSUB                                           As Currency
    Dim VPART                                          As Currency
    INDEX = 49
    Dim CNT_ITEM                                       As Integer
    CNT_ITEM = 0
    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_CQIRPARTS WHERE DLR_CQIR_REFERENCENO = '" & txtDLR.Text & "' AND (ISSUBLET IS NULL OR ISSUBLET = 'S') ORDER BY ID")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        RSTMP.MoveFirst
        Dim cnt_TEN                                    As Integer
        Do While Not RSTMP.EOF
            cnt_TEN = cnt_TEN + 1
            RSTMP.MoveNext
            If cnt_TEN = 10 Then GoTo COUNT_TEN
        Loop

        Do While Not RSTMP.EOF
COUNT_TEN:
            If CNT_ITEM = 10 Then GoTo CONT_DISPLAY

            xlSheet.Cells(INDEX, "A") = Null2String(RSTMP!partno)
            xlSheet.Cells(INDEX, "B") = Null2String(RSTMP!partname)
            xlSheet.Cells(INDEX, "J") = Null2String(RSTMP!QTY)
            xlSheet.Cells(INDEX, "L") = Format(NumericVal(RSTMP!cost), "#,###,##0.00")
            xlSheet.Cells(INDEX, "O") = Null2String(RSTMP!OPCODE)
            xlSheet.Cells(INDEX, "S") = Format(NumericVal(RSTMP!lts), "#,###,##0.00")

            If Null2String(RSTMP!ISSUBLET) = "" Then
                VPART = VPART + NumericVal(RSTMP!cost)
            Else
                VSUB = VSUB + NumericVal(RSTMP!cost)
            End If

            CNT_ITEM = CNT_ITEM + 1
            RSTMP.MoveNext
            INDEX = INDEX + 1
        Loop
    End If
    Set RSTMP = Nothing

    VSUB = TXTSUBREP.Text
CONT_DISPLAY:
    xlSheet.Cells(59, "L") = VPART
    xlSheet.Cells(63, "I") = VSUB

    'ENGINE TEMP
    If optENG(0).Value = 1 Then xlSheet.Cells(23, "AC") = "X" Else xlSheet.Cells(23, "AC") = ""
    If optENG(1).Value = 1 Then xlSheet.Cells(23, "AF") = "X" Else xlSheet.Cells(23, "AF") = ""
    If optENG(2).Value = 1 Then xlSheet.Cells(23, "AI") = "X" Else xlSheet.Cells(23, "AI") = ""

    'WEATHER
    If optWEA(0).Value = 1 Then xlSheet.Cells(25, "AC") = "X" Else xlSheet.Cells(25, "AC") = ""
    If optWEA(1).Value = 1 Then xlSheet.Cells(25, "AF") = "X" Else xlSheet.Cells(25, "AF") = ""
    If optWEA(2).Value = 1 Then xlSheet.Cells(25, "AI") = "X" Else xlSheet.Cells(25, "AI") = ""

    'SHIFTING
    If optSHIF(0).Value = 1 Then xlSheet.Cells(27, "AA") = "X" Else xlSheet.Cells(27, "AA") = ""
    If optSHIF(1).Value = 1 Then xlSheet.Cells(27, "AF") = "X" Else xlSheet.Cells(27, "AF") = ""

    'SHIFT POSITION
    If optPOS(0).Value = 1 Then xlSheet.Cells(27, "AA") = "X" Else xlSheet.Cells(27, "AA") = ""
    If optPOS(1).Value = 1 Then xlSheet.Cells(29, "AH") = "X" Else xlSheet.Cells(29, "AH") = ""

    'M/T
    If optMT(0).Value = 1 Then xlSheet.Cells(31, "Z") = "X" Else xlSheet.Cells(31, "Z") = ""
    If optMT(1).Value = 1 Then xlSheet.Cells(31, "AC") = "X" Else xlSheet.Cells(31, "AC") = ""
    If optMT(2).Value = 1 Then xlSheet.Cells(31, "AF") = "X" Else xlSheet.Cells(31, "AF") = ""
    If optMT(3).Value = 1 Then xlSheet.Cells(31, "AI") = "X" Else xlSheet.Cells(31, "AI") = ""
    If optMT(4).Value = 1 Then xlSheet.Cells(33, "Z") = "X" Else xlSheet.Cells(33, "Z") = ""
    If optMT(5).Value = 1 Then xlSheet.Cells(33, "AC") = "X" Else xlSheet.Cells(33, "AC") = ""
    If optMT(6).Value = 1 Then xlSheet.Cells(33, "AG") = "X" Else xlSheet.Cells(33, "AG") = ""

    'A/T
    If optAT(0).Value = 1 Then xlSheet.Cells(35, "Z") = "X" Else xlSheet.Cells(35, "Z") = ""
    If optAT(1).Value = 1 Then xlSheet.Cells(35, "AD") = "X" Else xlSheet.Cells(35, "AD") = ""
    If optAT(2).Value = 1 Then xlSheet.Cells(35, "AH") = "X" Else xlSheet.Cells(35, "AH") = ""
    If optAT(3).Value = 1 Then xlSheet.Cells(35, "AK") = "X" Else xlSheet.Cells(35, "AK") = ""
    If optAT(4).Value = 1 Then xlSheet.Cells(37, "AA") = "X" Else xlSheet.Cells(37, "AA") = ""
    If optAT(5).Value = 1 Then xlSheet.Cells(37, "AE") = "X" Else xlSheet.Cells(37, "AE") = ""
    If optAT(6).Value = 1 Then xlSheet.Cells(37, "AI") = "X" Else xlSheet.Cells(37, "AI") = ""

    'ROAD
    If optROAD(0).Value = 1 Then xlSheet.Cells(39, "Z") = "X" Else xlSheet.Cells(39, "Z") = ""
    If optROAD(1).Value = 1 Then xlSheet.Cells(39, "AD") = "X" Else xlSheet.Cells(39, "AD") = ""
    If optROAD(2).Value = 1 Then xlSheet.Cells(39, "AI") = "X" Else xlSheet.Cells(39, "AI") = ""
    If optROAD(3).Value = 1 Then xlSheet.Cells(41, "AD") = "X" Else xlSheet.Cells(41, "AD") = ""

    'LOCATION
    If optLOC(0).Value = 1 Then xlSheet.Cells(43, "AA") = "X" Else xlSheet.Cells(43, "AA") = ""
    If optLOC(1).Value = 1 Then xlSheet.Cells(43, "AF") = "X" Else xlSheet.Cells(43, "AF") = ""
    If optLOC(2).Value = 1 Then xlSheet.Cells(43, "AJ") = "X" Else xlSheet.Cells(43, "AJ") = ""
    If optLOC(3).Value = 1 Then xlSheet.Cells(45, "AA") = "X" Else xlSheet.Cells(43, "AA") = ""


    'ACTION
    If optACT(0).Value = 1 Then xlSheet.Cells(47, "AA") = "X" Else xlSheet.Cells(47, "AA") = ""
    If optACT(1).Value = 1 Then xlSheet.Cells(47, "AF") = "X" Else xlSheet.Cells(47, "AF") = ""
    If optACT(2).Value = 1 Then xlSheet.Cells(47, "AJ") = "X" Else xlSheet.Cells(47, "AJ") = ""
    If optACT(3).Value = 1 Then xlSheet.Cells(49, "AA") = "X" Else xlSheet.Cells(49, "AA") = ""
    If optACT(4).Value = 1 Then xlSheet.Cells(49, "AG") = "X" Else xlSheet.Cells(49, "AG") = ""

    'OCCURENCE
    If optOCC(0).Value = 1 Then xlSheet.Cells(51, "AB") = "X" Else xlSheet.Cells(51, "AG") = ""
    If optOCC(1).Value = 1 Then xlSheet.Cells(51, "AB") = "" Else xlSheet.Cells(51, "AG") = ""

    'ACCESSORIES
    If optACC(0).Value = 1 Then xlSheet.Cells(53, "AB") = "X" Else xlSheet.Cells(53, "AB") = ""
    If optACC(1).Value = 1 Then xlSheet.Cells(53, "AF") = "X" Else xlSheet.Cells(53, "AF") = ""

    'VEHICLE MAINTENANCE
    If optDEL(0).Value = True Then xlSheet.Cells(59, "X") = "X": xlSheet.Cells(59, "AC") = "": xlSheet.Cells(59, "AI") = "": xlSheet.Cells(59, "AE") = ""
    If optDEL(1).Value = True Then xlSheet.Cells(59, "X") = "": xlSheet.Cells(59, "AC") = "X": xlSheet.Cells(59, "AI") = "": xlSheet.Cells(59, "AE") = ""
    If optDEL(2).Value = True Then xlSheet.Cells(59, "X") = "": xlSheet.Cells(59, "AC") = "": xlSheet.Cells(59, "AI") = "X": xlSheet.Cells(59, "AE") = ""
    If optDEL(3).Value = True Then xlSheet.Cells(59, "X") = "": xlSheet.Cells(59, "AC") = "": xlSheet.Cells(59, "AI") = "": xlSheet.Cells(59, "AE") = "X"


    xlSheet.Cells(57, "AA") = txtEVERY
    xlSheet.Cells(57, "AI") = txtKMS
    xlSheet.Cells(62, "AE") = txtSPE.Text


    xlSheet.Cells(64, "N") = Format(txtINV.Text, "000000")
    xlSheet.Cells(66, "A") = "Other Comments: " & txtCOMM.Text
    xlSheet.Cells(72, "A") = txtREQ.Text
    xlSheet.Cells(72, "I") = txtCHECK.Text
    xlSheet.Cells(72, "X") = txtAPP.Text

    xlApp.Windows.Item(1).Caption = "QIR NO: " & txtDLR
    xlApp.Visible = True
    'xlBook.Close
    Set xlApp = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
End Sub

Sub PRINTEXCEL()
    Dim RSTMP                                          As New ADODB.Recordset


    Dim xCNT                                           As Integer
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "QIR_BLANK.xlt")
    Set xlSheet = xlBook.Worksheets(1)

    xCNT = lsvPARTS.ListItems.Count

    If optPWAREQ1.Value = True Then
        xlSheet.Cells(7, "D") = "X"
        xlSheet.Cells(7, "G") = ""
    Else
        xlSheet.Cells(7, "G") = "X"
        xlSheet.Cells(7, "D") = ""
    End If

    xlSheet.Cells(5, "AA") = txtDLR
    xlSheet.Cells(9, "B") = txtPWA
    xlSheet.Cells(9, "Q") = txtPWAT
    xlSheet.Cells(10, "B") = dptDELDATTE.Value
    xlSheet.Cells(11, "B") = txtCUST.Text
    xlSheet.Cells(12, "B") = Mid(txtVIN, 1, 1)
    xlSheet.Cells(12, "C") = Mid(txtVIN, 2, 1)
    xlSheet.Cells(12, "D") = Mid(txtVIN, 3, 1)
    xlSheet.Cells(12, "E") = Mid(txtVIN, 4, 1)
    xlSheet.Cells(12, "F") = Mid(txtVIN, 5, 1)
    xlSheet.Cells(12, "G") = Mid(txtVIN, 6, 1)
    xlSheet.Cells(12, "H") = Mid(txtVIN, 7, 1)
    xlSheet.Cells(12, "I") = Mid(txtVIN, 8, 1)
    xlSheet.Cells(12, "J") = Mid(txtVIN, 9, 1)
    xlSheet.Cells(12, "K") = Mid(txtVIN, 10, 1)
    xlSheet.Cells(12, "L") = Mid(txtVIN, 11, 1)
    xlSheet.Cells(12, "M") = Mid(txtVIN, 12, 1)
    xlSheet.Cells(12, "N") = Mid(txtVIN, 13, 1)
    xlSheet.Cells(12, "O") = Mid(txtVIN, 14, 1)
    xlSheet.Cells(12, "P") = Mid(txtVIN, 15, 1)
    xlSheet.Cells(12, "Q") = Mid(txtVIN, 16, 1)
    xlSheet.Cells(12, "R") = Mid(txtVIN, 17, 1)

    xlSheet.Cells(13, "B") = Mid(txtENGINE, 1, 1)
    xlSheet.Cells(13, "C") = Mid(txtENGINE, 2, 1)
    xlSheet.Cells(13, "D") = Mid(txtENGINE, 3, 1)
    xlSheet.Cells(13, "E") = Mid(txtENGINE, 4, 1)
    xlSheet.Cells(13, "F") = Mid(txtENGINE, 5, 1)
    xlSheet.Cells(13, "G") = Mid(txtENGINE, 6, 1)
    xlSheet.Cells(13, "H") = Mid(txtENGINE, 7, 1)
    xlSheet.Cells(13, "I") = Mid(txtENGINE, 8, 1)
    xlSheet.Cells(13, "J") = Mid(txtENGINE, 9, 1)
    xlSheet.Cells(13, "K") = Mid(txtENGINE, 10, 1)
    xlSheet.Cells(13, "L") = Mid(txtENGINE, 11, 1)
    xlSheet.Cells(13, "M") = Mid(txtENGINE, 12, 1)

    xlSheet.Cells(14, "B") = Mid(txtAXLE, 1, 1)
    xlSheet.Cells(14, "C") = Mid(txtAXLE, 2, 1)
    xlSheet.Cells(14, "D") = Mid(txtAXLE, 3, 1)
    xlSheet.Cells(14, "E") = Mid(txtAXLE, 4, 1)
    xlSheet.Cells(14, "F") = Mid(txtAXLE, 5, 1)
    xlSheet.Cells(14, "G") = Mid(txtAXLE, 6, 1)
    xlSheet.Cells(14, "H") = Mid(txtAXLE, 7, 1)
    xlSheet.Cells(14, "J") = Mid(txtAXLE, 8, 1)
    xlSheet.Cells(14, "I") = Mid(txtAXLE, 9, 1)
    xlSheet.Cells(14, "K") = Mid(txtAXLE, 10, 1)
    xlSheet.Cells(14, "L") = Mid(txtAXLE, 11, 1)
    xlSheet.Cells(14, "M") = Mid(txtAXLE, 12, 1)

    If optTRAN1.Value = True Then
        xlSheet.Cells(15, "F") = "X": xlSheet.Cells(15, "L") = ""
    Else
        xlSheet.Cells(15, "F") = "": xlSheet.Cells(15, "L") = "X"
    End If

    If optATT1.Value = True Then
        xlSheet.Cells(16, "F") = "X": xlSheet.Cells(16, "L") = ""
    Else
        xlSheet.Cells(16, "F") = "": xlSheet.Cells(16, "L") = "X"
    End If

    xlSheet.Cells(17, "M") = txtModel.Text
    xlSheet.Cells(9, "AF") = txtDCODE.Text
    xlSheet.Cells(10, "AF") = dptINSD.Value
    xlSheet.Cells(11, "AF") = dptREPD.Value
    xlSheet.Cells(12, "AF") = txtRO.Text
    xlSheet.Cells(13, "AF") = txtKM.Text
    xlSheet.Cells(14, "AF") = txtPLATENO.Text

    If optSENT1.Value = True Then
        xlSheet.Cells(15, "AI") = "X": xlSheet.Cells(15, "AM") = ""
    Else
        xlSheet.Cells(15, "AI") = "": xlSheet.Cells(15, "AM") = "X"
    End If

    xlSheet.Cells(16, "AF") = cboSA
    xlSheet.Cells(17, "AF") = cboTECH

    xlSheet.Cells(19, "B") = Replace(RTrim(LTrim(txtSUBJ.Text)), vbCrLf, "")
    xlSheet.Cells(21, "B") = Replace(RTrim(LTrim(txtHIST)), vbCrLf, "")
    xlSheet.Cells(24, "B") = Replace(RTrim(LTrim(txtDESC.Text)), vbCrLf, "")
    xlSheet.Cells(26, "B") = Replace(RTrim(LTrim(txtANA.Text)), vbCrLf, "")
    xlSheet.Cells(38, "B") = Replace(RTrim(LTrim(txtCORR.Text)), vbCrLf, "")
    xlSheet.Cells(42, "B") = Replace(RTrim(LTrim(txtREC.Text)), vbCrLf, "")

    xlSheet.Cells(46, "A") = txtCASUAL.Text
    xlSheet.Cells(45, "I") = cboNCODE.Text
    xlSheet.Cells(46, "I") = cboCCODE.Text
    xlSheet.Cells(45, "R") = txtPCODE.Text
    xlSheet.Cells(46, "R") = txtSCODE.Text

    Dim INDEX                                          As Integer
    Dim VSUB                                           As Currency
    Dim VPART                                          As Currency
    Dim k_LTS                                          As Double

    INDEX = 49
    Dim CNT_ITEM                                       As Integer
    CNT_ITEM = 0
    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_CQIRPARTS WHERE DLR_CQIR_REFERENCENO = '" & txtDLR.Text & "' AND (ISSUBLET IS NULL OR ISSUBLET = 'S') ORDER BY ID")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        RSTMP.MoveFirst
        Dim tmp_cnt                                    As Integer
        Do While Not RSTMP.EOF
            tmp_cnt = tmp_cnt + 1
            RSTMP.MoveNext
        Loop

        RSTMP.MoveFirst
        Do While Not RSTMP.EOF
            If CNT_ITEM = 10 Then GoTo CONT_DISPLAY

            xlSheet.Cells(INDEX, "A") = Null2String(RSTMP!partno)
            xlSheet.Cells(INDEX, "B") = Null2String(RSTMP!partname)
            xlSheet.Cells(INDEX, "J") = Null2String(RSTMP!QTY)
            xlSheet.Cells(INDEX, "L") = Format(NumericVal(RSTMP!cost), "#,###,##0.00")
            xlSheet.Cells(INDEX, "O") = Null2String(RSTMP!OPCODE)
            xlSheet.Cells(INDEX, "S") = Format(NumericVal(RSTMP!lts), "#,###,##0.00")

            If Null2String(RSTMP!ISSUBLET) = "" Then
                VPART = VPART + NumericVal(RSTMP!cost)
            Else
                VSUB = VSUB + NumericVal(RSTMP!cost)
            End If

            k_LTS = k_LTS + NumericVal(RSTMP!lts)
            CNT_ITEM = CNT_ITEM + 1
            RSTMP.MoveNext
            INDEX = INDEX + 1
        Loop
    End If
    Set RSTMP = Nothing

    VSUB = TXTSUBREP.Text
CONT_DISPLAY:
    xlSheet.Cells(59, "L") = VPART
    xlSheet.Cells(60, "A") = "NOTE (Labor Cost): " & lblFLATRATE & " x Total LTS"
    xlSheet.Cells(62, "I") = NumericVal(lblFLATRATE * k_LTS)
    xlSheet.Cells(63, "I") = VSUB

    'ENGINE TEMP
    If optENG(0).Value = 1 Then xlSheet.Cells(23, "AC") = "X" Else xlSheet.Cells(23, "AC") = ""
    If optENG(1).Value = 1 Then xlSheet.Cells(23, "AF") = "X" Else xlSheet.Cells(23, "AF") = ""
    If optENG(2).Value = 1 Then xlSheet.Cells(23, "AI") = "X" Else xlSheet.Cells(23, "AI") = ""

    'WEATHER
    If optWEA(0).Value = 1 Then xlSheet.Cells(25, "AC") = "X" Else xlSheet.Cells(25, "AC") = ""
    If optWEA(1).Value = 1 Then xlSheet.Cells(25, "AF") = "X" Else xlSheet.Cells(25, "AF") = ""
    If optWEA(2).Value = 1 Then xlSheet.Cells(25, "AI") = "X" Else xlSheet.Cells(25, "AI") = ""

    'SHIFTING
    If optSHIF(0).Value = 1 Then xlSheet.Cells(27, "AA") = "X" Else xlSheet.Cells(27, "AA") = ""
    If optSHIF(1).Value = 1 Then xlSheet.Cells(27, "AF") = "X" Else xlSheet.Cells(27, "AF") = ""

    'SHIFT POSITION
    If optPOS(0).Value = 1 Then xlSheet.Cells(27, "AA") = "X" Else xlSheet.Cells(27, "AA") = ""
    If optPOS(1).Value = 1 Then xlSheet.Cells(29, "AH") = "X" Else xlSheet.Cells(29, "AH") = ""

    'M/T
    If optMT(0).Value = 1 Then xlSheet.Cells(31, "Z") = "X" Else xlSheet.Cells(31, "Z") = ""
    If optMT(1).Value = 1 Then xlSheet.Cells(31, "AC") = "X" Else xlSheet.Cells(31, "AC") = ""
    If optMT(2).Value = 1 Then xlSheet.Cells(31, "AF") = "X" Else xlSheet.Cells(31, "AF") = ""
    If optMT(3).Value = 1 Then xlSheet.Cells(31, "AI") = "X" Else xlSheet.Cells(31, "AI") = ""
    If optMT(4).Value = 1 Then xlSheet.Cells(33, "Z") = "X" Else xlSheet.Cells(33, "Z") = ""
    If optMT(5).Value = 1 Then xlSheet.Cells(33, "AC") = "X" Else xlSheet.Cells(33, "AC") = ""
    If optMT(6).Value = 1 Then xlSheet.Cells(33, "AG") = "X" Else xlSheet.Cells(33, "AG") = ""

    'A/T
    If optAT(0).Value = 1 Then xlSheet.Cells(35, "Z") = "X" Else xlSheet.Cells(35, "Z") = ""
    If optAT(1).Value = 1 Then xlSheet.Cells(35, "AD") = "X" Else xlSheet.Cells(35, "AD") = ""
    If optAT(2).Value = 1 Then xlSheet.Cells(35, "AH") = "X" Else xlSheet.Cells(35, "AH") = ""
    If optAT(3).Value = 1 Then xlSheet.Cells(35, "AK") = "X" Else xlSheet.Cells(35, "AK") = ""
    If optAT(4).Value = 1 Then xlSheet.Cells(37, "AA") = "X" Else xlSheet.Cells(37, "AA") = ""
    If optAT(5).Value = 1 Then xlSheet.Cells(37, "AE") = "X" Else xlSheet.Cells(37, "AE") = ""
    If optAT(6).Value = 1 Then xlSheet.Cells(37, "AI") = "X" Else xlSheet.Cells(37, "AI") = ""

    'ROAD
    If optROAD(0).Value = 1 Then xlSheet.Cells(39, "Z") = "X" Else xlSheet.Cells(39, "Z") = ""
    If optROAD(1).Value = 1 Then xlSheet.Cells(39, "AD") = "X" Else xlSheet.Cells(39, "AD") = ""
    If optROAD(2).Value = 1 Then xlSheet.Cells(39, "AI") = "X" Else xlSheet.Cells(39, "AI") = ""
    If optROAD(3).Value = 1 Then xlSheet.Cells(41, "AD") = "X" Else xlSheet.Cells(41, "AD") = ""

    'LOCATION
    If optLOC(0).Value = 1 Then xlSheet.Cells(43, "AA") = "X" Else xlSheet.Cells(43, "AA") = ""
    If optLOC(1).Value = 1 Then xlSheet.Cells(43, "AF") = "X" Else xlSheet.Cells(43, "AF") = ""
    If optLOC(2).Value = 1 Then xlSheet.Cells(43, "AJ") = "X" Else xlSheet.Cells(43, "AJ") = ""
    If optLOC(3).Value = 1 Then xlSheet.Cells(45, "AA") = "X" Else xlSheet.Cells(43, "AA") = ""


    'ACTION
    If optACT(0).Value = 1 Then xlSheet.Cells(47, "AA") = "X" Else xlSheet.Cells(47, "AA") = ""
    If optACT(1).Value = 1 Then xlSheet.Cells(47, "AF") = "X" Else xlSheet.Cells(47, "AF") = ""
    If optACT(2).Value = 1 Then xlSheet.Cells(47, "AJ") = "X" Else xlSheet.Cells(47, "AJ") = ""
    If optACT(3).Value = 1 Then xlSheet.Cells(49, "AA") = "X" Else xlSheet.Cells(49, "AA") = ""
    If optACT(4).Value = 1 Then xlSheet.Cells(49, "AG") = "X" Else xlSheet.Cells(49, "AG") = ""

    'OCCURENCE
    If optOCC(0).Value = 1 Then xlSheet.Cells(51, "AB") = "X" Else xlSheet.Cells(51, "AG") = ""
    If optOCC(1).Value = 1 Then xlSheet.Cells(51, "AB") = "" Else xlSheet.Cells(51, "AG") = ""

    'ACCESSORIES
    If optACC(0).Value = 1 Then xlSheet.Cells(53, "AB") = "X" Else xlSheet.Cells(53, "AB") = ""
    If optACC(1).Value = 1 Then xlSheet.Cells(53, "AF") = "X" Else xlSheet.Cells(53, "AF") = ""

    'VEHICLE MAINTENANCE
    If optDEL(0).Value = True Then xlSheet.Cells(59, "X") = "X": xlSheet.Cells(59, "AC") = "": xlSheet.Cells(59, "AI") = "": xlSheet.Cells(59, "AE") = ""
    If optDEL(1).Value = True Then xlSheet.Cells(59, "X") = "": xlSheet.Cells(59, "AC") = "X": xlSheet.Cells(59, "AI") = "": xlSheet.Cells(59, "AE") = ""
    If optDEL(2).Value = True Then xlSheet.Cells(59, "X") = "": xlSheet.Cells(59, "AC") = "": xlSheet.Cells(59, "AI") = "X": xlSheet.Cells(59, "AE") = ""
    If optDEL(3).Value = True Then xlSheet.Cells(59, "X") = "": xlSheet.Cells(59, "AC") = "": xlSheet.Cells(59, "AI") = "": xlSheet.Cells(59, "AE") = "X"


    xlSheet.Cells(57, "AA") = txtEVERY
    xlSheet.Cells(57, "AI") = txtKMS
    xlSheet.Cells(62, "AE") = txtSPE.Text


    xlSheet.Cells(64, "N") = Format(txtINV.Text, "000000")
    xlSheet.Cells(66, "A") = "Other Comments: " & txtCOMM.Text
    xlSheet.Cells(72, "A") = txtREQ.Text
    xlSheet.Cells(72, "I") = txtCHECK.Text
    xlSheet.Cells(72, "X") = txtAPP.Text

    xlApp.Windows.Item(1).Caption = "QIR NO: " & txtDLR
    xlApp.Visible = True
    'xlBook.Close
    Set xlApp = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing

    If tmp_cnt > 10 Then
        Call NextPage10
    End If
    If tmp_cnt > 20 Then

    End If
End Sub

Sub DisabledPics(COND As Boolean)
    picSEARCH.Enabled = COND
    picDET.Enabled = COND
    picHD.Enabled = Not COND
    picNOTES.Enabled = Not COND
    picCONDITIONS.Enabled = Not COND
End Sub

Sub DisabledSaveMenu(COND As Boolean)
    picSaves.Enabled = COND
    picAdds.Enabled = COND
    picMENU.Enabled = COND
    picSEARCH.Enabled = COND
    picHD.Enabled = COND
    picDET.Enabled = COND
    picNOTES.Enabled = COND
    picCONDITIONS.Enabled = COND
End Sub

Sub InitMemVars1()
    txtPartNO.Text = ""
    txtPartDesc.Text = ""
    txtQty.Text = ""
    txtPartCost.Text = ""
    txtOPCODE.Text = ""
    txtALTS.Text = ""
    txtSubletDesc.Text = ""
    txtSCOST.Text = ""
    txtJobDesc.Text = ""
    txtJobCost.Text = ""
End Sub

Sub InitCbo()
    Dim rsParts                                        As ADODB.Recordset
    Combo_Loadval cboCCODE, gconDMIS.Execute("SELECT CAUECODE FROM WWTCAUE ORDER BY CAUECODE")
    Combo_Loadval cboNCODE, gconDMIS.Execute("SELECT NATRCODE FROM WWTNATR ORDER BY NATRCODE")
    Combo_Loadval cboSA, gconDMIS.Execute("SELECT NAYM FROM CSMS_VW_EMPNO ORDER BY NAYM")
    Combo_Loadval cboTECH, gconDMIS.Execute("SELECT TECH_NAME FROM CSMS_VW_TECHNICIAN ORDER BY TECH_NAME")

'    Set rsParts = gconDMIS.Execute("SELECT STOCKNO FROM PMIS_STOCKMAS ORDER BY STOCKNO ASC")
'    Dim partno                                         As String
'    While Not rsParts.EOF
'        DoEvents
'        partno = (LTrim(RTrim(Null2String(rsParts!STOCKNO))))
'        cboPARTS.AddItem partno
'        txtCASUAL.AddItem partno
'        rsParts.MoveNext
'    Wend
'    Set rsParts = Nothing
End Sub

Sub fillCboParts()
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT STOCKNO FROM PMIS_STOCKMAS ORDER BY STOCKNO ASC")
    cboPARTS.Clear
    txtCASUAL.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboPARTS.AddItem LTrim(RTrim(Null2String(RSTMP!STOCKNO)))
            txtCASUAL.AddItem LTrim(RTrim(Null2String(RSTMP!STOCKNO)))
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub filltech()

End Sub

Sub FillDLR()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim Item                                           As ListItem

    Set RSTMP = gconDMIS.Execute("SELECT DLR_CQIR_REFERENCENO, ID FROM CSMS_CQIR ORDER BY ID")
    lsvDLR.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set Item = lsvDLR.ListItems.Add(, , Null2String(RSTMP!DLR_CQIR_REFERENCENO))
            Item.SubItems(1) = RSTMP!ID

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub FillParts(VNO As String)
    Dim RSTMP                                          As New ADODB.Recordset
    Dim Item                                           As ListItem
    Dim vAMOUNT                                        As Currency
    Dim VLTS                                           As Double
    Dim VPART                                          As Currency
    Dim VSUBLET                                        As Currency

    VPART = 0: VSUBLET = 0
    vAMOUNT = 0: VLTS = 0
    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_CQIRPARTS WHERE DLR_CQIR_REFERENCENO = '" & VNO & "'")
    lsvPARTS.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set Item = lsvPARTS.ListItems.Add(, , Null2String(RSTMP!partno))
            Item.SubItems(1) = Null2String(RSTMP!partname)
            Item.SubItems(2) = NumericVal(RSTMP!QTY)
            Item.SubItems(3) = Format(NumericVal(RSTMP!cost), "#,###,##0.00")
            Item.SubItems(4) = Null2String(RSTMP!OPCODE)
            Item.SubItems(5) = Format(NumericVal(RSTMP!lts), "#,###,##0.00")
            Item.SubItems(6) = Null2String(RSTMP!ISSUBLET)
            Item.SubItems(7) = RSTMP!ID

            If Null2String(RSTMP!ISSUBLET) = "" Then
                VPART = VPART + N2Str2Zero(RSTMP!cost)
            Else
                VSUBLET = VSUBLET + N2Str2Zero(RSTMP!cost)
            End If

            vAMOUNT = vAMOUNT + Item.SubItems(3)
            VLTS = VLTS + Item.SubItems(5)
            RSTMP.MoveNext
        Loop
    End If

    txtTCOST.Text = Format(VPART, "#,###,##0.00")
    txtLTS.Text = Format(VLTS, "#,###,##0.00")

    'txtLABORCOST.Text = Format(247.5 * txtLTS, "#,###,##0.00")
    txtLABORCOST.Text = Format(lblFLATRATE * txtLTS, "#,###,##0.00")
    'TXTSUBREP.Text = Format(VSUBLET, "#,###,##0.00")
    TXTGRAND.Text = Format(CCur(VPART) + CCur(txtLABORCOST) + CCur(TXTSUBREP.Text), "#,###,##0.00")

    If Not Null2String(RS!Status) = "P" Then
        gconDMIS.Execute ("UPDATE CSMS_CQIR SET TOTALPARTCOST = " & CCur(txtTCOST.Text) & ", TOTALLABORCOST = " & CCur(txtLABORCOST.Text) & ",TOTALSUBLETREPAIR = " & CCur(TXTSUBREP.Text) & ",GRANDTOTAL = " & CCur(TXTGRAND) & " WHERE ID = " & LABID.Caption & "")
    End If

    Set RSTMP = Nothing
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM CSMS_CQIR ORDER BY DLR_CQIR_REFERENCENO desc", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub InitMemVars()
    TabControl1.Tab = 0

    lsvPARTS.ListItems.Clear

    txtPREVACL.Text = ""
    txtPREVRO.Text = ""
    txtSubletType.Text = ""
    txtTCOST.Text = "0.00"
    txtLTS.Text = "0.00"

    txtRO.Text = ""
    txtDLR.Text = ""

    txtModel.Text = ""
    txtDCODE.Text = COMPANY_CODE
    dptDELDATTE.Value = Date
    dtpTranDate.Value = Date
    dptINSD.Value = Date
    dptREPD.Value = Date

    txtKM.Text = ""
    txtPLATENO.Text = ""
    txtPWA.Text = ""
    txtPWAT.Text = ""

    cboSA.Text = ""
    cboTECH.Text = ""

    '    txtDLR.Text = GenerateDLRCODE & xREFNO

    txtCUST.Text = ""
    txtVIN.Text = ""
    txtENGINE.Text = ""
    txtAXLE.Text = ""

    optPWAREQ1.Value = True
    optTRAN1.Value = True
    optATT1.Value = 1

    txtSUBJ.Text = ""

    txtHIST.Text = ""
    txtDESC.Text = ""
    txtANA.Text = ""
    txtCORR.Text = ""
    txtREC.Text = ""
    txtCASUAL.Text = ""
    txtNCODE.Text = ""
    txtCCODE.Text = ""
    cboNCODE.ListIndex = -1
    cboCCODE.ListIndex = -1
    txtNDESC.Text = ""
    txtCDESC.Text = ""

    txtPCODE.Text = ""
    txtSCODE.Text = ""

    txtLABORCOST.Text = "0.00"
    TXTSUBREP.Text = "0.00"
    TXTGRAND.Text = "0.00"
    txtINV.Text = ""
    txtCOMM.Text = ""

    optSENT1.Value = True

    'txtsaleAdvisor.Text = Null2String(RS!serviceAdvisor)
    'cboTECH = Null2String(RS!TECHNICIAN)

    optENG(0).Value = 0: optENG(1).Value = 0: optENG(2).Value = 0
    optWEA(0).Value = 0: optWEA(1).Value = 0: optWEA(2).Value = 0
    optSHIF(0).Value = 0: optSHIF(1).Value = 0
    optPOS(0).Value = 0: optPOS(0).Value = 0
    optMT(0).Value = 0: optMT(1).Value = 0: optMT(2).Value = 0: optMT(3).Value = 0: optMT(4).Value = 0: optMT(5).Value = 0: optMT(6).Value = 0
    optAT(0).Value = 0: optAT(1).Value = 0: optAT(2).Value = 0: optAT(3).Value = 0: optAT(4).Value = 0: optAT(5).Value = 0: optAT(6).Value = 0
    optROAD(0).Value = 0: optROAD(1).Value = 0: optROAD(2).Value = 0: optROAD(3).Value = 0
    optLOC(0).Value = 0: optLOC(1).Value = 0: optLOC(2).Value = 0: optLOC(3).Value = 0
    optACT(0).Value = 0: optACT(1).Value = 0: optACT(2).Value = 0: optACT(3).Value = 0: optACT(4).Value = 0
    optOCC(0).Value = 0: optOCC(1).Value = 0
    optACC(0).Value = 0: optACC(1).Value = 0
    optDEL(0).Value = 1

    txtSPE.Text = ""
    txtEVERY.Text = ""
    txtKMS.Text = ""

    txtREQ.Text = ""
    txtCHECK.Text = ""
    txtAPP.Text = ""
End Sub

Function GetFlatRate() As Double
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT FLATRATE FROM ALL_MAKE WHERE CODE = 'H'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetFlatRate = NumericVal(RSTMP!FLATRATE)
    End If

    Set RSTMP = Nothing
End Function

Sub StoreMemvars()
    If Not RS.EOF And Not RS.BOF Then
        lblTRANNO.Caption = Format(Null2String(RS!TRANNO), "0000")
        LABID.Caption = Null2String(RS!ID)
        'txtFLATRATE.Locked = True
        lblFLATRATE.Caption = Format(NumericVal(RS!FLATRATE), MAXIMUM_DIGIT)
        txtFLATRATE.Text = Format(NumericVal(RS!FLATRATE), MAXIMUM_DIGIT)

        'If Null2String(RS!DELDATE) = "" Then
        '    Check3.Value = 0
        '    dptDELDATTE.Value = CDate(Date)
        'Else
        '    Check3.Value = 1
            dptDELDATTE.Value = Null2Date(RS!DELDATE)
        'End If
        dtpTranDate.Value = Null2String(RS!TRANDATE)
        
        If Null2String(RS!Status) = "" Then
            lblSTATUS.Caption = ""
            cmdPost.Enabled = True: cmdUnPost.Enabled = False
            cmdPrint.Enabled = True: cmdCancelCO.Enabled = True
            cmdEdit.Enabled = True: cmdDelete.Enabled = True
        ElseIf Null2String(RS!Status) = "C" Then
            lblSTATUS.Caption = "** CANCEL **"
            cmdPost.Enabled = False: cmdUnPost.Enabled = False
            cmdPrint.Enabled = False:
            cmdEdit.Enabled = False: cmdDelete.Enabled = False
        ElseIf Null2String(RS!Status) = "P" Then
            lblSTATUS.Caption = "** POSTED **"
            cmdPost.Enabled = False: cmdUnPost.Enabled = True
            cmdPrint.Enabled = True: cmdCancelCO.Enabled = False
            cmdEdit.Enabled = False: cmdDelete.Enabled = False
        ElseIf Null2String(RS!Status) = "W" Then
            lblSTATUS.Caption = "** WAITING FOR CONFIRMATION **"
            cmdPost.Enabled = False: cmdUnPost.Enabled = True
            cmdPrint.Enabled = True:
            cmdEdit.Enabled = True: cmdDelete.Enabled = True
        ElseIf Null2String(RS!Status) = "A" Then
            lblSTATUS.Caption = "** APPROVED CQIR **"
            cmdPost.Enabled = False: cmdUnPost.Enabled = False
            cmdPrint.Enabled = True: cmdCancelCO.Enabled = False
            cmdEdit.Enabled = False: cmdDelete.Enabled = False
        ElseIf Null2String(RS!Status) = "T" Then
            lblSTATUS.Caption = "** ATTACHED TO ACL **"
            cmdPost.Enabled = False: cmdUnPost.Enabled = False
            cmdPrint.Enabled = True: cmdCancelCO.Enabled = False
            cmdEdit.Enabled = False: cmdDelete.Enabled = False
        End If

        If Null2String(RS!pwa_request) = "Y" Then
            optPWAREQ1.Value = True
        Else
            optPWAREQ2.Value = True
        End If

        txtPREVACL.Text = Null2String(RS!PREVACLNO)
        txtPREVRO.Text = Null2String(RS!PREVRONO)
        txtSubletType.Text = Null2String(RS!SUBLETTYPE)
        cboCType.Text = Null2String(RS!CLAIMTYPE)
        txtModel.Text = Null2String(RS!MODEL)
        txtPWA.Text = Null2String(RS!PWANO)
        txtPWAT.Text = Null2String(RS!PWATYPE)

        txtRO.Text = Null2String(RS!RO_NO)
        txtDLR.Text = Null2String(RS!DLR_CQIR_REFERENCENO)

        txtDCODE.Text = Null2String(RS!DEALERCODE)

        'If Null2String(RS!InspectionDate) = "" Then
        '    Check1.Value = 0
        '    dptINSD.Value = CDate(Date)
        'Else
        '    Check1.Value = 1
            dptINSD.Value = Null2Date(RS!InspectionDate)
        'End If


        'If Null2String(RS!RepairDate) = "" Then
        '    Check2.Value = 0
        '    dptREPD.Value = CDate(Date)
        'Else
        '    Check2.Value = 1
            dptREPD.Value = Null2Date(RS!RepairDate)
        'End If

        txtKM.Text = Null2String(RS!MILEAGE)
        txtPLATENO.Text = Null2String(RS!plateno)
        cboSA.Text = Null2String(RS!serviceAdvisor)
        cboTECH.Text = Null2String(RS!Technician)

        txtCUST.Text = Null2String(RS!Customer)
        txtVIN.Text = Null2String(RS!VINNO)
        txtENGINE.Text = Null2String(RS!EngineNo)
        txtAXLE.Text = Null2String(RS!tm_axleno)

        If RS!transmissiontype = "MANUAL" Then
            optTRAN1.Value = True
        Else
            optTRAN2.Value = True
        End If

        If RS!attachments = "PHOTO" Then
            optATT1.Value = True
        Else
            optATT2.Value = True
        End If

        txtSUBJ.Text = Null2String(RS!Subject)
        txtHIST.Text = Null2String(RS!history)
        txtDESC.Text = Null2String(RS!Description)
        txtANA.Text = Null2String(RS!ANALYSIS)
        txtCORR.Text = Null2String(RS!correctiveAction)
        txtCOMM.Text = Null2String(RS!othercomments)
        txtREC.Text = Null2String(RS!RECOMMENDATION)

        txtCASUAL.Text = Null2String(RS!CAUSALPARTNO)
        txtNCODE.Text = Null2String(RS!NATURECODE)
        txtCCODE.Text = Null2String(RS!CAUSECODE)

        cboNCODE.Text = Null2String(RS!NATURECODE)
        cboCCODE.Text = Null2String(RS!CAUSECODE)
        txtNDESC.Text = Null2String(RS!NATUREDESC)
        txtCDESC.Text = Null2String(RS!CAUSEDESC)
        txtPCODE.Text = Null2String(RS!paintcode)
        txtSCODE.Text = Null2String(RS!subletcode)
        TXTSUBREP.Text = Format(NumericVal(RS!TotalSUBLETREPAIR), "#,###,##0.00")

        Call FillParts(Null2String(RS!DLR_CQIR_REFERENCENO))

        txtLABORCOST.Text = Format(NumericVal(RS!TotalLaborCost), "#,###,##0.00")
        TXTSUBREP.Text = Format(NumericVal(RS!TotalSUBLETREPAIR), "#,###,##0.00")
        TXTGRAND.Text = Format(NumericVal(RS!grandtotal), "#,###,##0.00")
        txtINV.Text = Null2String(RS!invoiceno_orno)


        If RS!sentregistrationcard = "YES" Then
            optSENT1.Value = True
        Else
            optSENT2.Value = True
        End If

        'txtsaleAdvisor.Text = Null2String(RS!serviceAdvisor)
        'cboTECH = Null2String(RS!TECHNICIAN)

        'ENGINE TEMPERATURE
        If RS!con_enginetemp = True Then optENG(0).Value = 1 Else optENG(0).Value = 0
        If RS!con_enginetemp1 = True Then optENG(1).Value = 1 Else optENG(1).Value = 0
        If RS!con_enginetemp2 = True Then optENG(2).Value = 1 Else optENG(2).Value = 0

        'WEATHER
        If RS!con_weather = True Then optWEA(0).Value = 1 Else optWEA(0).Value = 0
        If RS!con_weather1 = True Then optWEA(1).Value = 1 Else optWEA(1).Value = 0
        If RS!Con_weather2 = True Then optWEA(2).Value = 1 Else optWEA(2).Value = 0

        'SHIFTING
        If RS!con_shifting = True Then optSHIF(0).Value = 1 Else optSHIF(0).Value = 0
        If RS!con_shifting1 = True Then optSHIF(1).Value = 1 Else optSHIF(0).Value = 0

        'SHIFT POSITION
        If RS!con_shifposition = True Then optPOS(0).Value = 1 Else optPOS(0).Value = 0
        If RS!con_shifposition1 = True Then optPOS(1).Value = 1 Else optPOS(1).Value = 0

        'MANUAL
        If RS!con_MT = True Then optMT(0).Value = 1 Else optMT(0).Value = 0
        If RS!con_MT1 = True Then optMT(1).Value = 1 Else optMT(1).Value = 0
        If RS!con_MT2 = True Then optMT(2).Value = 1 Else optMT(2).Value = 0
        If RS!con_MT3 = True Then optMT(3).Value = 1 Else optMT(3).Value = 0
        If RS!con_MT4 = True Then optMT(4).Value = 1 Else optMT(4).Value = 0
        If RS!con_MT5 = True Then optMT(5).Value = 1 Else optMT(5).Value = 0
        If RS!con_MT6 = True Then optMT(6).Value = 1 Else optMT(6).Value = 0

        'AUTOMATIC
        If RS!con_AT = True Then optAT(0).Value = 1 Else optAT(0).Value = 0
        If RS!con_AT1 = True Then optAT(1).Value = 1 Else optAT(1).Value = 0
        If RS!con_AT2 = True Then optAT(2).Value = 1 Else optAT(2).Value = 0
        If RS!con_AT3 = True Then optAT(3).Value = 1 Else optAT(3).Value = 0
        If RS!con_AT4 = True Then optAT(4).Value = 1 Else optAT(4).Value = 0
        If RS!con_AT5 = True Then optAT(5).Value = 1 Else optAT(5).Value = 0
        If RS!con_AT6 = True Then optAT(6).Value = 1 Else optAT(6).Value = 0

        'ROAD CONDITION
        If RS!con_road = True Then optROAD(0).Value = 1 Else optROAD(0).Value = 0
        If RS!con_road1 = True Then optROAD(1).Value = 1 Else optROAD(1).Value = 0
        If RS!con_road2 = True Then optROAD(2).Value = 1 Else optROAD(2).Value = 0
        If RS!con_road3 = True Then optROAD(3).Value = 1 Else optROAD(3).Value = 0

        'LOCATION
        If RS!con_location = True Then optLOC(0).Value = 1 Else optLOC(0).Value = 0
        If RS!con_location = True Then optLOC(1).Value = 1 Else optLOC(1).Value = 0
        If RS!con_location = True Then optLOC(2).Value = 1 Else optLOC(2).Value = 0
        If RS!con_location = True Then optLOC(3).Value = 1 Else optLOC(3).Value = 0

        'ACTION
        If RS!con_action = True Then optACT(0).Value = 1 Else optACT(0).Value = 0
        If RS!con_action1 = True Then optACT(1).Value = 1 Else optACT(1).Value = 0
        If RS!con_action2 = True Then optACT(2).Value = 1 Else optACT(2).Value = 0
        If RS!con_action3 = True Then optACT(3).Value = 1 Else optACT(3).Value = 0
        If RS!con_action4 = True Then optACT(4).Value = 1 Else optACT(4).Value = 0

        'OCCURENCE
        If RS!con_Occurence = True Then optOCC(0).Value = 1 Else optOCC(0).Value = 0
        If RS!con_Occurence1 = True Then optOCC(1).Value = 1 Else optOCC(1).Value = 0

        'ACCESSORIES
        If RS!con_accessories = True Then optACC(0).Value = 1 Else optACC(0).Value = 0
        If RS!con_accessories1 = True Then optACC(1).Value = 1 Else optACC(1).Value = 0

        If RS!VechicleMaintenance = "DEALER" Then optDEL(0).Value = True
        If RS!VechicleMaintenance = "3-STAR SHOP" Then optDEL(1).Value = True
        If RS!VechicleMaintenance = "GAS STATION" Then optDEL(2).Value = True
        If RS!VechicleMaintenance = "OTHERS" Then optDEL(3).Value = True

        txtSPE.Text = Null2String(RS!othercondition)
        txtEVERY.Text = Null2String(RS!everymonth)
        txtKMS.Text = Null2String(RS!everykms)

        txtREQ.Text = Null2String(RS!requestedby)
        txtCHECK.Text = Null2String(RS!CheckedBy)
        txtAPP.Text = Null2String(RS!ApprovedBy)
    Else
        ShowNoRecord
        cmdAdd_Click
    End If


    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub

End Sub

Sub FindVin(vPLATENO As String)
    Dim RSMAKOY                                        As New ADODB.Recordset

    Set RSMAKOY = gconDMIS.Execute("SELECT VIN,ENGINE FROM CSMS_CUSVEH WHERE PLATE_NO = '" & vPLATENO & "'")
    If Not (RSMAKOY.BOF And RSMAKOY.EOF) Then
        txtVIN.Text = Null2String(RSMAKOY!VIN)
        txtENGINE.Text = Null2String(RSMAKOY!Engine)
    End If

    Set RSMAKOY = Nothing
End Sub

Sub FINDREPOR(vREPOR As String)
    Dim RSTMP                                          As New ADODB.Recordset
    Dim RSDER                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE REP_OR = '" & vREPOR & "'")
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        txtCUST.Text = Null2String(RSTMP!NIYM)
        txtKM.Text = Null2String(RSTMP!km_rdg)
        txtPLATENO.Text = Null2String(RSTMP!PLATE_NO)
        dptREPD.Value = CDate(RSTMP!DTE_RECD)
        cboSA.Text = FindSAName(RSTMP!RECD_BY)
        txtVIN.Text = Null2String(RSTMP!VIN)
        txtINV.Text = Format(Null2String(RSTMP!INVOICE), "00000000")

        Call FindVin(txtPLATENO)
    Else
        MsgBox "Repair Order no. Cannot be Found", vbInformation, "CSMS"
        Exit Sub
    End If

    Set RSTMP = Nothing
End Sub

Sub FillInfo()
    Dim INDEX                                          As Long
    Dim vID                                            As Long
    Dim vRONO                                          As String

    vRONO = txtRO
    Call FINDREPOR(vRONO)
End Sub

Sub FillSearchGridRepor(XXX As String)
    Dim RSTMP                                          As New ADODB.Recordset
    lsvDLR.Sorted = False: lsvDLR.ListItems.Clear
    lsvDLR.Enabled = False
    Set RSTMP = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))

    Set RSTMP = gconDMIS.Execute("select DLR_CQIR_REFERENCENO, ID from CSMS_CQIR where DLR_CQIR_REFERENCENO Like '%" & XXX & "%' ORDER BY DLR_CQIR_REFERENCENO")
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        Listview_Loadval Me.lsvDLR.ListItems, RSTMP
        lsvDLR.Refresh
        lsvDLR.Enabled = True
    End If
    Set RSTMP = Nothing
End Sub

Sub FillGridRepor()
    Dim RSTMP                                          As New ADODB.Recordset
    lsvDLR.Enabled = False
    lsvDLR.Sorted = False: lsvDLR.ListItems.Clear
    Set RSTMP = New ADODB.Recordset


    Set RSTMP = gconDMIS.Execute("select DLR_CQIR_REFERENCENO, ID from CSMS_CQIR Order by DLR_CQIR_REFERENCENO")

    If Not (RSTMP.EOF And RSTMP.BOF) Then
        Listview_Loadval Me.lsvDLR.ListItems, RSTMP
        lsvDLR.Refresh
        lsvDLR.Enabled = True
    End If

    Set RSTMP = Nothing
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSTMP                                          As New ADODB.Recordset
    lsvDLR.Sorted = False: lsvDLR.ListItems.Clear
    lsvDLR.Enabled = False
    Set RSTMP = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    If XXX = "" Then
        Set RSTMP = gconDMIS.Execute("select TOP 30 DLR_CQIR_REFERENCENO, ID from CSMS_CQIR Order by DLR_CQIR_REFERENCENO desc")
    Else
        If cboSearchBy.ListIndex = 0 Then
            Set RSTMP = gconDMIS.Execute("select TOP 30 DLR_CQIR_REFERENCENO, ID from CSMS_CQIR where DLR_CQIR_REFERENCENO Like '%" & XXX & "%' ORDER BY DLR_CQIR_REFERENCENO desc")
        ElseIf cboSearchBy.ListIndex = 1 Then
            Set RSTMP = gconDMIS.Execute("select TOP 30 VINNO, ID from CSMS_CQIR where VINNO Like '%" & XXX & "%' ORDER BY DLR_CQIR_REFERENCENO desc")
        ElseIf cboSearchBy.ListIndex = 2 Then
            Set RSTMP = gconDMIS.Execute("select TOP 30 RO_NO, ID from CSMS_CQIR where RO_NO Like '%" & XXX & "%' ORDER BY DLR_CQIR_REFERENCENO desc")
        Else
            Set RSTMP = gconDMIS.Execute("select TOP 30 CUSTOMER, ID from CSMS_CQIR where CUSTOMER Like '%" & XXX & "%' ORDER BY DLR_CQIR_REFERENCENO desc")
        End If
    End If
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        Listview_Loadval Me.lsvDLR.ListItems, RSTMP
        lsvDLR.Refresh
        lsvDLR.Enabled = True
    End If
    Set RSTMP = Nothing
End Sub

Sub ReportControlAddColumnHeader(LST As ReportControl, StringHeaders As String)
    Dim ar()                                           As String
    Dim I                                              As Integer

    ar = Split(StringHeaders, ",")
    LST.Columns.DeleteAll
    For I = LBound(ar) To UBound(ar)
        LST.Columns.Add I, ar(I), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Sub ReportControlPaintManager(LST As ReportControl)
    With LST
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnExplorer
    End With
End Sub

Private Sub cboCCODE_Change()
    If ADD_OR_EDIT = "" Then Exit Sub
    txtCDESC = FindCodeDesc("WWTCAUE", "CAUEENGL", "CAUECODE", cboCCODE.Text)
End Sub

Private Sub cboCCODE_Click()
    If ADD_OR_EDIT = "" Then Exit Sub
    txtCDESC = FindCodeDesc("WWTCAUE", "CAUEENGL", "CAUECODE", cboCCODE.Text)
End Sub

Private Sub cboNCODE_Change()
    If ADD_OR_EDIT = "" Then Exit Sub
    txtNDESC = FindCodeDesc("WWTNATR", "NATRENGL", "NATRCODE", cboNCODE.Text)
End Sub

Private Sub cboNCODE_Click()
    txtNDESC = FindCodeDesc("WWTNATR", "NATRENGL", "NATRCODE", cboNCODE.Text)
End Sub


Private Sub cboPARTS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call frm.SelectSQl("select top 100 STOCKNO, STOCKDESC from PMIS_STOCKMAS WHERE STOCKNO LIKE", "PARTS")
        frm.Show 1
    End If
End Sub

Private Sub cboPARTS_LostFocus()
    On Error Resume Next
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT STOCKDESC FROM PMIS_STOCKMAS WHERE LTRIM(RTRIM(STOCKNO)) = '" & RTrim(LTrim(cboPARTS.Text)) & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        txtPartDesc.Text = Null2String(RSTMP!STOCKDESC)
    Else
        txtPartDesc.Text = "Part no. not found"
    End If
    Set RSTMP = Nothing
End Sub

Private Sub cmd1_Click()
    'If COMPANY_CODE = "HBK" Then Exit Sub
    DisabledSaveMenu False

    Call DisplayWarrantyRO
    picSEARCHI.Visible = True
    picSEARCHI.ZOrder 0
End Sub

Private Sub cmd3_Click()
    picSEARCHI.Visible = False
    picSEARCHI.ZOrder 1

    Call DisabledSaveMenu(True)
    picSEARCH.Enabled = False
End Sub

Private Sub cmd4_Click()
    If Option4.Value = True Then                      'EXCEL PRINTING
        Call PRINTEXCEL
    Else

    End If
    gconDMIS.Execute ("UPDATE CSMS_CQIR SET STATUS = 'W' WHERE ID = " & LABID & "")
    cmd6_Click

    rsRefresh
    RS.Find "ID = " & LABID & ""
    StoreMemvars
End Sub

Private Sub cmd6_Click()
    picPRINT.Visible = False
    picPRINT.ZOrder 1
End Sub

Private Sub cmd7_Click()
    picSeachPlateNo.Visible = False
    picSeachPlateNo.ZOrder 1

    Call DisabledSaveMenu(True)
    picSEARCH.Enabled = False
End Sub

Private Sub cmd8_Click()
    'If COMPANY_CODE = "HBK" Then Exit Sub

    DisabledSaveMenu False

    Call DisplayVEHICLES
    picSeachPlateNo.Visible = True
    picSeachPlateNo.ZOrder 0
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "QUALITY INFORMATION REPORT") = False Then Exit Sub
    Dim xREFNO                                         As String

    ADD_OR_EDIT = "ADD"

    lblTRANNO.Caption = GenerateNewTranno
    Call InitMemVars

    Dim RSTMP                                          As New ADODB.Recordset
    Dim COCODE                                         As String
    Set RSTMP = gconDMIS.Execute("SELECT COMPANYCODE FROM ALL_PROFILE WHERE MODULENAME = 'CSMS'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        COCODE = Null2String(RSTMP!COMPANYCODE)
    End If
    Set RSTMP = Nothing

    xREFNO = GENERATEREFNO
    If COMPANY_CODE = "HGC" Then
        txtDLR.Text = Right(Year(Date), 2) + "-" + Format(Month(Date), "00") + "-" + Format(Day(Date), "00") + "-" + Format(xREFNO, "000")
    Else
        txtDLR.Text = Right(Year(Date), 2) + Format(Month(Date), "00") + Format(xREFNO, "00000")
    End If
    txtDLR.Text = COCODE + txtDLR

    picMENU.Visible = False
    Call DisabledPics(False)
    picAdds.Visible = False
    picSaves.Visible = True
    TabControl1.Tab = 0
End Sub

Private Sub cmdCancel_Click()
    picMENU.Visible = True
    picSaves.Visible = False
    picAdds.Visible = True
    DisabledPics True

    StoreMemvars
End Sub

Private Sub cmdCloseA_Click()
    picAUTHOR.Visible = False
    picAUTHOR.ZOrder 1
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "QUALITY INFORMATION REPORT") = False Then Exit Sub

    If MsgBox("Cancel this Transaction", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

    Dim xROID                                          As String
    If Not txtRO.Text = "" Then
        xROID = FindTransactionID(N2Str2Null(txtRO), "REP_OR", "CSMS_REPOR")
    Else
        xROID = ""
    End If

    SQL_STATEMENT = "UPDATE CSMS_CQIR SET STATUS = 'C' WHERE ID = " & LABID & ""
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT --------------------------------------------------------------------------------
    Call NEW_LogAudit("C", "QUALITY INFORMATION", SQL_STATEMENT, LABID.Caption, "", "DLR no. " & txtDLR.Text & " ,RO no: " & txtRO, "", xROID)
    Call NEW_LogAudit("DT", "BILLING SYSTEM", SQL_STATEMENT, xROID, "", "DLR no. " & txtDLR.Text & " ,RO no: " & txtRO, "", "")
    'NEW LOG AUDIT --------------------------------------------------------------------------------


    txtSEARCH.Text = "A": txtSEARCH.Text = ""
    rsRefresh
    RS.MoveFirst
    RS.Find "id = " & LABID & ""
    StoreMemvars
End Sub

Private Sub cmdCancelPWA_Click()
    Call DisabledPics(True)
    picPWA.Visible = False
    picPWA.ZOrder 1
End Sub

Private Sub cmdCloseAP_Click()
    DisabledSaveMenu True

    picNOTES.Enabled = False
    picCONDITIONS.Enabled = False
    picHD.Enabled = False

    picMENU.Visible = True
    picAddPart.Visible = False
    picAddPart.ZOrder 1
End Sub

Private Sub cmdCloseAS_Click()
    DisabledSaveMenu True

    picNOTES.Enabled = False
    picCONDITIONS.Enabled = False
    picHD.Enabled = False

    picMENU.Visible = True
    picAddSub.Visible = False
    picAddSub.ZOrder 1
End Sub

Private Sub cmdDEF_Click()
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT TOP 1 REQUESTEDBY,CHECKEDBY FROM CSMS_CQIR ORDER BY ID")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        txtREQ.Text = Null2String(RSTMP!requestedby)
        txtCHECK.Text = Null2String(RSTMP!CheckedBy)
    End If
    Set RSTMP = Nothing
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "QUALITY INFORMATION REPORT") = False Then Exit Sub

    If MsgBox("Delete this Transaction", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

    gconDMIS.Execute ("DELETE FROM CSMS_CQIR WHERE ID = " & LABID & "")
    gconDMIS.Execute ("DELETE FROM CSMS_CQIRPARTS WHERE DLR_CQIR_REFERENCENO = '" & txtDLR.Text & "'")


    'NEW LOG AUDIT --------------------------------------------------------------------------------
    Dim xROID                                          As String
    If Not txtRO.Text = "" Then
        xROID = FindTransactionID(txtRO, "REP_OR", "CSMS_REPOR")
    Else
        xROID = ""
    End If

    AUDIT_SQL = "DELETE FROM CSMS_CQIR WHERE ID = " & LABID & ""
    Call NEW_LogAudit("X", "QUALITY INFORMATION", AUDIT_SQL, LABID.Caption, "", "DLR no. " & txtDLR.Text & " ,RO no: " & txtRO, "", xROID)
    Call NEW_LogAudit("DT", "BILLING SYSTEM", AUDIT_SQL, xROID, "", "DLR no. " & txtDLR.Text & " ,RO no: " & txtRO, "", "")

    AUDIT_SQL = "DELETE FROM CSMS_CQIRPARTS WHERE DLR_CQIR_REFERENCENO = '" & txtDLR.Text & "'"
    Call NEW_LogAudit("XX", "QUALITY INFORMATION", AUDIT_SQL, "ALL DETAILS", "", "DLR no. " & txtDLR.Text, "", "ALL DETAILS")
    'NEW LOG AUDIT --------------------------------------------------------------------------------


    txtSEARCH.Text = "A": txtSEARCH.Text = ""
    ShowDeletedMsg

    rsRefresh
    StoreMemvars
End Sub

Private Sub cmdDeleteJob_Click()
    If MsgBox("Delete This Job", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

    gconDMIS.Execute ("DELETE FROM CSMS_CQIRPARTS WHERE ID = " & lblLITEMNO.Caption & "")
    ShowDeletedMsg

    FillParts txtDLR
    cmsCloseAJ_Click
End Sub

Private Sub cmdDeletePart_Click()
    If Function_Access(LOGID, "Acess_DELETE", "QUALITY INFORMATION") = False Then Exit Sub

    If MsgBox("Remove This Part no", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

    AUDIT_SQL = "DELETE FROM CSMS_CQIRPARTS WHERE ID = " & lblPITEMNO.Caption & ""
    gconDMIS.Execute (AUDIT_SQL)

    Call ShowDeletedMsg
    'NEW LOG AUDIT--------------------------------------------------------------------------------
    Call NEW_LogAudit("XX", "QUALITY INFORMATION", AUDIT_SQL, LABID.Caption, "P", "DLR no: " & txtDLR.Text & " ,PART no: " & txtPartNO, "", lblPITEMNO.Caption)
    'NEW LOG AUDIT--------------------------------------------------------------------------------

    FillParts txtDLR
    cmdCloseAP_Click
End Sub

Private Sub cmdDeleteSublet_Click()
    If Function_Access(LOGID, "Acess_DELETE", "QUALITY INFORMATION") = False Then Exit Sub

    If MsgBox("Delete Sublet Job", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

    AUDIT_SQL = "DELETE FROM CSMS_CQIRPARTS WHERE ID = " & lblSITEMNO.Caption & ""
    gconDMIS.Execute (AUDIT_SQL)
    ShowDeletedMsg

    'NEW LOG AUDIT-----------------------------------------------------------------------------
    Call NEW_LogAudit("XX", "QUALITY INFORMATION", AUDIT_SQL, LABID.Caption, "S", "DLR no:" & txtDLR.Text & " ,SUB code: " & txtSubCode, "", lblSITEMNO)
    'NEW LOG AUDIT-----------------------------------------------------------------------------

    FillParts txtDLR
    cmdCloseAS_Click
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "QUALITY INFORMATION REPORT") = False Then Exit Sub
    ADD_OR_EDIT = "EDIT"

    picMENU.Visible = False
    picAdds.Visible = False
    picSaves.Visible = True
    Call DisabledPics(False)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    RS.MoveFirst
    StoreMemvars
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    RS.MoveLast
    StoreMemvars
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_POST", "QUALITY INFORMATION REPORT") = False Then Exit Sub

    If MsgBox("Post This Transaction", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub
    AUDIT_SQL = "UPDATE CSMS_CQIR SET STATUS = 'P' WHERE ID = " & LABID.Caption & ""
    gconDMIS.Execute (AUDIT_SQL)

    'NEW LOG AUDIT--------------------------------------------------------------------------------
    Call NEW_LogAudit("P", "QUALITY INFORMATION", AUDIT_SQL, LABID, "", "DLR no: " & txtDLR, "", "")
    'NEW LOG AUDIT--------------------------------------------------------------------------------

    ShowSuccessFullyUpdated

    rsRefresh
    RS.Find "ID = " & LABID.Caption & ""
    StoreMemvars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "QUALITY INFORMATION REPORT") = False Then Exit Sub

    If Option4.Value = True Then                      'EXCEL PRINTING
        Screen.MousePointer = 11
        Call PRINTEXCEL

        'NEW LOG AUDIT----------------------------------------------------------------------------
        Call NEW_LogAudit("V", "QUALITY INFORMATION", "", LABID, "", "DLR NO. " & txtDLR, "", "")
        'NEW LOG AUDIT----------------------------------------------------------------------------
        Screen.MousePointer = 0
    End If

    Call cmd6_Click

    Call rsRefresh
    RS.Find "ID = " & LABID & ""
    Call StoreMemvars
End Sub

Private Sub cmdRefresh_Click()
    rsRefresh
    RS.Find "ID = " & LABID.Caption & ""
    StoreMemvars
End Sub

Private Sub cmdSaveJobs_Click()
    On Error Resume Next
    Dim SDESC                                          As String
    Dim SCOST                                          As Double
    Dim VDLRNO                                         As String

    If txtJobDesc.Text = "" Then
        ShowIsRequiredMsg "JOB DESCRIPTION CANNOT BE BLANK"
        txtJobDesc.SetFocus
        Exit Sub
    End If

    If txtJobCost.Text = "" Then
        ShowIsRequiredMsg "JOB COST CANNOT BE BLANK"
        txtJobCost.SetFocus
        Exit Sub
    End If

    If IsNumeric(txtJobCost) = False Then
        MsgBox "INVALID COST FORMAT", vbExclamation, "CSMS"
        txtJobCost.SetFocus
        Exit Sub
    End If

    SDESC = N2Str2Null(txtJobDesc)
    SCOST = NumericVal(txtJobCost)
    VDLRNO = N2Str2Null(txtDLR)

    If ADD_OR_EDIT_ITEM = "ADD" Then
        AUDIT_SQL = "INSERT INTO CSMS_CQIRPARTS (DLR_CQIR_REFERENCENO,PARTNAME,QTY, COST,LTS,ISSUBLET) " & _
                  " VALUES (" & VDLRNO & "," & SDESC & "," & 1 & "," & SCOST & "," & 0 & ",'J')"
        gconDMIS.Execute (AUDIT_SQL)

        ShowSuccessFullyAdded

        'NEW LOG AUDIT----------------------------------------------------------------------------
        Call NEW_LogAudit("AA", "QUALITY INFORMATION", AUDIT_SQL, LABID, "P", "DLR NO. " & txtDLR & " ,PART NO. " & txtPartNO, "", "")
        'NEW LOG AUDIT----------------------------------------------------------------------------
    Else
        AUDIT_SQL = "UPDATE CSMS_CQIRPARTS SET PARTNAME = " & SDESC & _
                    ",COST = " & SCOST & " WHERE ID = " & lblLITEMNO & ""
        gconDMIS.Execute (AUDIT_SQL)

        ShowSuccessFullyUpdated

        'NEW LOG AUDIT----------------------------------------------------------------------------
        Call NEW_LogAudit("EE", "QUALITY INFORMATION", AUDIT_SQL, LABID, "P", "DLR NO. " & txtDLR, "", lblPITEMNO)
        'NEW LOG AUDIT----------------------------------------------------------------------------
    End If

    FillParts txtDLR
    cmsCloseAJ_Click
End Sub

Private Sub cmdSavePart_Click()
    On Error Resume Next
    Dim VPARTNO                                        As String
    Dim SDESC                                          As String
    Dim SCOST                                          As Double
    Dim VQTY                                           As Double
    Dim VOPCODE                                        As String
    Dim VLTS                                           As String
    Dim VDLRNO                                         As String

    Dim RSTMP                                          As New ADODB.Recordset

    '    If optMAN.Value = True Then
    '        Set rstmp = gconDMIS.Execute("SELECT STOCKDESC FROM PMIS_STOCKMAS WHERE STOCKNO = '" & Trim(txtPartNO.Text) & "'")
    '        If (rstmp.BOF And rstmp.EOF) Then
    '            MsgBox "Invalid Part no.", vbInformation, "CSMS"
    '            txtPartNO.SetFocus
    '            Exit Sub
    '        End If
    '        Set rstmp = Nothing
    '
    '        Set rstmp = gconDMIS.Execute("SELECT STOCKDESC FROM PMIS_STOCKMAS WHERE LTRIM(RTRIM(STOCKNO)) = '" & LTrim(RTrim(cboPARTS.Text)) & "'")
    '        If (rstmp.BOF And rstmp.EOF) Then
    '            ShowIsRequiredMsg "INVALID PART NO."
    '            cboPARTS.SetFocus
    '            Exit Sub
    '        End If
    '        Set rstmp = Nothing
    '
    '        If cboPARTS.Text = "" Then
    '            ShowIsRequiredMsg "PART NUMBER CANNOT BE BLANK"
    '            cboPARTS.SetFocus
    '            Exit Sub
    '        End If
    '    End If

    '    If optMAN.Value = False Then
    'If txtPartNo.Text = "" Then
    If cboPARTS.Text = "" Then
        ShowIsRequiredMsg "PART NUMBER CANNOT BE BLANK"
        'txtPartNo.SetFocus
        cboPARTS.SetFocus
        Exit Sub
    End If
    '    End If



    If txtPartDesc.Text = "" Then
        ShowIsRequiredMsg "PART DESCRIPTION CANNOT BE BLANK"
        txtPartDesc.SetFocus
        Exit Sub
    End If

    If txtQty.Text = "" Then
        ShowIsRequiredMsg "QTY CANNOT BE BLANK"
        txtQty.SetFocus
        Exit Sub
    End If

    If IsNumeric(txtQty) = False Then
        MsgBox "INVALID COST FORMAT", vbExclamation, "CSMS"
        txtQty.SetFocus
        Exit Sub
    End If

    If txtPartCost.Text = "" Then
        MsgBox "INVALID COST FORMAT", vbExclamation, "CSMS"
        txtPartCost.SetFocus
        Exit Sub
    End If

    If IsNumeric(txtPartCost) = False Then
        MsgBox "INVALID COST FORMAT", vbExclamation, "CSMS"
        txtPartCost.SetFocus
        Exit Sub
    End If

    '    If txtOPCODE.Text = "" Then
    '        ShowIsRequiredMsg "OPCODE CANNOT BE BLANK"
    '        txtOPCODE.SetFocus
    '        Exit Sub
    '    End If

    '    If txtALTS.Text = "" Then
    '        ShowIsRequiredMsg "LTS CANNOT BE BLANK"
    '        txtALTS.SetFocus
    '        Exit Sub
    '    End If

    If IsNumeric(txtALTS.Text) = False Then
        MsgBox "INVALID LTS FORMAT", vbExclamation, "CSMS"
        txtALTS.SetFocus
        Exit Sub
    End If

    '    If optMAN.Value = False Then
    '        VPARTNO = N2Str2Null(txtPartNo)
    '    Else
    VPARTNO = N2Str2Null(cboPARTS)
    '    End If

    SDESC = N2Str2Null(txtPartDesc)
    VQTY = NumericVal(txtQty)
    SCOST = NumericVal(txtPartCost)
    VOPCODE = N2Str2Null(txtOPCODE)
    VLTS = NumericVal(txtALTS)
    VDLRNO = N2Str2Null(txtDLR)

    If ADD_OR_EDIT_ITEM = "ADD" Then
        AUDIT_SQL = "INSERT INTO CSMS_CQIRPARTS (DLR_CQIR_REFERENCENO,PARTNO,PARTNAME,QTY,COST,OPCODE,LTS) " & _
                  " VALUES (" & VDLRNO & _
                    "," & VPARTNO & _
                    "," & SDESC & _
                    "," & VQTY & _
                    "," & SCOST & _
                    "," & VOPCODE & _
                    "," & VLTS & ")"
        gconDMIS.Execute (AUDIT_SQL)

        ShowSuccessFullyAdded

        'NEW LOG AUDIT-----------------------------------------------------------------------------
        Call NEW_LogAudit("AA", "QUALITY INFORMATION", AUDIT_SQL, LABID.Caption, "P", "QIR no: " & txtDLR.Text & " ,PART no: " & cboPARTS, "", "")
        'NEW LOG AUDIT-----------------------------------------------------------------------------
    Else
        AUDIT_SQL = "UPDATE CSMS_CQIRPARTS SET " & _
                  " PARTNO = " & VPARTNO & _
                    ",PARTNAME = " & SDESC & _
                    ",QTY = " & VQTY & _
                    ",COST = " & SCOST & _
                    ",OPCODE = " & VOPCODE & _
                    ",LTS = " & VLTS & _
                  " WHERE ID = " & lblPITEMNO & ""
        gconDMIS.Execute (AUDIT_SQL)

        ShowSuccessFullyUpdated
        'NEW LOG AUDIT-----------------------------------------------------------------------------
        Call NEW_LogAudit("EE", "QUALITY INFORMATION", AUDIT_SQL, LABID.Caption, "P", "QIR no: " & txtDLR.Text & " ,PART no: " & cboPARTS, "", lblPITEMNO.Caption)
        'NEW LOG AUDIT-----------------------------------------------------------------------------
    End If

    FillParts txtDLR
    cmdCloseAP_Click
End Sub

Private Sub cmdSavePWA_Click()
    Dim SQL                                            As String
    Dim vPWANO                                         As String
    Dim VAPPROVEDBY                                    As String

    If txtPWAno1.Text = "" Then
        ShowIsRequiredMsg ("PWA NO CANNOT BE BLANK")
        txtPWAno1.SetFocus
        Exit Sub
    End If

    If txtAPP1.Text = "" Then
        ShowIsRequiredMsg "APPROVED BY CANNOT BE BLANK"
        txtAPP1.SetFocus
        Exit Sub
    End If

    vPWANO = N2Str2Null(txtPWAno1.Text)
    VAPPROVEDBY = N2Str2Null(txtAPP1.Text)

    AUDIT_SQL = "UPDATE CSMS_CQIR SET STATUS = 'A',PWANO = " & vPWANO & _
                ",APPROVEDBY = " & VAPPROVEDBY & ",DATEAPPROVED = '" & Date & "' WHERE ID = " & LABID & ""
    gconDMIS.Execute (AUDIT_SQL)

    'NEW LOG AUDIT---------------------------------------------------------------------------------
    Call NEW_LogAudit("AP", "QUALITY INFORMATION", AUDIT_SQL, LABID, "", "DLR no: " & txtDLR.Text & " ,PWA no: " & txtPWAno1, "", "")
    'NEW LOG AUDIT---------------------------------------------------------------------------------

    MessagePop InfoFriend, "QIR Information Updated", "QIR Sucessfully Approved!", 1000
    Call rsRefresh
    RS.Find "ID = " & LABID & ""

    Call cmdCancelPWA_Click
    Call StoreMemvars
End Sub

Private Sub cmdSaveSub_Click()
    On Error Resume Next
    Dim SDESC                                          As String
    Dim SCOST                                          As Double
    Dim VDLRNO                                         As String
    Dim SSUBCODE                                       As String
    Dim SOPCODE                                        As String
    Dim SLTS                                           As String
    Dim SUBQTY                                         As Integer

    If txtSubletDesc.Text = "" Then
        ShowIsRequiredMsg "SUBLET DESCRIPTION CANNOT BE BLANK"
        txtSubletDesc.SetFocus
        Exit Sub
    End If

    '    If txtSCOST.Text = "" Then
    '        ShowIsRequiredMsg "SUBLET COST CANNOT BE BLANK"
    '        txtSCOST.SetFocus
    '        Exit Sub
    '    End If

    '    If IsNumeric(txtSCOST) = False Then
    '        MsgBox "INVALID COST FORMAT", vbExclamation, "CSMS"
    '        txtSCOST.SetFocus
    '        Exit Sub
    '    End If

    SSUBCODE = N2Str2Null(txtSubCode.Text)
    SDESC = N2Str2Null(txtSubletDesc)
    SUBQTY = NumericVal(txtSUBQTY)
    SCOST = NumericVal(txtSCOST)
    SOPCODE = N2Str2Null(txtSOPCODE.Text)
    SLTS = N2Str2Zero(txtSLTS.Text)
    VDLRNO = N2Str2Null(txtDLR)


    If ADD_OR_EDIT_ITEM = "ADD" Then
        AUDIT_SQL = "INSERT INTO CSMS_CQIRPARTS (DLR_CQIR_REFERENCENO, PARTNO, PARTNAME, QTY, COST, OPCODE, LTS, ISSUBLET) " & _
                  " VALUES (" & VDLRNO & "," & SSUBCODE & "," & SDESC & "," & SUBQTY & "," & SCOST & "," & SOPCODE & "," & SLTS & ",'S')"
        gconDMIS.Execute (AUDIT_SQL)

        ShowSuccessFullyAdded
        'NEW LOG AUDIT-----------------------------------------------------------------------------
        Call NEW_LogAudit("AA", "QUALITY INFORMATION", AUDIT_SQL, LABID.Caption, "S", "QIR no: " & txtDLR.Text & ",SUB code: " & txtSubCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------------------------------
    Else
        AUDIT_SQL = "UPDATE CSMS_CQIRPARTS SET PARTNAME = " & SDESC & _
                    ", COST = " & SCOST & _
                    ", QTY = " & SUBQTY & _
                    ", PARTNO = " & SSUBCODE & _
                    ", OPCODE = " & SOPCODE & _
                    ", LTS = " & SLTS & _
                    ", DLR_CQIR_REFERENCENO = " & VDLRNO & _
                  " WHERE ID = " & lblSITEMNO & ""
        gconDMIS.Execute (AUDIT_SQL)

        ShowSuccessFullyUpdated
        'NEW LOG AUDIT-----------------------------------------------------------------------------
        Call NEW_LogAudit("EE", "QUALITY INFORMATION", AUDIT_SQL, LABID.Caption, "S", "QIR no: " & txtDLR.Text & ",SUB code: " & txtSubCode, "", lblSITEMNO.Caption)
        'NEW LOG AUDIT-----------------------------------------------------------------------------
    End If

    FillParts txtDLR
    cmdCloseAS_Click
End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UNPOST", "QUALITY INFORMATION REPORT") = False Then Exit Sub

    If MsgBox("Unpost This Transaction", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

    AUDIT_SQL = "UPDATE CSMS_CQIR SET STATUS = NULL, PWANO = NULL, APPROVEDBY = NULL WHERE ID = " & LABID & ""
    gconDMIS.Execute (AUDIT_SQL)

    'NEW LOG AUDIT--------------------------------------------------------------------------------
    Call NEW_LogAudit("U", "QUALITY INFORMATION", AUDIT_SQL, LABID, "", "DLR no: " & txtDLR, "", "")
    'NEW LOG AUDIT--------------------------------------------------------------------------------

    Call rsRefresh
    RS.Find "ID = " & LABID & ""

    Call StoreMemvars
End Sub

Private Sub cmsCloseAJ_Click()
    Call DisabledSaveMenu(True)
    picMENU.Visible = True
    picAddJob.Visible = False
    picAddJob.ZOrder 1
End Sub

Private Sub Command1_Click()
    'picCOND.Visible = True
    'picCOND.ZOrder 0
End Sub

Private Sub Command2_Click()
    picAUTHOR.Visible = True
    picAUTHOR.ZOrder 0
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        Call ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        Call ShowLastRecordMsg
    End If
    Call StoreMemvars
End Sub

Private Sub cmdSave_Click()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim VDLRNO                                         As String
    Dim VDELDATE As String: Dim VNAME As String: Dim vVin As String: Dim VENGINENO As String
    Dim VAXLE As String: Dim VTRANS As String: Dim VATT As String: Dim vDCODE As String
    Dim VINSDATE As String: Dim VREPDATE As String: Dim vRO As String: Dim VKM As String
    Dim VPLATE As String: Dim VSENT As String: Dim VSA As String: Dim VTECH As String
    Dim VSUBJ As String: Dim VHIST As String: Dim VDESC As String: Dim VANA As String
    Dim VCORR As String: Dim VOTHER As String: Dim VREC As String: Dim VCAS As String
    Dim VCCODE As String: Dim VPCODE As String: Dim VSCODE As String:
    Dim VENGTEMP As String: Dim VENGTEMP1 As String: Dim VENGTEMP2 As String
    Dim VWEA As String: Dim VWEA1 As String: Dim VWEA2 As String:
    Dim VSHIFT As String: Dim VSHIFT1                  As String:
    Dim vPOSITION As String: Dim vPOSITION1            As String
    Dim VMT As String: Dim VMT1 As String: Dim VMT2 As String: Dim VMT3 As String: Dim VMT4 As String: Dim VMT5 As String: Dim VMT6 As String
    Dim VAT As String: Dim VAT1 As String: Dim VAT2 As String: Dim VAT3 As String: Dim VAT4 As String: Dim VAT5 As String: Dim VAT6 As String
    Dim VROAD As String: Dim VROAD1 As String: Dim VROAD2 As String: Dim VROAD3 As String
    Dim VLOC As String: Dim VLOC1 As String: Dim VLOC2 As String: Dim VLOC3 As String
    Dim VACT As String: Dim VACT1 As String: Dim VACT2 As String: Dim VACT3 As String: Dim VACT4 As String
    Dim VOCCU As String: Dim VOCCU1                    As String
    Dim VACC As String: Dim VACC1                      As String
    Dim VMODEL As String: Dim VCTYPE                   As String
    Dim VSUBLETAMOUNT                                  As Currency
    Dim vlineNo                                        As String
    Dim VDETCODE                                       As String
    Dim SUBLETTYPE                                     As String
    Dim vPREVACLNO                                     As String
    Dim vPREVRONO                                      As String

    Dim VMONTH                                         As Integer:
    Dim VKMS                                           As Integer
    Dim VMAIN                                          As String:
    Dim VSPEC                                          As String:
    Dim VREQ                                           As String:
    Dim VCHECK                                         As String
    Dim VAPP                                           As String:
    Dim VINV                                           As String:
    Dim VNCODE                                         As String
    Dim VPWAREQ                                        As String

    Dim vNATUREDESC                                    As String
    Dim vCAUSEDESC                                     As String
    Dim VFLATRATE                                      As String

    If IsNumeric(txtPWAT) = False Then
        ShowIsRequiredMsg "INVALID PWA TYPE"
        TabControl1.Tab = 0
        txtPWAT.SetFocus
        Exit Sub
    End If

    If cboCType.Text = "" Then
        ShowIsRequiredMsg "CLAIM TYPE CANNOT BE BLANK"
        TabControl1.Tab = 0
        cboCType.SetFocus
        Exit Sub
    End If

    If txtDLR.Text = "" Then
        ShowIsRequiredMsg "QIR REFERENCE NO. CANNOT BE BLANK"
        TabControl1.Tab = 0
        txtDLR.SetFocus
        Exit Sub
    End If

    If txtPWAT.Text = "" Then
        ShowIsRequiredMsg "PWA TYPE CANNOT BE BLANK"
        TabControl1.Tab = 0
        txtPWAT.SetFocus
        Exit Sub
    End If

    If cboNCODE.Text = "" Then
        ShowIsRequiredMsg "NATURE CODE CANNOT BE BLANK"
        cboNCODE.SetFocus
        Exit Sub
    End If

    If cboCCODE.Text = "" Then
        ShowIsRequiredMsg "CAUSE CODE CANNOT BE BLANK"
        cboCCODE.SetFocus
        Exit Sub
    End If
    
    If DEALER_CODE = "HMH" Then
        If txtRO.Text = "" Then
            ShowIsRequiredMsg "REPAIR ORDER NO. CANNOT BE BLANK"
            TabControl1.Tab = 0
            txtRO.SetFocus
            Exit Sub
        End If
    End If
    
    If txtKM.Text = "" Then
        ShowIsRequiredMsg "MILEAGE CANNOT BE BLANK"
        TabControl1.Tab = 0
        txtKM.SetFocus
        Exit Sub
    End If

    If txtPLATENO.Text = "" Then
        ShowIsRequiredMsg "PLATE NO CANNOT BE BLANK"
        TabControl1.Tab = 0
        txtPLATENO.SetFocus
        Exit Sub
    End If

    If txtVIN.Text = "" Then
        ShowIsRequiredMsg "VIN NO CANNOT BE BLANK"
        TabControl1.Tab = 0
        txtVIN.SetFocus
        Exit Sub
    End If

    If cboSA.Text = "" Then
        ShowIsRequiredMsg "SERVICE ADVISER CANNOT BE BLANK"
        TabControl1.Tab = 0
        cboSA.SetFocus
        Exit Sub
    End If

    '    If cboTECH.Text = "" Then
    '        ShowIsRequiredMsg "TECHNICIAN CANNOT BE BLANK"
    '        TabControl1.Tab = 0
    '        cboTECH.SetFocus
    '        Exit Sub
    '    End If

    If txtREQ.Text = "" Then
        ShowIsRequiredMsg "REQUESTED BY CANNOT BE BLANK"
        TabControl1.Tab = 2
        txtREQ.SetFocus
        Exit Sub
    End If

    If txtCHECK.Text = "" Then
        ShowIsRequiredMsg "CHECK BY CANNOT BE BLANK"
        TabControl1.Tab = 2
        txtCHECK.SetFocus
        Exit Sub
    End If

    If ADD_OR_EDIT = "ADD" Then
        Set RSTMP = gconDMIS.Execute("SELECT DLR_CQIR_REFERENCENO FROM CSMS_CQIR WHERE DLR_CQIR_REFERENCENO = '" & txtDLR.Text & "'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            MsgBox "QIR REFERENCE NO ALREADY EXIST", vbInformation, "CSMS"
            TabControl1.Tab = 0
            txtDLR.SetFocus
            Exit Sub
        End If
    Else
        Set RSTMP = gconDMIS.Execute("SELECT DLR_CQIR_REFERENCENO,ID FROM CSMS_CQIR WHERE DLR_CQIR_REFERENCENO = '" & txtDLR.Text & "'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            If Not LABID.Caption = RSTMP!ID Then
                MsgBox "QIR REFERENCE NO ALREADY EXIST", vbInformation, "CSMS"
                TabControl1.Tab = 0
                txtDLR.SetFocus
                Exit Sub
            End If
        End If
    End If

    If txtREC.Text = "" Then
        If MsgBox("Recommendation is blank, proceed", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then
            TabControl1.Tab = 1
            txtREC.SetFocus
            Exit Sub
        End If
    End If

    VFLATRATE = Format(NumericVal(txtFLATRATE), "#,###,##0.00")
    If optPWAREQ1.Value = True Then
        VPWAREQ = N2Str2Null("Y")
    Else
        VPWAREQ = N2Str2Null("N")
    End If

    vPREVACLNO = N2Str2Null(txtPREVACL.Text)
    vPREVRONO = N2Str2Null(txtPREVRO)
    VDLRNO = N2Str2Null(txtDLR.Text)
    'If Check3.Value = 0 Then
    '    VDELDATE = N2Str2Null("")
    'Else
        VDELDATE = N2Date2Null(dptDELDATTE)
    'End If
    VNAME = N2Str2Null(txtCUST.Text)
    vVin = N2Str2Null(txtVIN.Text)
    VENGINENO = N2Str2Null(txtENGINE.Text)
    VAXLE = N2Str2Null(txtAXLE.Text)

    'TRANSMISSION TYPE
    If optTRAN1.Value = True Then VTRANS = "MANUAL"
    If optTRAN2.Value = True Then VTRANS = "AUTO"

    'ATTACHMENT
    If optATT1.Value = True Then VATT = "PHOTO"
    If optATT2.Value = True Then VATT = "SAMPLE PART"

    vDCODE = N2Str2Null(txtDCODE.Text)

    'If Check1.Value = 0 Then
    '    VINSDATE = N2Str2Null("")
    'Else
        VINSDATE = N2Date2Null(dptINSD)
    'End If
    'If Check2.Value = 0 Then
    '    VREPDATE = N2Str2Null("")
    'Else
        VREPDATE = N2Date2Null(dptREPD)
    'End If
    'VREPDATE = N2Date2Null(dptREPD)
    vRO = N2Str2Null(txtRO.Text)
    VKM = N2Str2Null(txtKM.Text)
    VPLATE = N2Str2Null(txtPLATENO.Text)

    'SENT REGISTRATION
    If optSENT1.Value = True Then VSENT = "YES"
    If optSENT2.Value = True Then VSENT = "NO"

    VSA = N2Str2Null(cboSA.Text)
    VTECH = N2Str2Null(cboTECH.Text)

    VSUBJ = N2Str2Null(RTrim(LTrim(txtSUBJ.Text)))
    VHIST = N2Str2Null(RTrim(LTrim(txtHIST.Text)))
    VDESC = N2Str2Null(RTrim(LTrim(txtDESC.Text)))
    VANA = N2Str2Null(RTrim(LTrim(txtANA.Text)))
    VCORR = N2Str2Null(RTrim(LTrim(txtCORR.Text)))
    VREC = N2Str2Null(RTrim(LTrim(txtREC.Text)))
    VOTHER = N2Str2Null(RTrim(LTrim(txtCOMM.Text)))

    VCAS = N2Str2Null(txtCASUAL.Text)
    VNCODE = N2Str2Null(cboNCODE.Text)
    VCCODE = N2Str2Null(cboCCODE.Text)
    vNATUREDESC = N2Str2Null(txtNDESC)
    vCAUSEDESC = N2Str2Null(txtCDESC)

    VPCODE = N2Str2Null(txtPCODE.Text)
    VSCODE = N2Str2Null(txtSCODE.Text)

    VINV = N2Str2Null(txtINV.Text)
    VREQ = N2Str2Null(txtREQ.Text)
    VCHECK = N2Str2Null(txtCHECK.Text)
    VAPP = N2Str2Null(txtAPP.Text)

    'ENGINE TEMP
    VENGTEMP = optENG(0).Value
    VENGTEMP1 = optENG(1).Value
    VENGTEMP2 = optENG(2).Value

    'WEATHER
    VWEA = optWEA(0).Value
    VWEA1 = optWEA(1).Value
    VWEA2 = optWEA(2).Value

    'SHIFTING
    VSHIFT = optSHIF(0).Value
    VSHIFT1 = optSHIF(1).Value

    'SHIFT POSITION
    vPOSITION = optPOS(0).Value
    vPOSITION1 = optPOS(1).Value

    'MT
    VMT = optMT(0).Value
    VMT1 = optMT(1).Value
    VMT2 = optMT(2).Value
    VMT3 = optMT(3).Value
    VMT4 = optMT(4).Value
    VMT5 = optMT(5).Value
    VMT6 = optMT(6).Value

    'AT
    VAT = optAT(0).Value
    VAT1 = optAT(1).Value
    VAT2 = optAT(2).Value
    VAT3 = optAT(3).Value
    VAT4 = optAT(4).Value
    VAT5 = optAT(5).Value
    VAT6 = optAT(6).Value

    'ROAD
    VROAD = optROAD(0).Value
    VROAD1 = optROAD(1).Value
    VROAD2 = optROAD(2).Value
    VROAD3 = optROAD(3).Value

    'LOCATION
    VLOC = optLOC(0).Value
    VLOC1 = optLOC(1).Value
    VLOC2 = optLOC(2).Value
    VLOC3 = optLOC(3).Value

    'ACTION
    VACT = optACT(0).Value
    VACT1 = optACT(1).Value
    VACT2 = optACT(2).Value
    VACT3 = optACT(3).Value
    VACT4 = optACT(4).Value

    'OCCURENCE
    VOCCU = optOCC(0).Value
    VOCCU1 = optOCC(1).Value

    'ACCESSORIES
    VACC = optACC(0).Value
    VACC1 = optACC(1).Value

    'MAINTENANCE
    If optDEL(0).Value = True Then VMAIN = "DEALER"
    If optDEL(1).Value = True Then VMAIN = "3-STAR SHOP"
    If optDEL(2).Value = True Then VMAIN = "GAS STATION"
    If optDEL(3).Value = True Then VMAIN = "OTHER"

    Dim VGTOTAL                                        As Currency
    VCTYPE = N2Str2Null(cboCType.Text)
    VMODEL = N2Str2Null(txtModel.Text)
    VMONTH = NumericVal(txtEVERY.Text)
    VKMS = NumericVal(txtKMS.Text)
    VSPEC = N2Str2Null(txtSPE.Text)
    VSUBLETAMOUNT = CCur(TXTSUBREP.Text)
    vlineNo = N2Str2Null(lblLINENO.Caption)
    VDETCODE = N2Str2Null(lblDETCODE.Caption)
    VGTOTAL = NumericVal(TXTGRAND.Text)
    SUBLETTYPE = N2Str2Null(txtSubletType)

    Dim xROID                                          As String
    If ADD_OR_EDIT = "ADD" Then
        SQL_STATEMENT = "INSERT INTO CSMS_CQIR " & _
                      " (FLATRATE, NATUREDESC, CAUSEDESC, PREVACLNO, PREVRONO, SUBLETTYPE, GRANDTOTAL, LINE_NO, DETCODE, TOTALSUBLETREPAIR, TRANNO, MODEL, DELDATE, CLAIMTYPE, Ro_no, PWATYPE, DLR_CQIR_ReferenceNo, PWA_REQUEST, Customer, VINNo, EngineNo, TM_AxleNO, TransmissionType, Attachments, Subject, History, Description " & _
                      " ,Analysis, CorrectiveAction, Recommendation, CausalPartNo, NatureCode, CauseCode, PaintCode, subletCode, " & _
                      " InvoiceNo_ORNo, OtherComments, DealerCode, InspectionDate, RepairDate, Mileage, PlateNo, SentRegistrationCard, ServiceAdvisor, Technician, " & _
                      " Con_EngineTemp, Con_EngineTemp1, Con_EngineTemp2, Con_weather, Con_weather1, Con_weather2, Con_Shifting, Con_Shifting1, Con_ShifPosition, Con_ShifPosition1, " & _
                      " Con_MT, Con_MT1, Con_MT2, Con_MT3, Con_MT4, Con_MT5, Con_MT6, Con_AT, Con_AT1, Con_AT2, Con_AT3, Con_AT4, Con_AT5, CON_AT6, " & _
                      " Con_Road, Con_Road1, Con_Road2, Con_Road3, Con_Location, Con_Location1, Con_Location2, Con_Location3, " & _
                      " Con_Action, Con_Action1, Con_Action2, Con_Action3, Con_Action4, Con_Occurence, Con_Occurence1, Con_Accessories, Con_Accessories1, " & _
                      " VechicleMaintenance, EveryMonth, EveryKMS ,OtherCondition, RequestedBy, CheckedBy, TRANDATE) VALUES " & _
                      " (" & NumericVal(txtFLATRATE) & ", " & vNATUREDESC & "," & vCAUSEDESC & "," & vPREVACLNO & "," & vPREVRONO & "," & SUBLETTYPE & "," & VGTOTAL & "," & vlineNo & "," & VDETCODE & "," & VSUBLETAMOUNT & "," & lblTRANNO.Caption & "," & VMODEL & "," & VDELDATE & "," & VCTYPE & "," & vRO & "," & txtPWAT.Text & "," & VDLRNO & "," & VPWAREQ & "," & VNAME & "," & vVin & "," & VENGINENO & "," & VAXLE & ",'" & VTRANS & "','" & VATT & "'," & VSUBJ & "," & VHIST & "," & VDESC & _
                        "," & VANA & "," & VCORR & "," & VREC & "," & VCAS & "," & VNCODE & "," & VCCODE & "," & VPCODE & "," & VSCODE & _
                        "," & VINV & "," & VOTHER & "," & vDCODE & "," & VINSDATE & "," & VREPDATE & "," & VKM & "," & VPLATE & ",'" & VSENT & "'," & VSA & "," & VTECH & _
                        "," & VENGTEMP & "," & VENGTEMP1 & "," & VENGTEMP2 & "," & VWEA & "," & VWEA1 & "," & VWEA2 & "," & VSHIFT & "," & VSHIFT1 & "," & vPOSITION & "," & vPOSITION1 & _
                        "," & VMT & "," & VMT1 & "," & VMT2 & "," & VMT3 & "," & VMT5 & "," & VMT6 & "," & VAT & "," & VAT1 & "," & VAT2 & "," & VAT3 & "," & VAT3 & "," & VAT4 & "," & VAT5 & "," & VAT6 & _
                        "," & VROAD & "," & VROAD1 & "," & VROAD2 & "," & VROAD3 & "," & VLOC & "," & VLOC1 & "," & VLOC2 & "," & VLOC3 & _
                        "," & VACT & "," & VACT1 & "," & VACT2 & "," & VACT3 & "," & VACT4 & "," & VOCCU & "," & VOCCU1 & "," & VACC & "," & VACC1 & _
                        ",'" & VMAIN & "'," & VMONTH & "," & VKMS & "," & VSPEC & "," & VREQ & "," & VCHECK & "," & N2Str2Null(dtpTranDate.Value) & ")"
        gconDMIS.Execute SQL_STATEMENT

        Dim RSID                                       As New ADODB.Recordset
        Set RSID = gconDMIS.Execute("SELECT ID FROM CSMS_CQIR WHERE DLR_CQIR_REFERENCENO = " & VDLRNO & "")
        If Not (RSID.BOF And RSID.EOF) Then
            LABID.Caption = RSID!ID
        End If

        'NEW LOG AUDIT----------------------------------------------------------------------------
        If Not txtRO.Text = "" Then
            xROID = FindTransactionID(N2Str2Null(txtRO), "REP_OR", "CSMS_REPOR")
        Else
            xROID = ""
        End If

        Call NEW_LogAudit("A", "QUALITY INFORMATION", SQL_STATEMENT, LABID.Caption, "", "QIR NO: " & txtDLR.Text, "", xROID)
        Call NEW_LogAudit("AT", "BILLING SYSTEM", SQL_STATEMENT, xROID, "", "QIR NO: " & txtDLR, "", "")
        'NEW LOG AUDIT----------------------------------------------------------------------------

        ShowSuccessFullyAdded
    Else
        'xBEFORE_SAVE = GetBeforeValue(LABID, "CSMS_CQIR")

        SQL_STATEMENT = "UPDATE CSMS_CQIR SET FLATRATE = " & NumericVal(txtFLATRATE) & ", NATUREDESC = " & vNATUREDESC & ",CAUSEDESC = " & vCAUSEDESC & ", PREVACLNO = " & vPREVACLNO & ", PREVRONO = " & vPREVRONO & ", SUBLETTYPE = " & SUBLETTYPE & ", TRANDATE = " & N2Str2Null(dtpTranDate.Value) & ", GRANDTOTAL = " & VGTOTAL & ", TRANNO = " & lblTRANNO.Caption & ",MODEL = " & VMODEL & ",DELDATE = " & VDELDATE & ",CLAIMTYPE = " & VCTYPE & ", TOTALSUBLETREPAIR = " & VSUBLETAMOUNT & ",LINE_NO = " & vlineNo & ",DETCODE = " & VDETCODE & _
                        ", Ro_no = " & vRO & ", PWATYPE = " & txtPWAT.Text & ", DLR_CQIR_ReferenceNo = " & VDLRNO & ",PWA_REQUEST = " & VPWAREQ & ", Customer = " & VNAME & _
                        ", VINNo = " & vVin & " ,EngineNo = " & VENGINENO & ", TM_AxleNO = " & VAXLE & _
                        ", TransmissionType = '" & VTRANS & "', Attachments = '" & VATT & "', Subject = " & VSUBJ & ", History = " & VHIST & ", Description = " & VDESC & _
                      " ,Analysis = " & VANA & ", CorrectiveAction = " & VCORR & ", Recommendation = " & VREC & ", CausalPartNo = " & VCAS & ", NatureCode = " & VNCODE & ", CauseCode = " & VCCODE & _
                      " ,PaintCode = " & VPCODE & ", subletCode = " & VSCODE & _
                      " ,InvoiceNo_ORNo = " & VINV & ", OtherComments = " & VOTHER & ", DealerCode = " & vDCODE & ", InspectionDate = " & VINSDATE & ", RepairDate = " & VREPDATE & _
                        ", Mileage = " & VKM & ", PlateNo = " & VPLATE & ", SentRegistrationCard = '" & VSENT & "', ServiceAdvisor = " & VSA & ", Technician = " & VTECH & _
                        ", Con_EngineTemp = " & VENGTEMP & ", Con_EngineTemp1 = " & VENGTEMP1 & ", Con_EngineTemp2 = " & VENGTEMP2 & ", Con_weather = " & VWEA & ", Con_weather1 = " & VWEA1 & ", Con_weather2 = " & VWEA2 & _
                        ", Con_Shifting = " & VSHIFT & ", Con_Shifting1 = " & VSHIFT1 & ", Con_ShifPosition = " & vPOSITION & ",Con_ShifPosition1 = " & vPOSITION1 & _
                        ", Con_MT = " & VMT & ", Con_MT1 = " & VMT1 & ", Con_MT2 = " & VMT2 & ", Con_MT3 = " & VMT3 & ", Con_MT4 = " & VMT4 & ", Con_MT5 = " & VMT5 & ", Con_MT6 = " & VMT6 & _
                        ", Con_AT = " & VAT & ", Con_AT1 = " & VAT1 & ", Con_AT2 = " & VAT2 & ", Con_AT3 = " & VAT3 & ", Con_AT4 = " & VAT4 & ", Con_AT5 = " & VAT5 & ", Con_AT6 = " & VAT6 & _
                        ", Con_Road = " & VROAD & ", Con_Road1 = " & VROAD1 & ", Con_Road2 = " & VROAD2 & ", Con_Road3 = " & VROAD3 & ", Con_Location = " & VLOC & ", Con_Location1 = " & VLOC1 & ", Con_Location2 = " & VLOC2 & ", Con_Location3 = " & VLOC3 & _
                        ", Con_Action = " & VACT & ", Con_Action1 = " & VACT1 & ", Con_Action2 = " & VACT2 & ", Con_Action3 = " & VACT3 & ", Con_Action4 = " & VACT4 & _
                        ", Con_Occurence = " & VOCCU & ", Con_Occurence1 = " & VOCCU1 & ", Con_Accessories = " & VACC & ", Con_Accessories1 = " & VACC1 & _
                        ", VechicleMaintenance = '" & VMAIN & "', EveryMonth = " & VMONTH & ", EveryKMS = " & VKMS & ",OTHERCONDITION = " & VSPEC & _
                        ", RequestedBy = " & VREQ & ", CheckedBy = " & VCHECK & " WHERE ID = " & LABID.Caption & ""

        gconDMIS.Execute (SQL_STATEMENT)

        'NEW LOG AUDIT----------------------------------------------------------------------------
        If Not txtRO.Text = "" Then
            xROID = FindTransactionID(N2Str2Null(txtRO), "REP_OR", "CSMS_REPOR")
        Else
            xROID = ""
        End If

        Call NEW_LogAudit("E", "QUALITY INFORMATION", SQL_STATEMENT, LABID.Caption, "", "QIR NO: " & txtDLR.Text, "", xROID)
        Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, xROID, "", "QIR NO: " & txtDLR, "", "")
        'NEW LOG AUDIT----------------------------------------------------------------------------
        Dim OLD_QIR_NO                                 As String

        OLD_QIR_NO = Null2String(RS!DLR_CQIR_REFERENCENO)

        'EDIT CSMS_CQIRPARTS
        If Not lsvPARTS.ListItems.Count = 0 Then
            AUDIT_SQL = "UPDATE CSMS_CQIRPARTS SET DLR_CQIR_REFERENCENO = " & VDLRNO & " WHERE DLR_CQIR_REFERENCENO = '" & OLD_QIR_NO & "'"
            gconDMIS.Execute (AUDIT_SQL)

            'NEW LOG AUDIT------------------------------------------------------------------------------
            Call NEW_LogAudit("EE", "QUALITY INFORMATION", AUDIT_SQL, LABID.Caption, "", "DLR no. " & txtDLR.Text & " ,RO no: " & txtRO, "", xROID)
            'NEW LOG AUDIT------------------------------------------------------------------------------
        End If

        Call ShowSuccessFullyUpdated
    End If

    txtSEARCH.Text = "A": txtSEARCH.Text = ""
    rsRefresh
    RS.MoveFirst
    RS.Find "ID = " & LABID.Caption & ""
    picMENU.Visible = True
    cmdCancel_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim SQL                                            As String

    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If picMENU.Visible = True Then
                Unload frmALL_AuditInquiry

                frmALL_AuditInquiry.Show
                frmALL_AuditInquiry.ZOrder 0
                frmALL_AuditInquiry.Caption = "Audit Inquiry (Quality Information)"
                Call frmALL_AuditInquiry.DisplayHistory(LABID, "QUALITY INFORMATION")
            End If

        Case vbKeyF9:
            If Null2String(RS!Status) = "P" Then
                If Function_Access(LOGID, "Acess_POST", "QUALITY INFORMATION REPORT") = False Then Exit Sub
                If MsgBox("Approve QIR", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

                Call DisabledPics(False)
                picPWA.Visible = True
                picPWA.ZOrder 0
                txtPWAno1.Text = txtPWA
                txtAPP1.Text = txtAPP.Text
                txtPWAno1.SetFocus
            End If

        Case vbKeyF12:
            If Null2String(RS!Status) = "A" Then
                If Function_Access(LOGID, "Acess_UNPOST", "QUALITY INFORMATION REPORT") = False Then Exit Sub
                If MsgBox("Disapproved CQIR", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

                AUDIT_SQL = "UPDATE CSMS_CQIR SET APPROVEDBY = NULL, PWANO = NULL, STATUS = 'P',DATEAPPROVED = NULL WHERE ID = " & LABID.Caption & ""
                gconDMIS.Execute (AUDIT_SQL)

                'NEW LOG AUDIT---------------------------------------------------------------------
                Call NEW_LogAudit("DS", "QUALITY INFORMATION", AUDIT_SQL, LABID.Caption, "", "DLR NO: " & txtDLR.Text, "", "")
                'NEW LOG AUDIT---------------------------------------------------------------------
                MessagePop InfoFriend, "QIR Information Updated", "QIR Sucessfully Disaproved!", 1000
                rsRefresh
                RS.Find "ID = " & LABID.Caption & ""
                StoreMemvars
            End If

        Case vbKeyF3:                                 'ADD PARTS
            If picMENU.Visible = True Then
                If Null2String(RS!Status) = "" Then
                    If Function_Access(LOGID, "Acess_ADD", "QUALITY INFORMATION REPORT") = False Then Exit Sub

                    picMENU.Visible = False
                    Call DisabledSaveMenu(False)

                    txtPartNO.Text = "": txtPartDesc.Text = "": txtQty.Text = "": txtPartCost.Text = "": txtOPCODE.Text = "": txtALTS.Text = ""
                    picAddPart.Visible = True
                    picAddPart.ZOrder 0
                    'txtPartNo.SetFocus
                    cboPARTS.SetFocus
                    cmdDeletePart.Visible = False

                    ADD_OR_EDIT_ITEM = "ADD"
                End If
            End If

        Case vbKeyF4:
            If picMENU.Visible = True Then            'ADD SUBLET
                If Null2String(RS!Status) = "" Then
                    If Function_Access(LOGID, "Acess_ADD", "QUALITY INFORMATION REPORT") = False Then Exit Sub

                    picMENU.Visible = False
                    Call DisabledSaveMenu(False)
                    txtSubletDesc.Text = "":          'txtSCOST.Text = ""
                    picAddSub.Visible = True
                    picAddSub.ZOrder 0
                    cmdDeleteSublet.Visible = False

                    ADD_OR_EDIT_ITEM = "ADD"
                End If
            End If

        Case vbKeyF5:
            If picAdds.Visible = True Then
                Call cmdRefresh_Click
            End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    Set frm = New frmCSMS_MasterSearch
    
    Call InitMemVars
    Call InitCbo
    Call rsRefresh
    Call StoreMemvars
    cboSearchBy.ListIndex = 0

    Call txtSearch_Change
    Screen.MousePointer = 0
End Sub

Private Sub frm_SelectionMade(ByVal Code As String, SearchType As String)
    If SearchType = "PARTS" Then
        cboPARTS.Text = Code
    ElseIf SearchType = "CAUSAL PARTS" Then
        txtCASUAL.Text = Code
    End If
    Unload frm
End Sub

Private Sub lsvDLR_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsvDLR
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

Private Sub lsvDLR_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RS.MoveFirst
    RS.Find "ID = " & Item.ListSubItems(1) & ""
    Call StoreMemvars
End Sub

Private Sub lsvPARTS_DblClick()
    If Function_Access(LOGID, "Acess_EDIT", "QUALITY INFORMATION REPORT") = False Then Exit Sub

    If Null2String(RS!Status) = "A" Or Null2String(RS!Status) = "W" Or Null2String(RS!Status) = "P" Or Null2String(RS!Status) = "T" Or Null2String(RS!Status) = "C" Then Exit Sub

    Dim RSTMP                                          As New ADODB.Recordset
    Dim INDEX                                          As Double
    Dim vID                                            As Integer

    If lsvPARTS.ListItems.Count = 0 Then Exit Sub
    INDEX = lsvPARTS.SelectedItem.INDEX
    vID = lsvPARTS.ListItems(INDEX).ListSubItems(7)

    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_CQIRPARTS WHERE DLR_CQIR_REFERENCENO = '" & txtDLR & "' AND ID = " & vID & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        ADD_OR_EDIT_ITEM = "EDIT"
        InitMemVars1
        picMENU.Visible = False
        DisabledSaveMenu False

        If Trim(Null2String(RSTMP!ISSUBLET)) = "" Then    'PARTS
            picAddPart.Visible = True
            picAddPart.ZOrder 0

            cboPARTS.Text = Null2String(RSTMP!partno)
            'txtPartNo.Text = Null2String(RSTMP!partno)

            txtPartDesc.Text = Null2String(RSTMP!partname)
            txtQty.Text = NumericVal(RSTMP!QTY)
            txtPartCost.Text = NumericVal(RSTMP!cost)
            txtOPCODE.Text = Null2String(RSTMP!OPCODE)
            txtALTS.Text = NumericVal(RSTMP!lts)

            lblPITEMNO.Caption = RSTMP!ID
            cmdDeletePart.Visible = True
            'SendKeys "{F3}"
        ElseIf Trim(Null2String(RSTMP!ISSUBLET)) = "J" Then    'JOBS
            picAddJob.Visible = True
            picAddJob.ZOrder 0

            txtJobDesc.Text = Null2String(RSTMP!partname)
            txtJobCost.Text = NumericVal(RSTMP!cost)

            lblLITEMNO.Caption = RSTMP!ID
            'SendKeys "{F4}"
        Else                                          'SUBLETS
            picAddSub.Visible = True
            picAddSub.ZOrder 0

            txtSubCode.Text = Null2String(RSTMP!partno)
            txtSubletDesc.Text = Null2String(RSTMP!partname)
            txtSCOST.Text = NumericVal(RSTMP!cost)
            txtSOPCODE.Text = Null2String(RSTMP!OPCODE)
            txtSLTS.Text = Null2String(RSTMP!lts)

            lblSITEMNO.Caption = RSTMP!ID
            cmdDeleteSublet.Visible = True
            'SendKeys "{F5}"
        End If
    End If

    Set RSTMP = Nothing
End Sub

Private Sub optMAN_Click()
    If optMAN.Value = True Then
        txtPartNO.Visible = False
        cboPARTS.Visible = True
    Else
        txtPartNO.Visible = True
        cboPARTS.Visible = False
    End If
End Sub

Private Sub rptLIST_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim INDEX                                          As Long
    Dim vID                                            As Long
    Dim vRONO                                          As String

    If Row.Record Is Nothing Then: Exit Sub

    vRONO = Row.Record(0).Value                       'RO NO
    'cboTECH.Text = Row.Record(4).Value                      'TECHNICIAN NAME
    lblLINENO.Caption = Row.Record(5).Value
    lblDETCODE.Caption = Row.Record(6).Value

    txtRO.Text = vRONO
    Call FINDREPOR(vRONO)

    picSEARCHI.Visible = False
    picSEARCHI.ZOrder 1
    Call DisabledSaveMenu(True)
    picSEARCH.Enabled = False
End Sub

Private Sub rptPlate_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim INDEX                                          As Long
    Dim vID                                            As Long
    Dim vPLATENO                                       As String

    If Row.Record Is Nothing Then: Exit Sub

    txtPLATENO.Text = Null2String(Row.Record(0).Value)    'RO NO
    txtVIN.Text = Null2String(Row.Record(2).Value)
    txtENGINE.Text = Null2String(Row.Record(3).Value)
    txtModel.Text = Null2String(Row.Record(4).Value)

    picSeachPlateNo.Visible = False
    picSeachPlateNo.ZOrder 1

    Call DisabledSaveMenu(True)
    picSEARCH.Enabled = False
End Sub

Private Sub Timer1_Timer()
    If lblSTATUS.ForeColor = vbBlack Then
        lblSTATUS.ForeColor = vbRed
    Else
        lblSTATUS.ForeColor = vbBlack
    End If
End Sub

Private Sub txtALTS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789.", KeyAscii)
    End If
End Sub

Private Sub txtALTS_LostFocus()
    If IsNumeric(txtALTS.Text) = False Then txtALTS.Text = "0.00"
    'If txtALTS.Text = 0 Then txtALTS.Text = 1
End Sub

Private Sub txtANA_LostFocus()
    txtANA.Text = RTrim(LTrim(txtANA))
End Sub

Private Sub txtCASUAL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call frm.SelectSQl("select top 100 STOCKNO, STOCKDESC from PMIS_STOCKMAS WHERE STOCKNO LIKE", "CAUSAL PARTS")
        frm.Show 1
    End If
End Sub

Private Sub txtCDESC_Change()
    If Add_o_Edit = "" Then Exit Sub
    txtCDESC.ToolTipText = txtCDESC
End Sub

Private Sub txtCOMM_LostFocus()
    txtCOMM.Text = RTrim(LTrim(txtCOMM))
End Sub

Private Sub txtCORR_LostFocus()
    txtCORR.Text = LTrim(RTrim(txtCORR))
End Sub

Private Sub txtDESC_LostFocus()
    txtDESC.Text = RTrim(LTrim(txtDESC))
End Sub

Private Sub txtEVERY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789", KeyAscii)
    End If
End Sub

Private Sub txtEVERY_LostFocus()
    If IsNumeric(txtEVERY) = False Then
        txtEVERY.Text = 1
    Else
        If CInt(txtEVERY) > 12 Then
            txtEVERY = 12
        End If
    End If
End Sub

Private Sub txtFLATRATE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890.", KeyAscii)
    End If
End Sub

Private Sub txtHIST_Change()
    'txtHIST.Text = RTrim(LTrim(txtHIST))
End Sub

Private Sub txtJobCost_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789.", KeyAscii)
    End If
End Sub

Private Sub txtJobCost_LostFocus()
    If IsNumeric(txtJobCost.Text) = False Then txtJobCost = 0
    If txtJobCost.Text = "" Then txtJobCost = 0
    'If txtJobCost = 0 Then txtJobCost = 0
End Sub

Private Sub txtKM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789", KeyAscii)
    End If
End Sub

Private Sub txtKMS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789", KeyAscii)
    End If
End Sub

Private Sub txtLABORCOST_Change()
    On Error Resume Next

End Sub

Private Sub txtLABORCOST_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890.", KeyAscii)
    End If
End Sub

Private Sub txtLABORCOST_LostFocus()
    If IsNumeric(txtLABORCOST) = False Then
        txtLABORCOST.Text = "0.00"
    Else
        txtLABORCOST.Text = Format(txtLABORCOST.Text, "#,###,##0.00")
    End If
End Sub

Private Sub txtNDESC_GotFocus()
    txtNDESC.ToolTipText = txtNDESC
End Sub

Private Sub txtPartCost_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789.", KeyAscii)
    End If
End Sub

Private Sub txtPartCost_LostFocus()
    If txtPartCost.Text = "" Then txtPartCost = 0
    If txtPartCost.Text = 0 Then txtPartCost = 0
End Sub

Private Sub txtPartNO_LostFocus()
    On Error Resume Next
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT STOCKDESC FROM PMIS_STOCKMAS WHERE STOCKNO = '" & txtPartNO.Text & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        txtPartDesc.Text = Null2String(RSTMP!STOCKDESC)
    Else
        'txtPartDesc.Text = "Part no. not found"
    End If
    Set RSTMP = Nothing
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789", KeyAscii)
    End If
End Sub

Private Sub txtQty_LostFocus()
    If txtQty.Text = "" Then txtQty = 1
    If txtQty.Text = 0 Then txtQty = 1
End Sub

Private Sub txtREC_LostFocus()
    txtREC.Text = LTrim(RTrim(txtREC))
End Sub

Private Sub txtRO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtRO_LostFocus()
    If Not txtRO.Text = "" Then
        If Not Left(txtRO.Text, 1) = "R" Then
            txtRO.Text = "R-" & Format(txtRO, "00000000")
            If Not txtRO.Text = "" Then
                Call FillInfo
            End If
        End If
    End If
End Sub

Private Sub txtSCOST_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789.", KeyAscii)
    End If
End Sub

Private Sub txtSCOST_LostFocus()
    If IsNumeric(txtSCOST.Text) = False Then txtSCOST = 0
    If txtSCOST.Text = "" Then txtSCOST = 0
    'If txtSCOST = 0 Then txtSCOST = 0
End Sub

Private Sub txtSeachP_Change()
    rptPlate.FilterText = txtSeachP.Text
    rptPlate.Populate
End Sub

Private Sub txtSeachP_GotFocus()
    txtSeachP.BackColor = &HC0FFFF
End Sub

Private Sub txtSeachP_LostFocus()
    txtSeachP.BackColor = vbWhite
End Sub

Private Sub txtSearch_Change()
    Call FillSearchGrid(txtSEARCH.Text)
End Sub

Private Sub txtSEARCHI_Change()
    rptLIST.FilterText = txtSEARCHI.Text
    rptLIST.Populate
End Sub

Private Sub txtSEARCHI_GotFocus()
    txtSEARCHI.BackColor = &HC0FFFF
End Sub

Private Sub txtSEARCHI_LostFocus()
    txtSEARCHI.BackColor = vbWhite
End Sub

Private Sub txtSLTS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = LimitChar("1234567890.", KeyAscii)
    End If
End Sub

Private Sub txtSLTS_LostFocus()
    If IsNumeric(txtSLTS.Text) = False Then
        txtSLTS.Text = "0"
    End If
End Sub

Private Sub txtSUBJ_LostFocus()
    txtSUBJ.Text = RTrim(LTrim(txtSUBJ))
End Sub

Private Sub txtSUBQTY_LostFocus()
    If IsNumeric(txtSUBQTY) = False Then txtSUBQTY.Text = 0
End Sub

Private Sub TXTSUBREP_LostFocus()
    If IsNumeric(TXTSUBREP) = False Then TXTSUBREP.Text = "0.00"

    TXTGRAND.Text = Format(CCur(TXTSUBREP) + CCur(txtLABORCOST) + CCur(txtTCOST), "#,###,##0.00")
End Sub

Private Sub txtTCOST_Change()
    'txtLABORCOST.Text = Format(txtTCOST.Text * txtLTS, "#,###,##0.00")
End Sub

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                                           As String
    Dim cWidth                                         As Long
    Dim I                                              As Integer
    Dim scwidth                                        As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For I = LBound(ar) To UBound(ar)
            If I <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.ColumnHeaders(I + 1).Width = scwidth
            End If
        Next
    ElseIf TypeOf grd Is ReportControl Then
        For I = LBound(ar) To UBound(ar)
            If I < grd.Columns.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.Columns(I).Width = scwidth
            End If
        Next

    End If

    Erase ar
    grd.Visible = True
End Sub


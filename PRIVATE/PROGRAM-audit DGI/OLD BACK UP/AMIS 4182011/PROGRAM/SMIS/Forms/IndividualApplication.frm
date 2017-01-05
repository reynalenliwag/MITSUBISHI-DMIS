VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_Trans_ApplicationIndividual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Application Data Entry for Individual"
   ClientHeight    =   17775
   ClientLeft      =   195
   ClientTop       =   375
   ClientWidth     =   11505
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
   Icon            =   "IndividualApplication.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   17775
   ScaleWidth      =   11505
   Begin VB.PictureBox picBottoms 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   11505
      TabIndex        =   187
      Top             =   14985
      Width           =   11505
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   30
         ScaleHeight     =   1140
         ScaleWidth      =   11445
         TabIndex        =   188
         Top             =   30
         Width           =   11445
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   10350
            MouseIcon       =   "IndividualApplication.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   192
            ToolTipText     =   "Exit Window"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   795
            Left            =   9660
            MouseIcon       =   "IndividualApplication.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   191
            ToolTipText     =   "Print this Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdDocumentCheckList 
            Caption         =   "Documents"
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
            Left            =   8970
            MouseIcon       =   "IndividualApplication.frx":123A
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":138C
            Style           =   1  'Graphical
            TabIndex        =   216
            ToolTipText     =   "Add/Remove Require Document for Loan Application"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdUpdateStatus 
            Caption         =   "&Status"
            Height          =   795
            Left            =   8280
            MouseIcon       =   "IndividualApplication.frx":19FF
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":1B51
            Style           =   1  'Graphical
            TabIndex        =   217
            ToolTipText     =   "Update Loan Status"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdCancelCO 
            Caption         =   "Cancel Transaction"
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
            Left            =   7590
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "IndividualApplication.frx":2183
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":22D5
            Style           =   1  'Graphical
            TabIndex        =   211
            ToolTipText     =   "Cancel this Transaction"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdUnPost 
            Caption         =   "Unpost"
            Height          =   795
            Left            =   6900
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "IndividualApplication.frx":260F
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":2761
            Style           =   1  'Graphical
            TabIndex        =   210
            ToolTipText     =   "Unpost this Transaction"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "Post"
            Height          =   795
            Left            =   6210
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "IndividualApplication.frx":2AA6
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":2BF8
            Style           =   1  'Graphical
            TabIndex        =   212
            ToolTipText     =   "Post this Transaction"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   5520
            MouseIcon       =   "IndividualApplication.frx":2F1D
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":306F
            Style           =   1  'Graphical
            TabIndex        =   190
            ToolTipText     =   "Edit Selected Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   4830
            MouseIcon       =   "IndividualApplication.frx":33CB
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":351D
            Style           =   1  'Graphical
            TabIndex        =   189
            ToolTipText     =   "Add Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdLast 
            Caption         =   "Last"
            Height          =   795
            Left            =   4140
            MouseIcon       =   "IndividualApplication.frx":3830
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":3982
            Style           =   1  'Graphical
            TabIndex        =   197
            ToolTipText     =   "Move to Last Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdFirst 
            Caption         =   "First"
            Height          =   795
            Left            =   3450
            MouseIcon       =   "IndividualApplication.frx":3CD2
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":3E24
            Style           =   1  'Graphical
            TabIndex        =   196
            ToolTipText     =   "Move to First Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   2760
            MouseIcon       =   "IndividualApplication.frx":4182
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":42D4
            Style           =   1  'Graphical
            TabIndex        =   195
            ToolTipText     =   "Find a Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   2070
            MouseIcon       =   "IndividualApplication.frx":45CE
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":4720
            Style           =   1  'Graphical
            TabIndex        =   194
            ToolTipText     =   "Move to Next Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   1380
            MouseIcon       =   "IndividualApplication.frx":4A78
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":4BCA
            Style           =   1  'Graphical
            TabIndex        =   193
            ToolTipText     =   "Move to Previous Record"
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   7560
         ScaleHeight     =   885
         ScaleWidth      =   5940
         TabIndex        =   198
         Top             =   30
         Visible         =   0   'False
         Width           =   5940
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   2835
            MouseIcon       =   "IndividualApplication.frx":4F29
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":507B
            Style           =   1  'Graphical
            TabIndex        =   200
            ToolTipText     =   "Cancel"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   2145
            MouseIcon       =   "IndividualApplication.frx":53B9
            MousePointer    =   99  'Custom
            Picture         =   "IndividualApplication.frx":550B
            Style           =   1  'Graphical
            TabIndex        =   199
            ToolTipText     =   "Save this Record"
            Top             =   45
            Width           =   705
         End
      End
   End
   Begin VB.PictureBox picTops 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11505
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   9630
         Top             =   0
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   5250
         Top             =   75
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtAPL_No 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   75
         Width           =   2595
      End
      Begin VB.Label labTstatus 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   6690
         TabIndex        =   215
         Top             =   90
         Width           =   2790
      End
      Begin VB.Label labID 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   9000
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label labLStatus 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   3810
         TabIndex        =   2
         Top             =   120
         Width           =   2850
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apl Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Index           =   1
         Left            =   90
         TabIndex        =   1
         Top             =   75
         Width           =   960
      End
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3315
      Left            =   4080
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   3285
      ScaleWidth      =   3720
      TabIndex        =   174
      Top             =   1290
      Visible         =   0   'False
      Width           =   3750
      Begin VB.CommandButton cmdCancelStatus 
         Caption         =   "&Cancel"
         Height          =   675
         Index           =   0
         Left            =   2880
         MouseIcon       =   "IndividualApplication.frx":585B
         MousePointer    =   99  'Custom
         Picture         =   "IndividualApplication.frx":59AD
         Style           =   1  'Graphical
         TabIndex        =   213
         ToolTipText     =   "Cancel"
         Top             =   2490
         Width           =   675
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   177
         Top             =   570
         Width           =   3225
      End
      Begin VB.TextBox txtReasonNote 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1245
         Left            =   180
         MaxLength       =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   179
         Top             =   1230
         Width           =   3390
      End
      Begin VB.CommandButton cmdCancelStatus 
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
         Height          =   255
         Index           =   1
         Left            =   3420
         TabIndex        =   180
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Update"
         Height          =   675
         Left            =   2190
         MouseIcon       =   "IndividualApplication.frx":5CEB
         MousePointer    =   99  'Custom
         Picture         =   "IndividualApplication.frx":5E3D
         Style           =   1  'Graphical
         TabIndex        =   214
         ToolTipText     =   "Save Changes"
         Top             =   2490
         Width           =   705
      End
      Begin VB.Label Label68 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Note:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   178
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label68 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   176
         Top             =   360
         Width           =   600
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   330
         Left            =   0
         TabIndex        =   175
         Top             =   0
         Width           =   3735
         _Version        =   655364
         _ExtentX        =   6588
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   ":: Update Status ::"
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
   Begin VB.PictureBox picMiddles 
      Align           =   1  'Align Top
      Height          =   14520
      Left            =   0
      ScaleHeight     =   14460
      ScaleWidth      =   11445
      TabIndex        =   5
      Top             =   465
      Width           =   11505
      Begin VB.VScrollBar ScrollBar1 
         Height          =   2835
         LargeChange     =   500
         Left            =   11130
         Max             =   11160
         SmallChange     =   250
         TabIndex        =   6
         Top             =   0
         Value           =   10
         Width           =   265
      End
      Begin VB.PictureBox picIndividual 
         BorderStyle     =   0  'None
         Height          =   14400
         Left            =   0
         ScaleHeight     =   14400
         ScaleWidth      =   11160
         TabIndex        =   7
         Top             =   0
         Width           =   11160
         Begin VB.Frame fraSourceOfIncome 
            Caption         =   "Source of Income"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3240
            Left            =   60
            TabIndex        =   45
            Top             =   3555
            Width           =   11055
            Begin VB.CommandButton cmdCopyApplicant 
               Caption         =   "<="
               Height          =   315
               Left            =   5250
               TabIndex        =   64
               ToolTipText     =   "Copy Spouse Income Source To Applicant Income Source"
               Top             =   675
               Width           =   465
            End
            Begin VB.CommandButton cmdCopySpouse 
               Caption         =   "=>"
               Height          =   315
               Left            =   5250
               TabIndex        =   63
               ToolTipText     =   "Copy Applicant Income Source To Spouse Income Source"
               Top             =   300
               Width           =   465
            End
            Begin VB.Frame fraApplicantEmployment 
               Caption         =   "Applicant"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2970
               Left            =   90
               TabIndex        =   46
               Top             =   180
               Width           =   5085
               Begin VB.TextBox txtInd_Apl_EmpBusName 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   50
                  Text            =   " "
                  Top             =   420
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Apl_PrevAddress 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   62
                  Text            =   " "
                  Top             =   2565
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Apl_PreviousEmp 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   60
                  Text            =   " "
                  Top             =   2205
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Apl_TelNo 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   56
                  Text            =   " "
                  Top             =   1485
                  Width           =   3135
               End
               Begin VB.OptionButton optAplBusiness 
                  Caption         =   "Business"
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
                  Left            =   3240
                  TabIndex        =   48
                  Top             =   150
                  Width           =   1125
               End
               Begin VB.OptionButton optAplEmployment 
                  Caption         =   "Employment"
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
                  Left            =   1800
                  TabIndex        =   47
                  Top             =   150
                  Width           =   1365
               End
               Begin VB.TextBox txtInd_Apl_LengthOfStay 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   58
                  Text            =   " "
                  Top             =   1845
                  Width           =   765
               End
               Begin VB.TextBox txtInd_Apl_Position 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   54
                  Text            =   " "
                  Top             =   1125
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Apl_Address 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   52
                  Text            =   " "
                  Top             =   765
                  Width           =   3135
               End
               Begin VB.Label Label35 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Emp/Bus. Name : "
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
                  Left            =   285
                  TabIndex        =   49
                  Top             =   480
                  Width           =   1500
               End
               Begin VB.Label Label30 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Previous Address : "
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
                  Left            =   195
                  TabIndex        =   61
                  Top             =   2595
                  Width           =   1590
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Previous Emp. : "
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
                  Left            =   450
                  TabIndex        =   59
                  Top             =   2145
                  Width           =   1335
               End
               Begin VB.Label Label32 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tel. No(s) : "
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
                  Left            =   840
                  TabIndex        =   55
                  Top             =   1485
                  Width           =   945
               End
               Begin VB.Label Label41 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Length of Stay : "
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
                  Left            =   495
                  TabIndex        =   57
                  Top             =   1815
                  Width           =   1290
               End
               Begin VB.Label Label39 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Position : "
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
                  Left            =   975
                  TabIndex        =   53
                  Top             =   1155
                  Width           =   810
               End
               Begin VB.Label Label37 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address : "
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
                  Left            =   960
                  TabIndex        =   51
                  Top             =   825
                  Width           =   825
               End
            End
            Begin VB.Frame fraSpouseEmployment 
               Caption         =   "Spouse"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2970
               Left            =   5775
               TabIndex        =   65
               Top             =   180
               Width           =   5160
               Begin VB.TextBox txtSpouseEmpBusName 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   69
                  Text            =   " "
                  Top             =   405
                  Width           =   3135
               End
               Begin VB.TextBox txtSpouseAddress 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   71
                  Text            =   " "
                  Top             =   765
                  Width           =   3135
               End
               Begin VB.TextBox txtSpousePosition 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   73
                  Text            =   " "
                  Top             =   1125
                  Width           =   3135
               End
               Begin VB.TextBox txtSpouseLengthOfStay 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   77
                  Text            =   " "
                  Top             =   1845
                  Width           =   795
               End
               Begin VB.OptionButton optSpsEmployment 
                  Caption         =   "Employment"
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
                  Left            =   1800
                  TabIndex        =   66
                  Top             =   120
                  Width           =   1365
               End
               Begin VB.OptionButton optSpsBusiness 
                  Caption         =   "Business"
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
                  Left            =   3330
                  TabIndex        =   67
                  Top             =   120
                  Width           =   1185
               End
               Begin VB.TextBox txtSpouseTelNo 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   75
                  Text            =   " "
                  Top             =   1485
                  Width           =   3135
               End
               Begin VB.TextBox txtSpousePreviousEmp 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   79
                  Text            =   " "
                  Top             =   2205
                  Width           =   3135
               End
               Begin VB.TextBox txtSpousePrevAddress 
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
                  Height          =   330
                  Left            =   1800
                  TabIndex        =   81
                  Text            =   " "
                  Top             =   2565
                  Width           =   3135
               End
               Begin VB.Label Label42 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Emp/Bus. Name : "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   285
                  TabIndex        =   68
                  Top             =   495
                  Width           =   1500
               End
               Begin VB.Label Label40 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address : "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   960
                  TabIndex        =   70
                  Top             =   840
                  Width           =   825
               End
               Begin VB.Label Label38 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Position : "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   975
                  TabIndex        =   72
                  Top             =   1200
                  Width           =   810
               End
               Begin VB.Label Label36 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Length of Stay : "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   495
                  TabIndex        =   76
                  Top             =   1905
                  Width           =   1290
               End
               Begin VB.Label Label34 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tel. No(s) : "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   840
                  TabIndex        =   74
                  Top             =   1545
                  Width           =   945
               End
               Begin VB.Label Label33 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Previous Emp. : "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   450
                  TabIndex        =   78
                  Top             =   2250
                  Width           =   1335
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Previous Address : "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   195
                  TabIndex        =   80
                  Top             =   2610
                  Width           =   1590
               End
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               Caption         =   "Previous Address (if aboive address is less that two years) : "
               Height          =   225
               Index           =   0
               Left            =   90
               TabIndex        =   82
               Top             =   3750
               Width           =   5295
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "References"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2025
            Index           =   0
            Left            =   45
            TabIndex        =   125
            Top             =   10335
            Width           =   11055
            Begin VB.TextBox txtRef_Credit_Name2 
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
               Height          =   330
               Left            =   1350
               TabIndex        =   140
               Text            =   " "
               Top             =   1590
               Width           =   2580
            End
            Begin VB.TextBox txtRef_Credit_Name1 
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
               Height          =   330
               Left            =   1350
               TabIndex        =   137
               Text            =   " "
               Top             =   1220
               Width           =   2580
            End
            Begin VB.TextBox txtRef_Pers_Name1 
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
               Height          =   330
               Left            =   1350
               TabIndex        =   130
               Text            =   " "
               Top             =   480
               Width           =   2580
            End
            Begin VB.TextBox txtRef_Pers_Name2 
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
               Height          =   330
               Left            =   1350
               TabIndex        =   133
               Text            =   " "
               Top             =   850
               Width           =   2580
            End
            Begin VB.TextBox txtRef_Credit_Add2 
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
               Height          =   330
               Left            =   4035
               TabIndex        =   141
               Text            =   " "
               Top             =   1590
               Width           =   4410
            End
            Begin VB.TextBox txtRef_Credit_Add1 
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
               Height          =   330
               Left            =   4035
               TabIndex        =   138
               Text            =   " "
               Top             =   1220
               Width           =   4410
            End
            Begin VB.TextBox txtRef_Pers_Add1 
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
               Height          =   330
               Left            =   4035
               TabIndex        =   131
               Text            =   " "
               Top             =   480
               Width           =   4410
            End
            Begin VB.TextBox txtRef_Pers_Add2 
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
               Height          =   330
               Left            =   4035
               TabIndex        =   134
               Text            =   " "
               Top             =   850
               Width           =   4410
            End
            Begin VB.TextBox txtRef_Credit_TelNo2 
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
               Height          =   330
               Left            =   8505
               TabIndex        =   142
               Text            =   " "
               Top             =   1590
               Width           =   2130
            End
            Begin VB.TextBox txtRef_Credit_TelNo1 
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
               Height          =   330
               Left            =   8505
               TabIndex        =   139
               Text            =   " "
               Top             =   1220
               Width           =   2130
            End
            Begin VB.TextBox txtRef_Pers_TelNo1 
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
               Height          =   330
               Left            =   8505
               TabIndex        =   132
               Text            =   " "
               Top             =   480
               Width           =   2130
            End
            Begin VB.TextBox txtRef_Pers_TelNo2 
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
               Height          =   330
               Left            =   8505
               TabIndex        =   135
               Text            =   " "
               Top             =   850
               Width           =   2130
            End
            Begin VB.Label Label64 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Credit "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   0
               Left            =   510
               TabIndex        =   136
               Top             =   1170
               Width           =   555
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Personal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   0
               Left            =   510
               TabIndex        =   129
               Top             =   510
               Width           =   765
            End
            Begin VB.Label Label66 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
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
               Left            =   2115
               TabIndex        =   126
               Top             =   225
               Width           =   525
            End
            Begin VB.Label Label61 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
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
               Left            =   5535
               TabIndex        =   127
               Top             =   225
               Width           =   735
            End
            Begin VB.Label Label62 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tel. No."
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
               Left            =   9435
               TabIndex        =   128
               Top             =   225
               Width           =   645
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Bank Account(s)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Index           =   0
            Left            =   45
            TabIndex        =   143
            Top             =   12360
            Width           =   11055
            Begin VB.ComboBox cboInd_BA_Type4 
               Height          =   330
               Left            =   3555
               TabIndex        =   161
               Top             =   1635
               Width           =   2190
            End
            Begin VB.ComboBox cboInd_BA_Type3 
               Height          =   330
               Left            =   3555
               TabIndex        =   157
               Top             =   1250
               Width           =   2190
            End
            Begin VB.ComboBox cboInd_BA_Type2 
               Height          =   330
               Left            =   3555
               TabIndex        =   153
               Top             =   865
               Width           =   2190
            End
            Begin VB.ComboBox cboInd_BA_Type1 
               Height          =   330
               Left            =   3555
               TabIndex        =   149
               Top             =   480
               Width           =   2190
            End
            Begin VB.TextBox txtInd_BA_Bal2 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   8505
               TabIndex        =   155
               Tag             =   "0.00"
               Text            =   " "
               Top             =   865
               Width           =   2130
            End
            Begin VB.TextBox txtInd_BA_Bal1 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   8505
               TabIndex        =   151
               Tag             =   "0.00"
               Text            =   " "
               Top             =   480
               Width           =   2130
            End
            Begin VB.TextBox txtInd_BA_Bal3 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   8505
               TabIndex        =   159
               Tag             =   "0.00"
               Text            =   " "
               Top             =   1250
               Width           =   2130
            End
            Begin VB.TextBox txtInd_BA_Bal4 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   8505
               TabIndex        =   163
               Tag             =   "0.00"
               Text            =   " "
               Top             =   1635
               Width           =   2130
            End
            Begin VB.TextBox txtInd_BA_AcctNo2 
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
               Height          =   330
               Left            =   5775
               TabIndex        =   154
               Text            =   " "
               Top             =   865
               Width           =   2700
            End
            Begin VB.TextBox txtInd_BA_AcctNo1 
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
               Height          =   330
               Left            =   5775
               TabIndex        =   150
               Text            =   " "
               Top             =   480
               Width           =   2700
            End
            Begin VB.TextBox txtInd_BA_AcctNo3 
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
               Height          =   330
               Left            =   5775
               TabIndex        =   158
               Text            =   " "
               Top             =   1250
               Width           =   2700
            End
            Begin VB.TextBox txtInd_BA_AcctNo4 
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
               Height          =   330
               Left            =   5775
               TabIndex        =   162
               Text            =   " "
               Top             =   1635
               Width           =   2700
            End
            Begin VB.TextBox txtInd_BA_Bank2 
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
               Height          =   330
               Left            =   315
               TabIndex        =   152
               Text            =   " "
               Top             =   865
               Width           =   3210
            End
            Begin VB.TextBox txtInd_BA_Bank1 
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
               Height          =   330
               Left            =   315
               TabIndex        =   148
               Text            =   " "
               Top             =   480
               Width           =   3210
            End
            Begin VB.TextBox txtInd_BA_Bank3 
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
               Height          =   330
               Left            =   315
               TabIndex        =   156
               Text            =   " "
               Top             =   1250
               Width           =   3210
            End
            Begin VB.TextBox txtInd_BA_Bank4 
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
               Height          =   330
               Left            =   315
               TabIndex        =   160
               Text            =   " "
               Top             =   1635
               Width           =   3210
            End
            Begin VB.Label Label63 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Balance"
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
               Left            =   9195
               TabIndex        =   147
               Top             =   240
               Width           =   705
            End
            Begin VB.Label Label67 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Account Number"
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
               Left            =   6360
               TabIndex        =   146
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label68 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Type of Account"
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
               Left            =   3915
               TabIndex        =   145
               Top             =   240
               Width           =   1395
            End
            Begin VB.Label Label69 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bank/Branch"
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
               Left            =   1365
               TabIndex        =   144
               Top             =   240
               Width           =   1125
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Applicant Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3555
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   0
            Width           =   11055
            Begin VB.TextBox txtAppInfo_Sps_BirthDate 
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
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   6705
               TabIndex        =   24
               Top             =   900
               Width           =   1200
            End
            Begin VB.TextBox txtAppInfo_App_BirthDate 
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
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   6690
               TabIndex        =   18
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txtAppInfo_LengthOfStay 
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
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   1200
               TabIndex        =   32
               Text            =   " "
               Top             =   1740
               Width           =   1215
            End
            Begin VB.TextBox txtAppInfo_App_LastName 
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
               Height          =   375
               Left            =   1200
               TabIndex        =   15
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtAppInfo_App_FirstName 
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
               Height          =   375
               Left            =   3030
               TabIndex        =   16
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtAppInfo_App_MiddleName 
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
               Height          =   375
               Left            =   4860
               TabIndex        =   17
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtAppInfo_Sps_LastName 
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
               Height          =   375
               Left            =   1200
               TabIndex        =   21
               Top             =   900
               Width           =   1815
            End
            Begin VB.TextBox txtAppInfo_Sps_FirstName 
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
               Height          =   375
               Left            =   3030
               TabIndex        =   22
               Text            =   " "
               Top             =   900
               Width           =   1815
            End
            Begin VB.TextBox txtAppInfo_Sps_MiddleName 
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
               Height          =   375
               Left            =   4860
               TabIndex        =   23
               Text            =   " "
               Top             =   900
               Width           =   1815
            End
            Begin VB.TextBox txtAppInfo_Address 
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
               Height          =   375
               Left            =   1200
               TabIndex        =   28
               Text            =   " "
               Top             =   1320
               Width           =   7155
            End
            Begin VB.TextBox txtAppInfo_Cellphone 
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
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   9060
               TabIndex        =   30
               Text            =   " "
               Top             =   1320
               Width           =   1875
            End
            Begin VB.TextBox txtAppInfo_Telephone 
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
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   9060
               TabIndex        =   27
               Text            =   " "
               Top             =   900
               Width           =   1875
            End
            Begin VB.TextBox txtAppInfo_App_Age 
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
               Height          =   375
               Left            =   7920
               TabIndex        =   19
               Text            =   " "
               Top             =   480
               Width           =   435
            End
            Begin VB.TextBox txtAppInfo_Sps_Age 
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
               Height          =   375
               Left            =   7920
               TabIndex        =   25
               Text            =   " "
               Top             =   900
               Width           =   435
            End
            Begin VB.TextBox txtAppInfo_PreviousAddress 
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               ForeColor       =   &H00400000&
               Height          =   375
               Left            =   5130
               TabIndex        =   31
               Text            =   " "
               Top             =   1725
               Width           =   5775
            End
            Begin VB.TextBox txtAppInfo_NoDependent 
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
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1200
               TabIndex        =   36
               Text            =   " "
               Top             =   3060
               Width           =   1095
            End
            Begin VB.TextBox txtAppInfo_DependentAge 
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
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   2340
               TabIndex        =   37
               Text            =   " "
               Top             =   3060
               Width           =   2055
            End
            Begin VB.ComboBox cboAppInfo_OnwnerShip 
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
               Left            =   1200
               TabIndex        =   34
               Text            =   "cboOwnerShip"
               Top             =   2220
               Width           =   3255
            End
            Begin VB.ComboBox cboAppInfo_App_CivilStatus 
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
               Left            =   8460
               TabIndex        =   20
               Text            =   "cboAppCivilStatus"
               Top             =   480
               Width           =   2475
            End
            Begin VB.ComboBox cboAppInfo_AppCitizen 
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
               Left            =   1200
               TabIndex        =   35
               Top             =   2640
               Width           =   3255
            End
            Begin VB.Frame fraRented 
               Caption         =   "If Rented (?) ..."
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
               Height          =   1350
               Left            =   5130
               TabIndex        =   38
               Top             =   2130
               Width           =   5805
               Begin VB.TextBox txtAppInfo_MonthlyRental 
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
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   3660
                  TabIndex        =   40
                  Tag             =   "0.00"
                  Text            =   " "
                  Top             =   240
                  Width           =   2055
               End
               Begin VB.TextBox txtAppInfo_NameofLandlord 
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
                  Height          =   330
                  Left            =   3660
                  TabIndex        =   42
                  Text            =   " "
                  Top             =   600
                  Width           =   2055
               End
               Begin VB.TextBox txtAppInfo_LandlordTelNo 
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
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   3660
                  TabIndex        =   44
                  Text            =   " "
                  Top             =   960
                  Width           =   2055
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Monthly Rental : "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   2205
                  TabIndex        =   39
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Name of Landlord :  "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   1905
                  TabIndex        =   41
                  Top             =   600
                  Width           =   1665
               End
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tel. No. : "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   2775
                  TabIndex        =   43
                  Top             =   960
                  Width           =   765
               End
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "No. of Dependents : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   0
               Left            =   90
               TabIndex        =   209
               Top             =   2970
               Width           =   1110
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Citizenship : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   150
               TabIndex        =   208
               Top             =   2640
               Width           =   1050
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ownership : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   45
               TabIndex        =   207
               Top             =   2280
               Width           =   1155
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Length of Stay : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   0
               Left            =   60
               TabIndex        =   206
               Top             =   1740
               Width           =   1140
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   375
               TabIndex        =   205
               Top             =   1350
               Width           =   825
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   420
               TabIndex        =   204
               Top             =   930
               Width           =   780
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Applicant : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   315
               TabIndex        =   203
               Top             =   555
               Width           =   885
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
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
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   0
               Left            =   1200
               TabIndex        =   9
               Top             =   255
               Width           =   945
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
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
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   0
               Left            =   2970
               TabIndex        =   10
               Top             =   255
               Width           =   945
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
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
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   0
               Left            =   4770
               TabIndex        =   11
               Top             =   255
               Width           =   1125
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Birthdate : "
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
               Height          =   225
               Index           =   0
               Left            =   6705
               TabIndex        =   12
               Top             =   255
               Width           =   870
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tel(s):"
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
               Left            =   8460
               TabIndex        =   26
               Top             =   975
               Width           =   525
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cell(s):"
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
               Left            =   8430
               TabIndex        =   29
               Top             =   1350
               Width           =   600
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Age"
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
               Height          =   225
               Index           =   0
               Left            =   7965
               TabIndex        =   13
               Top             =   255
               Width           =   345
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Civil Status : "
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
               Height          =   225
               Index           =   0
               Left            =   8460
               TabIndex        =   14
               Top             =   255
               Width           =   1050
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Previous Address (if Above  Address is less that two years) : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   390
               Index           =   0
               Left            =   2460
               TabIndex        =   33
               Top             =   1740
               Width           =   2430
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Monthly Income/Expense"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3555
            Index           =   0
            Left            =   5850
            TabIndex        =   107
            Top             =   6780
            Width           =   5265
            Begin VB.TextBox txtMonthlyIncome_Amort 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   3105
               TabIndex        =   124
               Tag             =   "0.00"
               Text            =   " "
               Top             =   3000
               Width           =   2055
            End
            Begin VB.TextBox txtMonthlyIncome_Rental 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   3105
               TabIndex        =   122
               Tag             =   "0.00"
               Text            =   " "
               Top             =   2613
               Width           =   2055
            End
            Begin VB.TextBox txtMonthlyIncome_LivingExpense 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   3105
               TabIndex        =   120
               Tag             =   "0.00"
               Text            =   " "
               Top             =   2230
               Width           =   2055
            End
            Begin VB.TextBox txtMonthlyIncome_OtherIncome3 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   3105
               TabIndex        =   118
               Tag             =   "0.00"
               Text            =   " "
               Top             =   1847
               Width           =   2055
            End
            Begin VB.TextBox txtMonthlyIncome_OtherIncome2 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   3105
               TabIndex        =   116
               Tag             =   "0.00"
               Text            =   " "
               Top             =   1464
               Width           =   2055
            End
            Begin VB.TextBox txtMonthlyIncome_OtherIncome1 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   3105
               TabIndex        =   114
               Tag             =   "0.00"
               Text            =   " "
               Top             =   1081
               Width           =   2055
            End
            Begin VB.TextBox txtMonthlyIncome_OtherIncomeDesc3 
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
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1390
               TabIndex        =   117
               Text            =   " "
               Top             =   1847
               Width           =   1665
            End
            Begin VB.TextBox txtMonthlyIncome_OtherIncomeDesc2 
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
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1390
               TabIndex        =   115
               Text            =   " "
               Top             =   1464
               Width           =   1665
            End
            Begin VB.TextBox txtMonthlyIncome_OtherIncomeDesc1 
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
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1390
               TabIndex        =   113
               Text            =   " "
               Top             =   1081
               Width           =   1665
            End
            Begin VB.TextBox txtMonthlyIncome_Spouse 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   3105
               TabIndex        =   111
               Tag             =   "0.00"
               Text            =   " "
               Top             =   698
               Width           =   2055
            End
            Begin VB.TextBox txtMonthlyIncome_Applicant 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   3105
               TabIndex        =   109
               Tag             =   "0.00"
               Text            =   " "
               Top             =   315
               Width           =   2055
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Other Income 3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   90
               TabIndex        =   202
               Top             =   1950
               Width           =   1260
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Other Income 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   90
               TabIndex        =   201
               Top             =   1575
               Width           =   1260
            End
            Begin VB.Label Label59 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Amortizations:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   1905
               TabIndex        =   123
               Top             =   3075
               Width           =   1155
            End
            Begin VB.Label Label58 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rental:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   2475
               TabIndex        =   121
               Top             =   2685
               Width           =   585
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Less: Living Expenses:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   1155
               TabIndex        =   119
               Top             =   2355
               Width           =   1905
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Other Income 1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   90
               TabIndex        =   112
               Top             =   1125
               Width           =   1260
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse Income: "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   1665
               TabIndex        =   110
               Top             =   750
               Width           =   1395
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Applicant Income: "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   1560
               TabIndex        =   108
               Top             =   375
               Width           =   1500
            End
         End
         Begin VB.Frame frame90 
            Caption         =   "Loan Applied For/ AOR Calculations"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3555
            Left            =   75
            TabIndex        =   83
            Top             =   6780
            Width           =   5670
            Begin VB.OptionButton optLoan_Private 
               Caption         =   "Private"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   2025
               TabIndex        =   104
               Top             =   3210
               Width           =   1095
            End
            Begin VB.TextBox txtLoan_PlaceOfUse 
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
               Height          =   330
               Left            =   1890
               TabIndex        =   102
               Text            =   " "
               Top             =   2850
               Width           =   3345
            End
            Begin VB.ComboBox cboLoan_SAE 
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
               Left            =   1890
               TabIndex        =   87
               Top             =   615
               Width           =   3345
            End
            Begin VB.TextBox txtLoan_DownpaymentPerct 
               Alignment       =   2  'Center
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
               Height          =   330
               Left            =   1890
               TabIndex        =   91
               Tag             =   "0.00"
               Text            =   " "
               Top             =   1380
               Width           =   705
            End
            Begin VB.TextBox txtLoan_AORPercentage 
               Alignment       =   2  'Center
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
               Height          =   330
               Left            =   1890
               TabIndex        =   96
               Tag             =   "0.00"
               Text            =   " "
               Top             =   2100
               Width           =   1080
            End
            Begin VB.TextBox txtLoan_UnitCost 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   1890
               TabIndex        =   89
               Tag             =   "0.00"
               Text            =   " "
               Top             =   1005
               Width           =   3345
            End
            Begin VB.ComboBox cboLoan_Model 
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
               Left            =   1890
               TabIndex        =   85
               Top             =   240
               Width           =   3345
            End
            Begin VB.TextBox txtLoan_MonthlyAmortization 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
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
               Height          =   330
               Left            =   1890
               Locked          =   -1  'True
               TabIndex        =   100
               TabStop         =   0   'False
               Tag             =   "0.00"
               Text            =   " "
               Top             =   2475
               Width           =   3345
            End
            Begin VB.TextBox txtLoan_FinBalAmount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
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
               Height          =   330
               Left            =   1890
               Locked          =   -1  'True
               TabIndex        =   94
               TabStop         =   0   'False
               Tag             =   "0.00"
               Text            =   " "
               Top             =   1740
               Width           =   3330
            End
            Begin VB.OptionButton optLoan_Public 
               Caption         =   "Public"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   4500
               TabIndex        =   106
               Top             =   3210
               Width           =   1095
            End
            Begin VB.OptionButton optLoan_Business 
               Caption         =   "Business"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3150
               TabIndex        =   105
               Top             =   3210
               Width           =   1095
            End
            Begin VB.TextBox txtLoan_BankTerms 
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
               Height          =   330
               Left            =   4110
               TabIndex        =   97
               Tag             =   "0"
               Text            =   " "
               Top             =   2100
               Width           =   1110
            End
            Begin VB.TextBox txtLoan_DownPayment 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   2595
               TabIndex        =   92
               Tag             =   "0.00"
               Text            =   " "
               Top             =   1380
               Width           =   2640
            End
            Begin VB.Label Label55 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Purpose : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   990
               TabIndex        =   103
               Top             =   3210
               Width           =   840
            End
            Begin VB.Label Label56 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Place of Use : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   570
               TabIndex        =   101
               Top             =   2850
               Width           =   1185
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bank Terms: "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   3030
               TabIndex        =   99
               Top             =   2175
               Width           =   1095
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Net Amortization: "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   330
               TabIndex        =   98
               Top             =   2535
               Width           =   1425
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit/Model : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   750
               TabIndex        =   84
               Top             =   300
               Width           =   1005
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Financing Balance: "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   135
               TabIndex        =   93
               Top             =   1740
               Width           =   1620
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Cost:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   945
               TabIndex        =   88
               Top             =   1020
               Width           =   810
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Downpayment: (%) "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   150
               TabIndex        =   90
               Top             =   1380
               Width           =   1605
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AOR/Bank Terms: "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   240
               TabIndex        =   95
               Top             =   2115
               Width           =   1515
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Executive : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   330
               TabIndex        =   86
               Top             =   660
               Width           =   1425
            End
         End
      End
   End
   Begin VB.PictureBox pic4EditSO 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   3060
      ScaleHeight     =   4755
      ScaleWidth      =   5835
      TabIndex        =   164
      Top             =   870
      Visible         =   0   'False
      Width           =   5865
      Begin VB.TextBox txtFindAPL 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   166
         Top             =   690
         Width           =   4155
      End
      Begin VB.CommandButton cmdCancelSO 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   4740
         Picture         =   "IndividualApplication.frx":618D
         Style           =   1  'Graphical
         TabIndex        =   173
         ToolTipText     =   "Cancel"
         Top             =   4005
         Width           =   855
      End
      Begin VB.TextBox txtSearch_APL 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   170
         Top             =   3630
         Width           =   1125
      End
      Begin VB.TextBox txtSearch_AplName 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   171
         Top             =   3630
         Width           =   4215
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   2535
         Left            =   150
         TabIndex        =   169
         Top             =   1050
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   4471
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
         MouseIcon       =   "IndividualApplication.frx":64CB
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "APL No."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Last Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "First Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "MI"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdSaveSO 
         Caption         =   "&Select"
         Height          =   645
         Left            =   3900
         Picture         =   "IndividualApplication.frx":662D
         Style           =   1  'Graphical
         TabIndex        =   172
         ToolTipText     =   "Select"
         Top             =   4005
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Customer Name"
         Height          =   345
         Index           =   0
         Left            =   300
         TabIndex        =   167
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Search For Loan Application"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   135
         TabIndex        =   168
         Top             =   330
         Width           =   3765
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   165
         Top             =   0
         Width           =   5820
         _Version        =   655364
         _ExtentX        =   10266
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Edit Individual Application:::"
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
   Begin VB.PictureBox picDocumentList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   3060
      ScaleHeight     =   4755
      ScaleWidth      =   5835
      TabIndex        =   181
      Top             =   870
      Visible         =   0   'False
      Width           =   5865
      Begin VB.CommandButton Command3 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   5010
         MouseIcon       =   "IndividualApplication.frx":6969
         MousePointer    =   99  'Custom
         Picture         =   "IndividualApplication.frx":6ABB
         Style           =   1  'Graphical
         TabIndex        =   186
         ToolTipText     =   "Cancel"
         Top             =   3870
         Width           =   705
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3525
         Left            =   60
         TabIndex        =   183
         Top             =   300
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   6218
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
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
         MouseIcon       =   "IndividualApplication.frx":6DF9
         NumItems        =   0
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Save"
         Height          =   795
         Left            =   4320
         MouseIcon       =   "IndividualApplication.frx":6F5B
         MousePointer    =   99  'Custom
         Picture         =   "IndividualApplication.frx":70AD
         Style           =   1  'Graphical
         TabIndex        =   185
         ToolTipText     =   "Save this Record"
         Top             =   3870
         Width           =   705
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   795
         Left            =   3630
         MouseIcon       =   "IndividualApplication.frx":73FD
         MousePointer    =   99  'Custom
         Picture         =   "IndividualApplication.frx":754F
         Style           =   1  'Graphical
         TabIndex        =   184
         ToolTipText     =   "Add Record"
         Top             =   3870
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   182
         Top             =   0
         Width           =   5820
         _Version        =   655364
         _ExtentX        =   10266
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Update Document Check List For Individual Application:::"
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
End
Attribute VB_Name = "frmSMIS_Trans_ApplicationIndividual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================
'FUNCTION / FEATURE :AORVALUE: Computes AOR Regardless of Financing Company
'DATE STARTED       :5/26/200717:32
'LAST UPDATED       :5/26/200717:32
'DATABASE UPDATES   :NONE
'WHO UPDATED        :AXP  5/26/2007
'UDPATING CODE      :AXP-526200717:32
'==========================================================================================

Option Explicit
Dim rsLoan                                                            As ADODB.Recordset
Dim rsS_Model                                                         As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim xDateApplied                                                      As String
Dim Ctl                                                               As Control
Dim rsBAType                                                          As ADODB.Recordset
Dim ProspectID                                                        As Long
Dim CUSCDE                                                            As String
Dim ProfileType                                                       As String
Private APLNO                                                         As String
Dim WithEvents FormSearch                                             As frmSMIS_Mis_SearchMaster
Attribute FormSearch.VB_VarHelpID = -1
Dim WithEvents FormLog                                                As frmSMIS_Inquiry_ViewLog
Attribute FormLog.VB_VarHelpID = -1

Private AddingLoan                                                    As Boolean
Dim ComputebyPert                                                     As Boolean
Private LoanID                                                        As Long



Public Function ShowLoanApp(IDX As Long) As Boolean
    If IDX > 0 Then
        AddorEdit = "EDIT"
        LoanID = IDX
    Else
        AddorEdit = "ADD"
    End If
End Function

Public Function AddFromProspects(IDXPROSPECTID As Long) As Boolean
    AddingLoan = True
    If IDXPROSPECTID <> 0 Then
        AddingLoan = True
    Else
        Unload Me
    End If

    initMemvars

    Dim oCusRs                                                        As ADODB.Recordset

    Set oCusRs = gconDMIS.Execute("SELECT  * FROM CRIS_PROSPECTS WHERE PROSPECTID=" & IDXPROSPECTID)
    If oCusRs.EOF = True Or oCusRs.BOF = True Then
        MsgBox " Error Fetching Record"
        Exit Function
    End If

    '    If IsDate(oCusRs!LogApplication) = True Then
    '        MsgBox rsLoan.RecordCount
    '       StoreMemvars
    '       Exit Function
    '    End If

    CUSCDE = Null2String(oCusRs!CUSCDE)
    ProspectID = Null2String(oCusRs!ProspectID)
    ProfileType = Null2String(oCusRs!ProspectType)
    cboLoan_Model.Text = Null2String(oCusRs!Variant)
    cboLoan_SAE.Text = Null2String(oCusRs!SAE)
    Dim temprs                                                        As ADODB.Recordset

    If CUSCDE <> "" Then
        Set temprs = gconDMIS.Execute("SELECT * FROM ALL_CUSTOMER WHERE CUSCDE=" & N2Str2Null(CUSCDE))
        If Not (temprs.EOF Or temprs.BOF) Then
            Dim oSpouse                                               As Variant
            txtAppInfo_App_LastName = Null2String(temprs!lastname)
            txtAppInfo_App_FirstName = Null2String(temprs!FIRSTNAME)
            txtAppInfo_App_MiddleName = Null2String(temprs!MiddleInitial)
            txtAppInfo_App_BirthDate.Text = Null2String(temprs!BirthDate)
            oSpouse = Split(Null2String(temprs!Spouse), " ")
            If UBound(oSpouse) = 1 Then
                txtAppInfo_Sps_FirstName = oSpouse(0)
                txtAppInfo_Sps_LastName = oSpouse(1)
                cboAppInfo_App_CivilStatus.Text = "Married"
            ElseIf UBound(oSpouse) = 2 Then
                txtAppInfo_Sps_FirstName = oSpouse(0)
                txtAppInfo_Sps_LastName = oSpouse(1)
                txtAppInfo_Sps_MiddleName = oSpouse(2)
                cboAppInfo_App_CivilStatus.Text = "Married"
            ElseIf UBound(oSpouse) = 0 Then

                cboAppInfo_App_CivilStatus.Text = "Unspecified"
            Else
                txtAppInfo_Sps_LastName = Null2String(temprs!Spouse)
                cboAppInfo_App_CivilStatus.Text = "Married"
            End If
            txtAppInfo_Address = Null2String(temprs!CUSTOMERADD)
            txtAppInfo_Cellphone = Null2String(temprs!Mobile)
            txtAppInfo_Telephone = Null2String(temprs!HOMEPHONE)
            txtInd_Apl_EmpBusName = Null2String(temprs!CUSCOMP)
            txtInd_Apl_Address = Null2String(temprs!CompanyAdd)
            txtInd_Apl_Position = Null2String(temprs!TITLE)
            txtInd_Apl_TelNo = Null2String(temprs!TelephoneNo)
            cboAppInfo_AppCitizen = "Filippino"
        End If
    Else
        Dim arName                                                    As Variant
        arName = Split(Null2String(oCusRs!AcctName), " ")

        If UBound(arName) = 1 Then
            txtAppInfo_App_LastName = arName(1)
            txtAppInfo_App_FirstName = arName(0)

        ElseIf UBound(arName) = 2 Then
            txtAppInfo_App_LastName = arName(1)
            txtAppInfo_App_FirstName = arName(0)
            txtAppInfo_App_MiddleName = arName(2)
        ElseIf UBound(arName) = 0 Then
            txtAppInfo_App_FirstName = arName(0)
        End If

        txtAppInfo_Address = Null2String(oCusRs!Address)
        txtAppInfo_Cellphone = Null2String(oCusRs!Mobile)
        txtAppInfo_Telephone = Null2String(oCusRs!Telephone)
        Erase arName
    End If

CustomerCode:

    txtApl_No = GenerateCode("SMIS_LoanIndiv", "APL_No ", "00000000")
    AddorEdit = "ADD"
    picIndividual.Enabled = True
    picSaves.Visible = True
    picAdds.Visible = False
End Function

Private Function AORVALUE(Principal, AOR, Term) As Double
'On Error Resume Next
'AXP-526200717:32
    If AOR <= 0 Then: AORVALUE = 0: Exit Function
    If Principal <= 0 Then: AORVALUE = 0: Exit Function
    If Term <= 0 Then: AORVALUE = 0: Exit Function
    Dim Interest                                                      As Double
    Interest = NumericVal(AOR)
    Interest = AOR / 1200
    AORVALUE = FormatNumber((Principal * Interest / (1 - ((1 / (1 + Interest) ^ Term)))), 2)
End Function

Private Sub cboAppInfo_App_CivilStatus_Click()
    Dim isbool                                                        As Boolean
    isbool = IIf(cboAppInfo_App_CivilStatus.Text = "Single", False, True)
    fraSpouseEmployment.Enabled = isbool
    cmdCopySpouse.Enabled = isbool
    ShadeControl txtSpouseAddress, isbool
    ShadeControl txtAppInfo_Sps_Age, isbool
    ShadeControl txtAppInfo_Sps_BirthDate, isbool
    ShadeControl txtSpouseEmpBusName, isbool
    ShadeControl txtAppInfo_Sps_FirstName, isbool
    ShadeControl txtAppInfo_Sps_LastName, isbool
    ShadeControl txtSpouseLengthOfStay, isbool
    ShadeControl txtAppInfo_Sps_MiddleName, isbool
    ShadeControl txtMonthlyIncome_Spouse, isbool
    ShadeControl txtSpousePosition, isbool
    ShadeControl txtSpouseTelNo, isbool
    ShadeControl txtSpousePrevAddress, isbool
    ShadeControl txtSpousePreviousEmp, isbool

End Sub

Private Sub cboAppInfo_App_CivilStatus_GotFocus()
    VBComBoBoxDroppedDown cboAppInfo_App_CivilStatus
    'Set cCombo.AttachCombo =
End Sub

Private Sub cboAppInfo_OnwnerShip_Change()
    If cboAppInfo_OnwnerShip.Text = "Rented" Then
        fraRented.Enabled = True
        txtAppInfo_MonthlyRental.BackColor = vbWhite
        txtAppInfo_NameofLandlord.BackColor = vbWhite
        txtAppInfo_LandlordTelNo.BackColor = vbWhite
    Else
        fraRented.Enabled = False
        txtAppInfo_MonthlyRental.BackColor = vbButtonFace
        txtAppInfo_NameofLandlord.BackColor = vbButtonFace
        txtAppInfo_LandlordTelNo.BackColor = vbButtonFace
    End If
End Sub

Private Sub cboAppInfo_OnwnerShip_Click()
    cboAppInfo_OnwnerShip_Change
End Sub

Private Sub cboAppInfo_OnwnerShip_GotFocus()
'Set cCombo.AttachCombo =
    VBComBoBoxDroppedDown cboAppInfo_OnwnerShip
End Sub

Private Sub cboLoan_Model_GotFocus()

    VBComBoBoxDroppedDown cboLoan_Model
    'Set cCombo.AttachCombo =
End Sub

Private Sub cboLoan_SAE_GotFocus()
    VBComBoBoxDroppedDown cboLoan_SAE
    'Set cCombo.AttachCombo =
End Sub

Private Sub cmdADD_Click()
    If Function_Access(LOGID, "Acess_ADD", "INDIVIDUAL LOAN APPLICATION") = False Then Exit Sub
    Set FormSearch = New frmSMIS_Mis_SearchMaster
    Call FormSearch.SearchForProspects(" (isdate(logapplication) =0 and status<>'C') AND ProspectType='P' ")
    FormSearch.Show 1

End Sub

Private Sub cmdCancel_Click()
    picIndividual.Enabled = False
    picSaves.Visible = False
    picAdds.Visible = True
    AddorEdit = ""
    StoreMemVars

End Sub

'Upating Code       : AXP-0707200713:22
Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "INDIVIDUAL LOAN APPLICATION") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If MsgBox("Do You want to Cancel this Applications", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    gconDMIS.Execute ("UPDATE SMIS_LOANINDIV SET STATUS='C', LSTATUS='C' WHERE ID=" & LabID)
    gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET LOGAPPLICATION=NULL,APPNO=Null, LOGAPPLICATIONTYPE=NULL WHERE APPNO=" & N2Str2Null(rsLoan!Apl_no) & " AND PROSPECTID=" & ProspectID)
    rsLoan.Requery
    rsLoan.Find ("ID=" & LabID)
    StoreMemVars
    '  MessagePop RecSaveOk, "Transaction Cancelled", "Transaction Sucessfully Cancelled"





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancelSO_Click()
    ShowHidePictureBox2 pic4EditSO, False
End Sub

Private Sub cmdCancelStatus_Click(Index As Integer)
    ShowHidePictureBox2 picStatus, False
End Sub

Private Sub cmdCopyApplicant_Click()
    optAplBusiness = optSpsBusiness
    optAplEmployment = optSpsEmployment
    txtInd_Apl_EmpBusName = txtSpouseEmpBusName
    txtInd_Apl_Address = txtSpouseAddress
    txtInd_Apl_TelNo = txtSpouseTelNo
    txtInd_Apl_Position = txtSpousePosition
    txtInd_Apl_LengthOfStay = txtSpouseLengthOfStay
End Sub

Private Sub cmdCopySpouse_Click()
    optSpsBusiness = optAplBusiness
    optSpsEmployment = optAplEmployment
    txtSpouseEmpBusName = txtInd_Apl_EmpBusName
    txtSpouseAddress = txtInd_Apl_Address
    txtSpouseTelNo = txtInd_Apl_TelNo
    txtSpousePosition = txtInd_Apl_Position
    txtSpouseLengthOfStay = txtInd_Apl_LengthOfStay


End Sub

'Private Sub cmdDelete_Click()
'    If labLStatus.Caption = "Approved" Then
'        Call MsgBox("Cannot Delete Current Record.. Loan Has Already Been Approved", vbExclamation)
'        Exit Sub
'    End If
'
'    If MsgBox("Confirm ", vbYesNo + vbExclamation) = vbYes Then
'        AddorEdit = ""
'        AddingLoan = False
'        gconDMIS.Execute ("DELETE FROM SMIS_LoanIndiv WHERE APL_No=" & N2Str2Null(txtAPL_No))
'        gconDMIS.Execute ("DELETE FROM SMIS_LoanDocument Where AplType='I' And APLCODE=" & N2Str2Null(txtAPL_No))
'        gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET LOGAPPLICATION=NULL, LOGAPPLICATIONTYPE=NULL WHERE PROSPECTID=" & ProspectID)
'        InitMemVars
'        rsRefresh
'        StoreMemvars
'
'        If FormExist("MainForm") Then
'            MainForm.ShowData
'        End If
'    End If
'End Sub

'Upating Code       : AXP-0707200713:22
Public Sub cmdDocumentCheckList_Click()
    Dim sql                                                           As String
    Dim lst                                                           As ListItem
    Dim rs                                                            As ADODB.Recordset

    On Error GoTo ErrorCode:

    sql = " Select Code, DocumentName , 1 chklist from SMIS_DOCUMENT where code in (select DocumentCode from SMIS_LoanDocument Where AplType='I' and AplCode=" & N2Str2Null(txtApl_No) & "  )" & vbCrLf
    sql = sql & " Union " & vbCrLf
    sql = sql & " Select Code, DocumentName , 0  chklist  from SMIS_DOCUMENT where code not in (select DocumentCode from SMIS_LoanDocument Where AplType='I' and AplCode=" & N2Str2Null(txtApl_No) & "  )" & vbCrLf

    Set rs = gconDMIS.Execute(sql)
    ListView1.ListItems.Clear
    While Not rs.EOF
        Set lst = ListView1.ListItems.Add(, , Null2String(rs!CODE))
        Call lst.ListSubItems.Add(, , Null2String(rs!DocumentName))
        Call lst.ListSubItems.Add(, , Null2String(rs!CODE))
        lst.Checked = CBool(rs!Chklist)
        rs.MoveNext
    Wend
    Set rs = Nothing
    ShowHidePictureBox2 picDocumentList, True





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "INDIVIDUAL LOAN APPLICATION") = False Then Exit Sub
    picIndividual.Enabled = True
    AddorEdit = "EDIT"
    picSaves.Visible = True: picAdds.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    pic4EditSO.Visible = True
    pic4EditSO.ZOrder 0
    ShowHidePictureBox2 pic4EditSO, True
    txtFindAPL = "": txtSearch_APL = "": txtSearch_AplName = ""
    cmdSaveSO.Enabled = False
    txtFindAPL_Change
End Sub

Private Sub cmdFirst_Click()
    If Not rsLoan.BOF Then
        rsLoan.MoveFirst
    End If
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    If Not rsLoan.EOF Then
        rsLoan.MoveLast
    Else
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    On Error GoTo ADDER:
    rsLoan.MoveNext
    If rsLoan.EOF Then
        rsLoan.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ADDER:
    Err.Clear
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200713:22
Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "INDIVIDUAL LOAN APPLICATION") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If MsgBox("Do You want to Post this Applications", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    gconDMIS.Execute ("UPDATE SMIS_LOANINDIV SET STATUS='P' WHERE ID=" & LabID)
    rsRefresh
    rsLoan.Find ("ID=" & LabID)
    StoreMemVars
    '  MessagePop RecSaveOk, "Transaction Posted", "Transaction Sucessfully Posted"





    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo ADDER:
    rsLoan.MovePrevious
    If rsLoan.BOF Then
        rsLoan.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ADDER:
    Err.Clear
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200713:23
Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "INDIVIDUAL LOAN APPLICATION") = False Then Exit Sub
    On Error GoTo ErrorCode:

    CrystalReport1.Formulas(0) = "Company = '" & Company_name & "'"
    'rptGenREP.Formulas(1) = "CompanyAddress = '" & Company_Address & "'"
    'rptReleased.WindowAllowDrillDown = False
    PrintSQLReport CrystalReport1, SMIS_REPORT_PATH & "LoanIndiv.rpt", "{LOAN.ID}= " & Trim(LabID), DMIS_REPORT_Connection, 1





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:23
Private Sub cmdSave_Click()


    Dim Apl_no, AplCode, DateApplied                                  As String
    Dim Ind_Apl_LastName, Ind_Apl_FirstName, Ind_Apl_MidName, Ind_Apl_Birthday, Ind_Apl_Age As String
    Dim Ind_Sps_LastName, Ind_Sps_FirstName, Ind_Sps_MidName, Ind_Sps_Birthday, Ind_Sps_Age As String
    Dim Ind_Address, Ind_TelNo, Ind_CpNo, Ind_Length_of_Stay, Ind_Ownership, Ind_Civil_Status, Ind_Citizenship, Ind_No_Of_dependents, Ind_Previous_Address As String
    Dim Ind_Monthly_Rental, Ind_Name_of_Landlord, Ind_Landlord_TelNo  As String
    Dim Ind_Apl_EmpBusName, Ind_Apl_Address, Ind_Apl_Position, Ind_Apl_TelNo, Ind_Apl_LengthOfStay, Ind_Apl_PreviousEmp, Ind_Apl_PrevAddress As String
    Dim Ind_Sps_EmpBusName, Ind_Sps_Address, Ind_Sps_Position, Ind_Sps_TelNo, Ind_Sps_LengthOfStay, Ind_Sps_PreviousEmp, Ind_Sps_PrevAddress As String
    Dim Ind_MI_Applicant, Ind_MI_Spouse                               As String
    Dim Ind_MI_OtherIncome1Desc, Ind_MI_OtherIncome1Amount            As String
    Dim Ind_MI_OtherIncome2Desc, Ind_MI_OtherIncome2Amount            As String
    Dim Ind_MI_OtherIncome3Desc, Ind_MI_OtherIncome3Amount            As String
    Dim Ind_MI_LivingExpense, Ind_MI_Rental, Ind_MI_Amortizations     As String
    Dim Ind_LoanApl_UnitModel, Ind_LoanApl_LCP, Ind_LoanApl_DP, Ind_LoanApl_Term, Ind_LoanApl_AOR As String
    Dim Ind_LoanApl_Monthly_Amortization, Ind_LoanApl_Balance_FI_Perc, Ind_LoanApl_Balance_FI_Amount As String
    Dim Ind_LoanApl_Purpose, Ind_LoanApl_PlaceOfUse, Ind_LoanApl_SAE  As String
    Dim Ind_Ref_Pers_Name1, Ind_Ref_Pers_Add1, Ind_Ref_Pers_TelNo1    As String
    Dim Ind_Ref_Pers_Name2, Ind_Ref_Pers_Add2, Ind_Ref_Pers_TelNo2    As String
    Dim Ind_Ref_Credit_Name1, Ind_Ref_Credit_Add1, Ind_Ref_Credit_TelNo1 As String
    Dim Ind_Ref_Credit_Name2, Ind_Ref_Credit_Add2, Ind_Ref_Credit_TelNo2 As String
    Dim Ind_BA_Bank1, Ind_BA_Type1, Ind_BA_AcctNo1, Ind_BA_Bal1       As String
    Dim Ind_BA_Bank2, Ind_BA_Type2, Ind_BA_AcctNo2, Ind_BA_Bal2       As String
    Dim Ind_BA_Bank3, Ind_BA_Type3, Ind_BA_AcctNo3, Ind_BA_Bal3       As String
    Dim Ind_BA_Bank4, Ind_BA_Type4, Ind_BA_AcctNo4, Ind_BA_Bal4
    Dim sql                                                           As String

    On Error GoTo ErrorCode:

    Apl_no = N2Str2Null(txtApl_No)
    AplCode = N2Str2Null(CUSCDE)
    DateApplied = N2Str2Null(FormatDateTime(Now, vbShortDate))
    Ind_Apl_LastName = N2Str2Null(txtAppInfo_App_LastName)
    Ind_Apl_FirstName = N2Str2Null(txtAppInfo_App_FirstName)
    Ind_Apl_MidName = N2Str2Null(txtAppInfo_App_MiddleName)
    Ind_Apl_Birthday = N2Str2Null(txtAppInfo_App_BirthDate)
    Ind_Apl_Age = N2Str2Null(txtAppInfo_App_Age)
    Ind_Sps_LastName = N2Str2Null(txtAppInfo_Sps_LastName)
    Ind_Sps_FirstName = N2Str2Null(txtAppInfo_Sps_FirstName)
    Ind_Sps_MidName = N2Str2Null(txtAppInfo_Sps_MiddleName)
    Ind_Sps_Birthday = N2Str2Null(txtAppInfo_Sps_BirthDate)
    Ind_Sps_Age = N2Str2Null(txtAppInfo_Sps_Age)
    Ind_Address = N2Str2Null(txtAppInfo_Address)
    Ind_TelNo = N2Str2Null(txtAppInfo_Telephone)
    Ind_CpNo = N2Str2Null(txtAppInfo_Cellphone)
    Ind_Length_of_Stay = N2Str2Null(txtAppInfo_LengthOfStay)
    Ind_Ownership = N2Str2Null(cboAppInfo_OnwnerShip)
    Ind_Civil_Status = N2Str2Null(cboAppInfo_App_CivilStatus)
    Ind_Citizenship = N2Str2Null(cboAppInfo_AppCitizen)
    Ind_No_Of_dependents = N2Str2Null(txtAppInfo_NoDependent)
    Ind_Previous_Address = N2Str2Null(txtAppInfo_PreviousAddress)
    Ind_Monthly_Rental = N2Str2Null(txtAppInfo_MonthlyRental)
    Ind_Name_of_Landlord = N2Str2Null(txtAppInfo_NameofLandlord)
    Ind_Landlord_TelNo = N2Str2Null(txtAppInfo_LandlordTelNo)

    Ind_Apl_EmpBusName = N2Str2Null(txtInd_Apl_EmpBusName)
    Ind_Apl_Address = N2Str2Null(txtInd_Apl_Address)
    Ind_Apl_Position = N2Str2Null(txtInd_Apl_Position)
    Ind_Apl_TelNo = N2Str2Null(txtInd_Apl_TelNo)
    Ind_Apl_LengthOfStay = N2Str2Null(txtInd_Apl_LengthOfStay)
    Ind_Apl_PreviousEmp = N2Str2Null(txtInd_Apl_PreviousEmp)
    Ind_Apl_PrevAddress = N2Str2Null(txtInd_Apl_PrevAddress)

    Ind_Sps_EmpBusName = N2Str2Null(txtSpouseEmpBusName)
    Ind_Sps_Address = N2Str2Null(txtSpouseAddress)
    Ind_Sps_Position = N2Str2Null(txtSpousePosition)
    Ind_Sps_TelNo = N2Str2Null(txtSpouseTelNo)
    Ind_Sps_LengthOfStay = N2Str2Null(txtSpouseLengthOfStay)
    Ind_Sps_PreviousEmp = N2Str2Null(txtSpousePreviousEmp)
    Ind_Sps_PrevAddress = N2Str2Null(txtSpousePrevAddress)

    ''INCOME TABLE
    Ind_MI_Applicant = N2Str2Null(txtMonthlyIncome_Applicant)
    Ind_MI_Spouse = N2Str2Null(txtMonthlyIncome_Spouse)

    Ind_MI_OtherIncome1Desc = N2Str2Null(txtMonthlyIncome_OtherIncomeDesc1)
    Ind_MI_OtherIncome1Amount = N2Str2Null(txtMonthlyIncome_OtherIncome1)

    Ind_MI_OtherIncome2Desc = N2Str2Null(txtMonthlyIncome_OtherIncomeDesc2)
    Ind_MI_OtherIncome2Amount = N2Str2Null(txtMonthlyIncome_OtherIncome2)

    Ind_MI_OtherIncome3Desc = N2Str2Null(txtMonthlyIncome_OtherIncomeDesc3)
    Ind_MI_OtherIncome3Amount = N2Str2Null(txtMonthlyIncome_OtherIncome3)

    Ind_MI_LivingExpense = N2Str2Null(txtMonthlyIncome_LivingExpense)
    Ind_MI_Rental = N2Str2Null(txtMonthlyIncome_Rental)
    Ind_MI_Amortizations = N2Str2Null(txtMonthlyIncome_Amort)
    ''END INCOME TABLE

    ''LOAN/AMORT TAB
    Ind_LoanApl_UnitModel = N2Str2Null(cboLoan_Model)
    Ind_LoanApl_LCP = N2Str2Null(txtLoan_UnitCost)
    Ind_LoanApl_DP = N2Str2Null(txtLoan_Downpayment)
    Ind_LoanApl_Term = N2Str2Null(txtLoan_BankTerms)
    Ind_LoanApl_AOR = N2Str2Null(txtLoan_AORPercentage)
    Ind_LoanApl_Monthly_Amortization = N2Str2Null(txtLoan_MonthlyAmortization)
    Ind_LoanApl_Balance_FI_Perc = N2Str2Null(txtLoan_DownpaymentPerct)
    Ind_LoanApl_Balance_FI_Amount = N2Str2Null(txtLoan_FinBalAmount)

    If optLoan_Private.Value = True Then
        Ind_LoanApl_Purpose = N2Str2Null("PRIVATE")
    ElseIf optLoan_Business.Value = True Then
        Ind_LoanApl_Purpose = N2Str2Null("BUSINESS")
    ElseIf optLoan_Public.Value = True Then
        Ind_LoanApl_Purpose = N2Str2Null("PUBLIC")
    End If



    Ind_LoanApl_PlaceOfUse = N2Str2Null(txtLoan_PlaceOfUse)
    Ind_LoanApl_SAE = N2Str2Null(cboLoan_SAE)
    ''END LOAN/AMORT TAB

    ''REFERENCES
    Ind_Ref_Pers_Name1 = N2Str2Null(txtRef_Pers_Name1)
    Ind_Ref_Pers_Add1 = N2Str2Null(txtRef_Pers_Add1)
    Ind_Ref_Pers_TelNo1 = N2Str2Null(txtRef_Pers_TelNo1)

    Ind_Ref_Pers_Name2 = N2Str2Null(txtRef_Pers_Name2)
    Ind_Ref_Pers_Add2 = N2Str2Null(txtRef_Pers_Add2)
    Ind_Ref_Pers_TelNo2 = N2Str2Null(txtRef_Pers_TelNo2)

    Ind_Ref_Credit_Name1 = N2Str2Null(txtRef_Credit_Name1)
    Ind_Ref_Credit_Add1 = N2Str2Null(txtRef_Credit_Add1)
    Ind_Ref_Credit_TelNo1 = N2Str2Null(txtRef_Credit_TelNo1)

    Ind_Ref_Credit_Name2 = N2Str2Null(txtRef_Credit_Name2)
    Ind_Ref_Credit_Add2 = N2Str2Null(txtRef_Credit_Add2)
    Ind_Ref_Credit_TelNo2 = N2Str2Null(txtRef_Credit_TelNo2)
    ''END REFERENCES

    'BANK INFO
    Ind_BA_Bank1 = N2Str2Null(txtInd_BA_Bank1)
    Ind_BA_Type1 = N2Str2Null(cboInd_BA_Type1)
    Ind_BA_AcctNo1 = N2Str2Null(txtInd_BA_AcctNo1)
    Ind_BA_Bal1 = N2Str2Null(txtInd_BA_Bal1)

    Ind_BA_Bank2 = N2Str2Null(txtInd_BA_Bank2)
    Ind_BA_Type2 = N2Str2Null(cboInd_BA_Type2)
    Ind_BA_AcctNo2 = N2Str2Null(txtInd_BA_AcctNo2)
    Ind_BA_Bal2 = N2Str2Null(txtInd_BA_Bal2)

    Ind_BA_Bank3 = N2Str2Null(txtInd_BA_Bank3)
    Ind_BA_Type3 = N2Str2Null(cboInd_BA_Type3)
    Ind_BA_AcctNo3 = N2Str2Null(txtInd_BA_AcctNo3)
    Ind_BA_Bal3 = N2Str2Null(txtInd_BA_Bal3)

    Ind_BA_Bank4 = N2Str2Null(txtInd_BA_Bank4)
    Ind_BA_Type4 = N2Str2Null(cboInd_BA_Type4)
    Ind_BA_AcctNo4 = N2Str2Null(txtInd_BA_AcctNo4)
    Ind_BA_Bal4 = N2Str2Null(txtInd_BA_Bal4)


    If AddorEdit = "ADD" Then
        sql = "INSERT INTO SMIS_LoanIndiv( " & vbCrLf
        sql = sql & "ProspectID, APL_No, AplCode, DateApplied,  " & vbCrLf
        sql = sql & "Ind_Apl_LastName, Ind_Apl_FirstName, Ind_Apl_MidName, Ind_Apl_Birthday, Ind_Apl_Age,  " & vbCrLf
        sql = sql & "Ind_Sps_LastName, Ind_Sps_FirstName, Ind_Sps_MidName, Ind_Sps_Birthday, Ind_Sps_Age,  " & vbCrLf
        sql = sql & "Ind_Address,Ind_TelNo, Ind_CpNo, Ind_Length_of_Stay,Ind_Previous_Address, Ind_Ownership, Ind_Civil_Status, Ind_Citizenship, Ind_No_Of_dependents , " & vbCrLf
        sql = sql & "Ind_Monthly_Rental, Ind_Name_of_Landlord, Ind_Landlord_TelNo,  " & vbCrLf

        sql = sql & "Ind_Apl_EmpBusName, Ind_Apl_Address, Ind_Apl_Position, Ind_Apl_TelNo, Ind_Apl_LengthOfStay, Ind_Apl_PreviousEmp, Ind_Apl_PrevAddress,  " & vbCrLf
        sql = sql & "Ind_Sps_EmpBusName, Ind_Sps_Address, Ind_Sps_Position, Ind_Sps_TelNo, Ind_Sps_LengthOfStay, Ind_Sps_PreviousEmp, Ind_Sps_PrevAddress,  " & vbCrLf

        sql = sql & "Ind_MI_Applicant, Ind_MI_Spouse,  " & vbCrLf
        sql = sql & "Ind_MI_OtherIncome1Desc, Ind_MI_OtherIncome1Amount,  " & vbCrLf
        sql = sql & "Ind_MI_OtherIncome2Desc, Ind_MI_OtherIncome2Amount,  " & vbCrLf
        sql = sql & "Ind_MI_OtherIncome3Desc, Ind_MI_OtherIncome3Amount,  " & vbCrLf
        sql = sql & "Ind_MI_LivingExpense, Ind_MI_Rental, Ind_MI_Amortizations,  " & vbCrLf
        sql = sql & "Ind_LoanApl_UnitModel, Ind_LoanApl_LCP, Ind_LoanApl_DP, Ind_LoanApl_Term, Ind_LoanApl_AOR,  " & vbCrLf
        sql = sql & "Ind_LoanApl_Monthly_Amortization, Ind_LoanApl_Balance_FI_Perc, Ind_LoanApl_Balance_FI_Amount,  " & vbCrLf
        sql = sql & "Ind_LoanApl_Purpose, Ind_LoanApl_PlaceOfUse, Ind_LoanApl_SAE,  " & vbCrLf
        sql = sql & "Ind_Ref_Pers_Name1, Ind_Ref_Pers_Add1, Ind_Ref_Pers_TelNo1,  " & vbCrLf
        sql = sql & "Ind_Ref_Pers_Name2, Ind_Ref_Pers_Add2, Ind_Ref_Pers_TelNo2,  " & vbCrLf
        sql = sql & "Ind_Ref_Credit_Name1, Ind_Ref_Credit_Add1, Ind_Ref_Credit_TelNo1,  " & vbCrLf
        sql = sql & "Ind_Ref_Credit_Name2, Ind_Ref_Credit_Add2, Ind_Ref_Credit_TelNo2,  " & vbCrLf
        sql = sql & "Ind_BA_Bank1, Ind_BA_Type1, Ind_BA_AcctNo1, Ind_BA_Bal1,  " & vbCrLf
        sql = sql & "Ind_BA_Bank2, Ind_BA_Type2, Ind_BA_AcctNo2, Ind_BA_Bal2,  " & vbCrLf
        sql = sql & "Ind_BA_Bank3, Ind_BA_Type3, Ind_BA_AcctNo3, Ind_BA_Bal3,  " & vbCrLf
        sql = sql & "Ind_BA_Bank4, Ind_BA_Type4, Ind_BA_AcctNo4, Ind_BA_Bal4, LStatus) " & vbCrLf
        sql = sql & "VALUES( " & vbCrLf
        sql = sql & ProspectID & " , " & Apl_no & " , " & AplCode & " , " & DateApplied & " , " & vbCrLf
        sql = sql & Ind_Apl_LastName & " , " & Ind_Apl_FirstName & " , " & Ind_Apl_MidName & " , " & Ind_Apl_Birthday & " , " & Ind_Apl_Age & " , " & vbCrLf
        sql = sql & Ind_Sps_LastName & " , " & Ind_Sps_FirstName & " , " & Ind_Sps_MidName & " , " & Ind_Sps_Birthday & " , " & Ind_Sps_Age & " , " & vbCrLf
        sql = sql & Ind_Address & " , " & Ind_TelNo & " , " & Ind_CpNo & " , " & Ind_Length_of_Stay & " , " & Ind_Previous_Address & "," & Ind_Ownership & " , " & Ind_Civil_Status & " , " & Ind_Citizenship & " , " & Ind_No_Of_dependents & " , " & vbCrLf
        sql = sql & Ind_Monthly_Rental & " , " & Ind_Name_of_Landlord & " , " & Ind_Landlord_TelNo & " , " & vbCrLf
        sql = sql & Ind_Apl_EmpBusName & " , " & Ind_Apl_Address & " , " & Ind_Apl_Position & " , " & Ind_Apl_TelNo & " , " & Ind_Apl_LengthOfStay & " , " & Ind_Apl_PreviousEmp & " , " & Ind_Apl_PrevAddress & " , " & vbCrLf
        sql = sql & Ind_Sps_EmpBusName & " , " & Ind_Sps_Address & " , " & Ind_Sps_Position & " , " & Ind_Sps_TelNo & " , " & Ind_Sps_LengthOfStay & " , " & Ind_Sps_PreviousEmp & " , " & Ind_Sps_PrevAddress & " , " & vbCrLf
        sql = sql & Ind_MI_Applicant & " , " & Ind_MI_Spouse & " , " & vbCrLf
        sql = sql & Ind_MI_OtherIncome1Desc & " , " & Ind_MI_OtherIncome1Amount & " , " & vbCrLf
        sql = sql & Ind_MI_OtherIncome2Desc & " , " & Ind_MI_OtherIncome2Amount & " , " & vbCrLf
        sql = sql & Ind_MI_OtherIncome3Desc & " , " & Ind_MI_OtherIncome3Amount & " , " & vbCrLf
        sql = sql & Ind_MI_LivingExpense & " , " & Ind_MI_Rental & " , " & Ind_MI_Amortizations & " , " & vbCrLf
        sql = sql & Ind_LoanApl_UnitModel & " , " & Ind_LoanApl_LCP & " , " & Ind_LoanApl_DP & " , " & Ind_LoanApl_Term & " , " & Ind_LoanApl_AOR & " , " & vbCrLf
        sql = sql & Ind_LoanApl_Monthly_Amortization & " , " & Ind_LoanApl_Balance_FI_Perc & " , " & Ind_LoanApl_Balance_FI_Amount & " , " & vbCrLf
        sql = sql & Ind_LoanApl_Purpose & " , " & Ind_LoanApl_PlaceOfUse & " , " & Ind_LoanApl_SAE & " , " & vbCrLf
        sql = sql & Ind_Ref_Pers_Name1 & " , " & Ind_Ref_Pers_Add1 & " , " & Ind_Ref_Pers_TelNo1 & " , " & vbCrLf
        sql = sql & Ind_Ref_Pers_Name2 & " , " & Ind_Ref_Pers_Add2 & " , " & Ind_Ref_Pers_TelNo2 & " , " & vbCrLf
        sql = sql & Ind_Ref_Credit_Name1 & " , " & Ind_Ref_Credit_Add1 & " , " & Ind_Ref_Credit_TelNo1 & " , " & vbCrLf
        sql = sql & Ind_Ref_Credit_Name2 & " , " & Ind_Ref_Credit_Add2 & " , " & Ind_Ref_Credit_TelNo2 & " , " & vbCrLf
        sql = sql & Ind_BA_Bank1 & " , " & Ind_BA_Type1 & " , " & Ind_BA_AcctNo1 & " , " & Ind_BA_Bal1 & " , " & vbCrLf
        sql = sql & Ind_BA_Bank2 & " , " & Ind_BA_Type2 & " , " & Ind_BA_AcctNo2 & " , " & Ind_BA_Bal2 & " , " & vbCrLf
        sql = sql & Ind_BA_Bank3 & " , " & Ind_BA_Type3 & " , " & Ind_BA_AcctNo3 & " , " & Ind_BA_Bal3 & " , " & vbCrLf
        sql = sql & Ind_BA_Bank4 & " , " & Ind_BA_Type4 & " , " & Ind_BA_AcctNo4 & " , " & Ind_BA_Bal4 & " , 'O' )"

    Else


        sql = "UPDATE SMIS_LoanIndiv "
        sql = sql & " SET ProspectID= " & ProspectID & " , APL_No= " & Apl_no & " , AplCode= " & N2Str2Null(CUSCDE) & " , DateApplied= " & DateApplied & " , " & vbCrLf
        sql = sql & " Ind_Apl_LastName= " & Ind_Apl_LastName & " , Ind_Apl_FirstName= " & Ind_Apl_FirstName & " , Ind_Apl_MidName= " & Ind_Apl_MidName & " , Ind_Apl_Birthday= " & Ind_Apl_Birthday & " , Ind_Apl_Age= " & Ind_Apl_Age & " ,  " & vbCrLf
        sql = sql & " Ind_Sps_LastName= " & Ind_Sps_LastName & " , Ind_Sps_FirstName= " & Ind_Sps_FirstName & " , Ind_Sps_MidName= " & Ind_Sps_MidName & " , Ind_Sps_Birthday= " & Ind_Sps_Birthday & " , Ind_Sps_Age= " & Ind_Sps_Age & " ,  " & vbCrLf
        sql = sql & " Ind_Address= " & Ind_Address & " , Ind_TelNo= " & Ind_TelNo & " , Ind_CpNo= " & Ind_CpNo & " , Ind_Length_of_Stay= " & Ind_Length_of_Stay & " ,  " & vbCrLf
        sql = sql & " Ind_Ownership= " & Ind_Ownership & " , Ind_Civil_Status= " & Ind_Civil_Status & " , Ind_Citizenship= " & Ind_Citizenship & " , Ind_No_Of_dependents= " & Ind_No_Of_dependents & " ,  " & vbCrLf
        sql = sql & " Ind_Previous_Address= " & Ind_Previous_Address & " ,  " & vbCrLf
        sql = sql & " Ind_Monthly_Rental= " & Ind_Monthly_Rental & " ,  " & vbCrLf
        sql = sql & " Ind_Name_of_Landlord= " & Ind_Name_of_Landlord & " ,  " & vbCrLf
        sql = sql & " Ind_Landlord_TelNo= " & Ind_Landlord_TelNo & " ,  " & vbCrLf
        sql = sql & " Ind_Apl_EmpBusName= " & Ind_Apl_EmpBusName & " ,  Ind_Apl_Address= " & Ind_Apl_Address & " , Ind_Apl_Position= " & Ind_Apl_Position & " , Ind_Apl_TelNo= " & Ind_Apl_TelNo & " , Ind_Apl_LengthOfStay= " & Ind_Apl_LengthOfStay & " , Ind_Apl_PreviousEmp= " & Ind_Apl_PreviousEmp & " ,  Ind_Apl_PrevAddress= " & Ind_Apl_PrevAddress & " ,  " & vbCrLf
        sql = sql & " Ind_Sps_EmpBusName= " & Ind_Sps_EmpBusName & " ,  Ind_Sps_Address= " & Ind_Sps_Address & " ,  Ind_Sps_Position= " & Ind_Sps_Position & " , Ind_Sps_TelNo= " & Ind_Sps_TelNo & " , Ind_Sps_LengthOfStay= " & Ind_Sps_LengthOfStay & " , Ind_Sps_PreviousEmp= " & Ind_Sps_PreviousEmp & " ,  Ind_Sps_PrevAddress= " & Ind_Sps_PrevAddress & " ,   " & vbCrLf
        sql = sql & " Ind_MI_Applicant= " & Ind_MI_Applicant & " ,  " & vbCrLf
        sql = sql & " Ind_MI_Spouse= " & Ind_MI_Spouse & " ,  " & vbCrLf
        sql = sql & " Ind_MI_OtherIncome1Desc= " & Ind_MI_OtherIncome1Desc & " , Ind_MI_OtherIncome1Amount= " & Ind_MI_OtherIncome1Amount & " ,  " & vbCrLf
        sql = sql & " Ind_MI_OtherIncome2Desc= " & Ind_MI_OtherIncome2Desc & " , Ind_MI_OtherIncome2Amount= " & Ind_MI_OtherIncome2Amount & " ,  " & vbCrLf
        sql = sql & " Ind_MI_OtherIncome3Desc= " & Ind_MI_OtherIncome3Desc & " , Ind_MI_OtherIncome3Amount= " & Ind_MI_OtherIncome3Amount & " ,  " & vbCrLf
        sql = sql & " Ind_MI_LivingExpense= " & Ind_MI_LivingExpense & " ,  " & vbCrLf
        sql = sql & " Ind_MI_Rental= " & Ind_MI_Rental & " ,  " & vbCrLf
        sql = sql & " Ind_MI_Amortizations= " & Ind_MI_Amortizations & " ,  " & vbCrLf
        sql = sql & " Ind_LoanApl_UnitModel= " & Ind_LoanApl_UnitModel & " ,  " & vbCrLf
        sql = sql & " Ind_LoanApl_LCP= " & Ind_LoanApl_LCP & " ,  " & vbCrLf
        sql = sql & " Ind_LoanApl_DP= " & Ind_LoanApl_DP & " ,  " & vbCrLf
        sql = sql & " Ind_LoanApl_Term= " & Ind_LoanApl_Term & " ,  " & vbCrLf
        sql = sql & " Ind_LoanApl_AOR= " & Ind_LoanApl_AOR & " ,  " & vbCrLf
        sql = sql & " Ind_LoanApl_Monthly_Amortization= " & Ind_LoanApl_Monthly_Amortization & " ,  " & vbCrLf
        sql = sql & " Ind_LoanApl_Balance_FI_Perc= " & Ind_LoanApl_Balance_FI_Perc & " ,  " & vbCrLf
        sql = sql & " Ind_LoanApl_Balance_FI_Amount= " & Ind_LoanApl_Balance_FI_Amount & " ,  " & vbCrLf
        sql = sql & " Ind_LoanApl_Purpose= " & Ind_LoanApl_Purpose & " ,  " & vbCrLf
        sql = sql & " Ind_LoanApl_PlaceOfUse= " & Ind_LoanApl_PlaceOfUse & " ,  " & vbCrLf
        sql = sql & " Ind_LoanApl_SAE= " & Ind_LoanApl_SAE & " ,  " & vbCrLf
        sql = sql & " Ind_Ref_Pers_Name1= " & Ind_Ref_Pers_Name1 & " , Ind_Ref_Pers_Add1= " & Ind_Ref_Pers_Add1 & " , Ind_Ref_Pers_TelNo1= " & Ind_Ref_Pers_TelNo1 & " ,  " & vbCrLf
        sql = sql & " Ind_Ref_Pers_Name2= " & Ind_Ref_Pers_Name2 & " , Ind_Ref_Pers_Add2= " & Ind_Ref_Pers_Add2 & " , Ind_Ref_Pers_TelNo2= " & Ind_Ref_Pers_TelNo2 & " ,  " & vbCrLf
        sql = sql & " Ind_Ref_Credit_Name1= " & Ind_Ref_Credit_Name1 & " , Ind_Ref_Credit_Add1= " & Ind_Ref_Credit_Add1 & " , Ind_Ref_Credit_TelNo1= " & Ind_Ref_Credit_TelNo1 & " ,  " & vbCrLf
        sql = sql & " Ind_Ref_Credit_Name2= " & Ind_Ref_Credit_Name2 & " , Ind_Ref_Credit_Add2= " & Ind_Ref_Credit_Add2 & " , Ind_Ref_Credit_TelNo2= " & Ind_Ref_Credit_TelNo2 & " ,  " & vbCrLf
        sql = sql & " Ind_BA_Bank1= " & Ind_BA_Bank1 & " , Ind_BA_Type1= " & Ind_BA_Type1 & " , Ind_BA_AcctNo1= " & Ind_BA_AcctNo1 & " , Ind_BA_Bal1= " & Ind_BA_Bal1 & " ,  " & vbCrLf
        sql = sql & " Ind_BA_Bank2= " & Ind_BA_Bank2 & " , Ind_BA_Type2= " & Ind_BA_Type2 & " , Ind_BA_AcctNo2= " & Ind_BA_AcctNo2 & " , Ind_BA_Bal2= " & Ind_BA_Bal2 & " ,  " & vbCrLf
        sql = sql & " Ind_BA_Bank3= " & Ind_BA_Bank3 & " , Ind_BA_Type3= " & Ind_BA_Type3 & " , Ind_BA_AcctNo3= " & Ind_BA_AcctNo3 & " , Ind_BA_Bal3= " & Ind_BA_Bal3 & " ,  " & vbCrLf
        sql = sql & " Ind_BA_Bank4= " & Ind_BA_Bank4 & " , Ind_BA_Type4= " & Ind_BA_Type4 & " , Ind_BA_AcctNo4= " & Ind_BA_AcctNo4 & " , Ind_BA_Bal4= " & Ind_BA_Bal4 & vbCrLf
        sql = sql & " WHERE ID=" & LabID


    End If
    gconDMIS.Execute sql
    gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET LOGAPPLICATION=GETDATE() ,LogApplicationType='C' WHERE PROSPECTID=" & ProspectID)
    rsRefresh
    rsLoan.Find ("APL_No=" & N2Str2Null(txtApl_No))
    cmdCancel.Value = True



    If FormExist("MainForm") Then
        MainForm.ShowData
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSaveSO_Click()
    If lstCustomer.SelectedItem Is Nothing Then: Exit Sub
    ProspectID = lstCustomer.SelectedItem.ListSubItems(6)
    SearchID lstCustomer.SelectedItem.ListSubItems(7)
    ShowHidePictureBox2 pic4EditSO, False

End Sub

Public Sub SearchID(xxx)

    rsLoan.MoveFirst
    rsLoan.Find ("ID=" & xxx)
    AddorEdit = ""
    StoreMemVars

End Sub
'Upating Code       : AXP-0707200713:22
Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost", "INDIVIDUAL LOAN APPLICATION") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If MsgBox("Do You want to Un Post this Applications", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    gconDMIS.Execute ("UPDATE SMIS_LOANINDIV SET STATUS='U' WHERE ID=" & LabID)
    rsRefresh
    rsLoan.Find ("ID=" & LabID)
    StoreMemVars
    '   MessagePop RecSaveOk, "Transaction Posted", "Transaction Sucessfully Un Posted"





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:22
Private Sub cmdUpdateStatus_Click()
    Combo1.ListIndex = SetStatus(Null2String(rsLoan!lStatus))
    txtReasonNote = Null2String(rsLoan!Notes)
    ShowHidePictureBox2 picStatus, True
End Sub

Private Sub Command3_Click()
    ShowHidePictureBox2 picDocumentList, False
End Sub

'Upating Code       : AXP-0707200713:22
Private Sub Command4_Click()
    Dim Item                                                          As ListItem
    On Error GoTo ErrorCode:

    gconDMIS.Execute ("Delete from SMIS_loanDocument where aplcode=" & N2Str2Null(txtApl_No) & " AND AplType='I'")
    For Each Item In ListView1.ListItems

        If Item.Checked = True Then
            gconDMIS.Execute (" insert into SMIS_loanDocument values (" & N2Str2Null(Item.Text) & ", " & N2Str2Null(txtApl_No) & ", 'I')")
        End If
    Next
    MessagePop RecSaveOk, "Updated", "Document Listing Updated"

    ShowHidePictureBox2 picDocumentList, False





    Exit Sub
ErrorCode:
    ShowVBError

End Sub

'Upating Code       : AXP-0707200713:21
Private Sub Command6_Click()

    On Error GoTo ErrorCode:

    gconDMIS.Execute ("Update smis_loanindiv set notes= " & _
                      N2Str2Null(txtReasonNote) & ", LStatus=" & N2Str2Null(Left(Combo1, 1)) & " Where Id=" & LabID)

    '    If Combo1 = "Approved" Then
    '        If MsgBox(" Do You Want to Post This Application.", vbYesNo) = vbYes Then
    '            gconDMIS.Execute ("Update smis_loanindiv set status='P' Where Id=" & labID)
    '        End If
    '
    '    End If

    rsRefresh
    rsLoan.Find ("ID=" & LabID)
    cmdCancel.Value = True
    ShowHidePictureBox2 picStatus, False

    If FormExist("MainForm") Then
        MainForm.ShowData
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:22
Private Sub Command7_Click()
    On Error GoTo ErrorCode:

    frmSMIS_Files_Document.Show
    frmSMIS_Files_Document.ZOrder 0





    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If picStatus.Visible = True Then
            ShowHidePictureBox2 picStatus, False
        ElseIf pic4EditSO.Visible = True Then
            ShowHidePictureBox2 pic4EditSO, False
        ElseIf picDocumentList.Visible = True Then
            ShowHidePictureBox2 picDocumentList, False
        End If
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_Load()

    Me.Height = Screen.TwipsPerPixelY * 540
    picMiddles.Height = Me.ScaleHeight - picBottoms.Height - picTops.Height
    ScrollBar1.Height = picMiddles.ScaleHeight - 15
    ScrollBar1.Max = Abs(picMiddles.ScaleHeight - picIndividual.Height) + 20    '& "--" & ScrollBar1.Value

    Call AddColumnHeader("CODE,DOCUMENT NAME ", ListView1)
    Call ResizeColumnHeader(ListView1, "35,65")

    picIndividual.Enabled = False
    CenterMe frmMain, Me, 1
    initMemvars
    InitCBO
    rsRefresh

    If LoanID > 0 Then
        rsLoan.Find ("ID='" & LoanID & "'")
    End If
    If AddingLoan = True Then
        Exit Sub
    Else
        StoreMemVars
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    APLNO = vbNullString

    Set rsLoan = Nothing
    Set rsS_Model = Nothing
    AddorEdit = vbNullString
    xDateApplied = vbNullString
    Set Ctl = Nothing

    ProspectID = 0
    CUSCDE = vbNullString
    ProfileType = vbNullString
    APLNO = vbNullString

End Sub

Private Sub FormLog_EventSelected(EventName As String, xProspectID As Long, AppointmentID As Long)

    Dim temprs                                                        As ADODB.Recordset


    Set temprs = gconDMIS.Execute("SELECT * FROM CRIS_SALESAPPOINTMENTS WHERE AppointmentID=" & AppointmentID)
    If Not (temprs.EOF Or temprs.BOF) Then
        Dim oSpouse                                                   As Variant

        'txtAppInfo_App_LastName = Null2String(temprs!lastname)
        'txtAppInfo_App_FirstName = Null2String(temprs!FirstName)
        'txtAppInfo_App_MiddleName = Null2String(temprs!MiddleInitial
        'txtAppInfo_App_BirthDate.Text = Null2String(temprs!BirthDate)
        'oSpouse = Split(Null2String(temprs!Spouse), " ")
        'If UBound(oSpouse) = 1 Then
        '    txtAppInfo_Sps_FirstName = oSpouse(0)
        '    txtAppInfo_Sps_LastName = oSpouse(1)
        '    cboAppInfo_App_CivilStatus.Text = "Married"
        'ElseIf UBound(oSpouse) = 2 Then
        '    txtAppInfo_Sps_FirstName = oSpouse(0)
        '    txtAppInfo_Sps_LastName = oSpouse(1)
        '    txtAppInfo_Sps_MiddleName = oSpouse(2)
        '    cboAppInfo_App_CivilStatus.Text = "Married"
        'ElseIf UBound(oSpouse) = 0 Then

        '   cboAppInfo_App_CivilStatus.Text = "Unspecified"
        'El'se
        '   txtAppInfo_Sps_LastName = Null2String(temprs!Spouse)
        '   cboAppInfo_App_CivilStatus.Text = "Married"
        'e 'nd If
        'txtAppInfo_Address = Null2String(temprs!CustomerAdd)
        'txtAppInfo_Cellphone = Null2String(temprs!Mobile)
        'txtAppInfo_Telephone = Null2String(temprs!HomePhone)
        'txtInd_Apl_EmpBusName = Null2String(temprs!CUSCOMP)
        'txtInd_Apl_Address = Null2String(temprs!COMPANYADD)
        'txtInd_Apl_Position = Null2String(temprs!TITLE)
        'txtInd_Apl_TelNo = Null2String(temprs!TELEPHONENO)
        'cboAppInfo_AppCitizen = "Filippino"
    End If

    Unload FormLog
    Set FormLog = Nothing
End Sub

Private Sub FormSearch_NoSelectionMade()
    If rsLoan.EOF Or rsLoan.BOF Then
        Unload Me
    End If

End Sub

Private Sub FormSearch_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    initMemvars
    CUSCDE = Null2String(oCusRs!CUSCDE)
    ProspectID = Null2String(oCusRs!ProspectID)
    ProfileType = Null2String(oCusRs!ProspectType)
    cboLoan_Model.Text = Null2String(oCusRs!Variant)
    cboLoan_SAE.Text = Null2String(oCusRs!SAE)

    '    If IsNull(oCusRs!LogAppointment) = False Then
    '       MsgBox " This Prospect Has Appointment Do You Want Preview Appointment to Load Application "
    '      Set FormLog = New frmSMIS_Inquiry_ViewLog
    '     Call FormLog.ShowCustomerAppointment(ProspectID, Null2String(oCusRs!AcctName))
    '    Load FormLog
    '   FormLog.Show
    'GoTo CustomerCode:
    'Exit Sub
    'End If

    Dim temprs                                                        As ADODB.Recordset

    If CUSCDE <> "" Then
        Set temprs = gconDMIS.Execute("SELECT * FROM ALL_CUSTOMER WHERE CUSCDE=" & N2Str2Null(CUSCDE))
        If Not (temprs.EOF Or temprs.BOF) Then
            Dim oSpouse                                               As Variant
            txtAppInfo_App_LastName = Null2String(temprs!lastname)
            txtAppInfo_App_FirstName = Null2String(temprs!FIRSTNAME)
            txtAppInfo_App_MiddleName = Null2String(temprs!MiddleInitial)
            txtAppInfo_App_BirthDate.Text = Null2String(temprs!BirthDate)
            oSpouse = Split(Null2String(temprs!Spouse), " ")
            If UBound(oSpouse) = 1 Then
                txtAppInfo_Sps_FirstName = oSpouse(0)
                txtAppInfo_Sps_LastName = oSpouse(1)
                cboAppInfo_App_CivilStatus.Text = "Married"
            ElseIf UBound(oSpouse) = 2 Then
                txtAppInfo_Sps_FirstName = oSpouse(0)
                txtAppInfo_Sps_LastName = oSpouse(1)
                txtAppInfo_Sps_MiddleName = oSpouse(2)
                cboAppInfo_App_CivilStatus.Text = "Married"
            ElseIf UBound(oSpouse) = 0 Then

                cboAppInfo_App_CivilStatus.Text = "Unspecified"
            Else
                txtAppInfo_Sps_LastName = Null2String(temprs!Spouse)
                cboAppInfo_App_CivilStatus.Text = "Married"
            End If
            txtAppInfo_Address = Null2String(temprs!CUSTOMERADD)
            txtAppInfo_Cellphone = Null2String(temprs!Mobile)
            txtAppInfo_Telephone = Null2String(temprs!HOMEPHONE)
            txtInd_Apl_EmpBusName = Null2String(temprs!CUSCOMP)
            txtInd_Apl_Address = Null2String(temprs!CompanyAdd)
            txtInd_Apl_Position = Null2String(temprs!TITLE)
            txtInd_Apl_TelNo = Null2String(temprs!TelephoneNo)
            cboAppInfo_AppCitizen = "Filippino"
        End If
    Else
        Dim arName                                                    As Variant
        arName = Split(Null2String(oCusRs!AcctName), " ")

        If UBound(arName) = 1 Then
            txtAppInfo_App_LastName = arName(0)
            txtAppInfo_App_FirstName = arName(1)

        ElseIf UBound(arName) = 2 Then
            txtAppInfo_App_LastName = arName(0)
            txtAppInfo_App_FirstName = arName(1)
            txtAppInfo_App_MiddleName = arName(2)
        ElseIf UBound(arName) = 0 Then
            txtAppInfo_App_LastName = arName(0)
        End If

        txtAppInfo_Address = Null2String(oCusRs!Address)
        txtAppInfo_Cellphone = Null2String(oCusRs!Mobile)
        txtAppInfo_Telephone = Null2String(oCusRs!Telephone)
        Erase arName
    End If

CustomerCode:

    txtApl_No = GenerateCode("SMIS_LoanIndiv", "APL_No ", "00000000")
    AddorEdit = "ADD"
    picIndividual.Enabled = True
    picSaves.Visible = True
    picAdds.Visible = False

    Unload FormSearch
    Set FormSearch = Nothing
End Sub


';Private Sub FormAOR_LineAOR(NetSalesPrice As Variant, DownPayment As Variant, Term As Variant, AOR As Variant, FinBaltoFinanced As Variant, NetMoAmort As Variant)

'''''''''''
'
'End Sub


Sub InitCBO()
    Dim sql                                                           As String
    Dim AccountType                                                   As String
    With Combo1
        .AddItem "On Process"
        .AddItem "Approved"
        .AddItem "Cancelled"
        .AddItem "Disapproved"
        .AddItem "Pending"
        .ListIndex = 0
    End With

    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select descript from All_Model order by descript asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboLoan_Model.Clear
        Do While Not rsS_Model.EOF
            cboLoan_Model.AddItem Null2String(rsS_Model!DESCRIPT)
            rsS_Model.MoveNext
        Loop
    End If

    With cboAppInfo_OnwnerShip
        .Clear
        .AddItem ("Owned")
        .AddItem ("Mortgaged")
        .AddItem ("Rented")
        .AddItem ("Provided")
    End With

    With cboAppInfo_App_CivilStatus
        .Clear
        .AddItem ("Single")
        .AddItem ("Married")
        .AddItem ("Windowed")
        .AddItem ("Separated")
        .AddItem ("Unspecified")
    End With

    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select CODE,NAME from SMIS_vw_SRep order by NAME asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboLoan_SAE.Clear
        Do While Not rsS_Model.EOF
            cboLoan_SAE.AddItem Null2String(rsS_Model!Name)
            rsS_Model.MoveNext
        Loop
    End If

    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select Ind_Citizenship from SMIS_LoanIndiv group by Ind_Citizenship order by Ind_Citizenship asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboAppInfo_AppCitizen.Clear
        Do While Not rsS_Model.EOF
            cboAppInfo_AppCitizen.AddItem Null2String(rsS_Model!Ind_Citizenship)
            rsS_Model.MoveNext
        Loop
    End If

    Set rsBAType = New ADODB.Recordset

    sql = "select DISTINCT IND_BA_TYPE1  from SMIS_LOANINDIV " & vbCrLf
    sql = sql & " Union " & vbCrLf
    sql = sql & " select DISTINCT IND_BA_TYPE2  from SMIS_LOANINDIV" & vbCrLf
    sql = sql & " Union " & vbCrLf
    sql = sql & " select DISTINCT IND_BA_TYPE3  from SMIS_LOANINDIV" & vbCrLf
    sql = sql & " Union " & vbCrLf
    sql = sql & " select DISTINCT IND_BA_TYPE4  from SMIS_LOANINDIV" & vbCrLf

    Set rsBAType = gconDMIS.Execute(sql)
    If Not rsBAType.EOF And Not rsBAType.BOF Then

        rsBAType.MoveFirst
        cboInd_BA_Type1.Clear
        cboInd_BA_Type2.Clear
        cboInd_BA_Type3.Clear
        cboInd_BA_Type4.Clear
        Do While Not rsBAType.EOF
            AccountType = Null2String(rsBAType.Fields(0).Value)
            cboInd_BA_Type1.AddItem AccountType
            cboInd_BA_Type2.AddItem AccountType
            cboInd_BA_Type3.AddItem AccountType
            cboInd_BA_Type4.AddItem AccountType

            rsBAType.MoveNext
        Loop
    End If





End Sub

Sub initMemvars()
    With Me
        For Each Ctl In .ControlS
            If TypeOf Ctl Is TextBox Then
                Ctl.Text = Ctl.Tag
            End If
        Next Ctl
    End With
    optLoan_Private.Value = True
    labTStatus = ""
    labLStatus = ""
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtSearch_APL = lstCustomer.SelectedItem
    txtSearch_AplName = Trim(lstCustomer.SelectedItem.SubItems(2)) & ", " & Trim(lstCustomer.SelectedItem.SubItems(3)) & " " & Trim(lstCustomer.SelectedItem.SubItems(4))
    cmdSaveSO.Enabled = True
End Sub

Sub rsRefresh()
    Set rsLoan = New ADODB.Recordset
    rsLoan.Open "SELECT * FROM SMIS_LoanIndiv ORDER BY id DESC", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Private Sub ScrollBar1_Change()
    picIndividual.Top = 0 - ScrollBar1.Value
End Sub

Function SetStatus(xx As String) As Integer

    If xx = "O" Then
        SetStatus = 0
    ElseIf xx = "A" Then
        SetStatus = 1
    ElseIf xx = "C" Then
        SetStatus = 2
    ElseIf xx = "D" Then
        SetStatus = 3
    ElseIf xx = "P" Then
        SetStatus = 4
    End If
End Function

Function GetStatus(xx As String) As String
    If xx = "O" Then
        GetStatus = "On Process"
    ElseIf xx = "A" Then
        GetStatus = "Approved"
    ElseIf xx = "C" Then
        GetStatus = "Cancelled"
    ElseIf xx = "D" Then
        GetStatus = "Disapproved"
    ElseIf xx = "P" Then
        GetStatus = "Pending"
    End If
End Function
Sub StoreMemVars()
    If Not rsLoan.EOF And Not rsLoan.BOF Then
        txtApl_No = Null2String(rsLoan!Apl_no)
        CUSCDE = Null2String(rsLoan!AplCode)
        xDateApplied = Null2String(rsLoan!DateApplied)

        txtAppInfo_App_LastName = Null2String(rsLoan!Ind_Apl_LastName)
        txtAppInfo_App_FirstName = Null2String(rsLoan!Ind_Apl_FirstName)
        txtAppInfo_App_MiddleName = Null2String(rsLoan!Ind_Apl_MidName)
        txtAppInfo_Sps_LastName = Null2String(rsLoan!Ind_Sps_LastName)
        txtAppInfo_Sps_FirstName = Null2String(rsLoan!Ind_Sps_FirstName)
        txtAppInfo_Sps_MiddleName = Null2String(rsLoan!Ind_Sps_MidName)
        txtAppInfo_Address = Null2String(rsLoan!Ind_Address)
        txtAppInfo_App_BirthDate = Null2String(rsLoan!Ind_Apl_Birthday)
        txtAppInfo_App_Age = Null2String(rsLoan!Ind_Apl_Age)
        txtAppInfo_Sps_BirthDate = Null2String(rsLoan!Ind_Sps_Birthday)
        txtAppInfo_Sps_Age = Null2String(rsLoan!Ind_Sps_Age)
        txtAppInfo_Cellphone = Null2String(rsLoan!Ind_CpNo)
        txtAppInfo_Telephone = Null2String(rsLoan!Ind_TelNo)
        txtAppInfo_LengthOfStay = Null2String(rsLoan!Ind_Length_of_Stay)
        cboAppInfo_OnwnerShip.Text = Null2String(rsLoan!Ind_Ownership)
        cboAppInfo_App_CivilStatus.Text = Null2String(rsLoan!Ind_Civil_Status)
        cboAppInfo_AppCitizen.Text = Null2String(rsLoan!Ind_Citizenship)
        txtAppInfo_NoDependent = Null2String(rsLoan!Ind_No_Of_dependents)
        txtAppInfo_MonthlyRental = Null2String(rsLoan!Ind_Monthly_Rental)
        txtAppInfo_NameofLandlord = Null2String(rsLoan!Ind_Name_of_Landlord)
        txtAppInfo_LandlordTelNo = Null2String(rsLoan!Ind_Landlord_TelNo)
        txtAppInfo_PreviousAddress = Null2String(rsLoan!Ind_Previous_Address)

        txtInd_Apl_EmpBusName = Null2String(rsLoan!Ind_Apl_EmpBusName)
        txtInd_Apl_Address = Null2String(rsLoan!Ind_Apl_Address)
        txtInd_Apl_Position = Null2String(rsLoan!Ind_Apl_Position)
        txtInd_Apl_TelNo = Null2String(rsLoan!Ind_Apl_TelNo)
        txtInd_Apl_LengthOfStay = Null2String(rsLoan!Ind_Apl_LengthOfStay)
        txtInd_Apl_PreviousEmp = Null2String(rsLoan!Ind_Apl_PreviousEmp)
        txtInd_Apl_PrevAddress = Null2String(rsLoan!Ind_Apl_PrevAddress)

        txtSpouseEmpBusName = Null2String(rsLoan!Ind_Sps_EmpBusName)
        txtSpouseAddress = Null2String(rsLoan!Ind_Sps_Address)
        txtSpousePosition = Null2String(rsLoan!Ind_Sps_Position)
        txtSpouseTelNo = Null2String(rsLoan!Ind_Sps_TelNo)
        txtSpouseLengthOfStay = Null2String(rsLoan!Ind_Sps_LengthOfStay)
        txtSpousePreviousEmp = Null2String(rsLoan!Ind_Sps_PreviousEmp)
        txtSpousePrevAddress = Null2String(rsLoan!Ind_Sps_PrevAddress)
        txtMonthlyIncome_Applicant = Null2String(rsLoan!Ind_MI_Applicant)
        txtMonthlyIncome_Spouse = Null2String(rsLoan!Ind_MI_Spouse)
        txtMonthlyIncome_OtherIncomeDesc1 = Null2String(rsLoan!Ind_MI_OtherIncome1Desc)
        txtMonthlyIncome_OtherIncome1 = Null2String(rsLoan!Ind_MI_OtherIncome1Amount)
        txtMonthlyIncome_OtherIncomeDesc2 = Null2String(rsLoan!Ind_MI_OtherIncome2Desc)
        txtMonthlyIncome_OtherIncome2 = Null2String(rsLoan!Ind_MI_OtherIncome2Amount)
        txtMonthlyIncome_OtherIncomeDesc3 = Null2String(rsLoan!Ind_MI_OtherIncome3Desc)
        txtMonthlyIncome_OtherIncome3 = Null2String(rsLoan!Ind_MI_OtherIncome3Amount)
        txtMonthlyIncome_LivingExpense = Null2String(rsLoan!Ind_MI_LivingExpense)
        txtMonthlyIncome_Rental = Null2String(rsLoan!Ind_MI_Rental)

        ProspectID = Null2String(rsLoan!ProspectID)
        cboLoan_Model.Text = Null2String(rsLoan!Ind_LoanApl_UnitModel)
        txtLoan_UnitCost = Null2String(rsLoan!Ind_LoanApl_LCP)
        txtLoan_Downpayment = Null2String(rsLoan!Ind_LoanApl_DP)
        txtLoan_BankTerms = Null2String(rsLoan!Ind_LoanApl_Term)



        txtLoan_AORPercentage = Null2String(rsLoan!Ind_LoanApl_AOR)
        txtMonthlyIncome_Amort = Null2String(rsLoan!Ind_MI_Amortizations)
        txtLoan_MonthlyAmortization = Null2String(rsLoan!Ind_LoanApl_Monthly_Amortization)

        txtLoan_DownpaymentPerct = Null2String(rsLoan!Ind_LoanApl_Balance_FI_Perc)

        txtLoan_FinBalAmount = Null2String(rsLoan!Ind_LoanApl_Balance_FI_Amount)

        If Null2String(rsLoan!Ind_LoanApl_Purpose) = optLoan_Private.Caption Then
            optLoan_Private.Value = True
        ElseIf Null2String(rsLoan!Ind_LoanApl_Purpose) = optLoan_Business.Caption Then
            optLoan_Business.Value = True
        ElseIf Null2String(rsLoan!Ind_LoanApl_Purpose) = optLoan_Public.Caption Then
            optLoan_Public.Value = True
        End If
        txtLoan_PlaceOfUse = Null2String(rsLoan!Ind_LoanApl_PlaceOfUse)
        cboLoan_SAE = Null2String(rsLoan!Ind_LoanApl_SAE)
        txtRef_Pers_Name1 = Null2String(rsLoan!Ind_Ref_Pers_Name1)
        txtRef_Pers_Add1 = Null2String(rsLoan!Ind_Ref_Pers_Add1)
        txtRef_Pers_TelNo1 = Null2String(rsLoan!Ind_Ref_Pers_TelNo1)
        txtRef_Pers_Name2 = Null2String(rsLoan!Ind_Ref_Pers_Name2)
        txtRef_Pers_Add2 = Null2String(rsLoan!Ind_Ref_Pers_Add2)
        txtRef_Pers_TelNo2 = Null2String(rsLoan!Ind_Ref_Pers_TelNo2)
        txtRef_Credit_Name1 = Null2String(rsLoan!Ind_Ref_Credit_Name1)
        txtRef_Credit_Add1 = Null2String(rsLoan!Ind_Ref_Credit_Add1)
        txtRef_Credit_TelNo1 = Null2String(rsLoan!Ind_Ref_Credit_TelNo1)
        txtRef_Credit_Name2 = Null2String(rsLoan!Ind_Ref_Credit_Name2)
        txtRef_Credit_Add2 = Null2String(rsLoan!Ind_Ref_Credit_Add2)
        txtRef_Credit_TelNo2 = Null2String(rsLoan!Ind_Ref_Credit_TelNo2)

        txtInd_BA_Bank1 = Null2String(rsLoan!Ind_BA_Bank1)
        cboInd_BA_Type1 = Null2String(rsLoan!Ind_BA_Type1)
        txtInd_BA_AcctNo1 = Null2String(rsLoan!Ind_BA_AcctNo1)
        txtInd_BA_Bal1 = Null2String(rsLoan!Ind_BA_Bal1)

        txtInd_BA_Bank2 = Null2String(rsLoan!Ind_BA_Bank2)
        cboInd_BA_Type2 = Null2String(rsLoan!Ind_BA_Type2)
        txtInd_BA_AcctNo2 = Null2String(rsLoan!Ind_BA_AcctNo2)
        txtInd_BA_Bal2 = Null2String(rsLoan!Ind_BA_Bal2)
        txtInd_BA_Bank3 = Null2String(rsLoan!Ind_BA_Bank3)
        cboInd_BA_Type3 = Null2String(rsLoan!Ind_BA_Type3)
        txtInd_BA_AcctNo3 = Null2String(rsLoan!Ind_BA_AcctNo3)
        txtInd_BA_Bal3 = Null2String(rsLoan!Ind_BA_Bal3)
        txtInd_BA_Bank4 = Null2String(rsLoan!Ind_BA_Bank4)
        cboInd_BA_Type4 = Null2String(rsLoan!Ind_BA_Type4)
        txtInd_BA_AcctNo4 = Null2String(rsLoan!Ind_BA_AcctNo4)
        txtInd_BA_Bal4 = Null2String(rsLoan!Ind_BA_Bal4)
        LabID = rsLoan!ID
        Dim TStatus, lStatus                                          As String
        TStatus = Null2String(rsLoan!status)
        lStatus = Null2String(rsLoan!lStatus)
        labLStatus = GetStatus(lStatus)

        If Null2String(rsLoan!IsProcessed) = True Then
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = True
            cmdUnPost.Enabled = False
            cmdCancelCO.Enabled = False
        Else
            If TStatus = "P" Then
                labTStatus.Visible = True
                labTStatus.Caption = "POSTED"
                cmdEdit.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = True
                cmdUnPost.Enabled = True
                cmdDocumentCheckList.Enabled = False
                cmdUpdateStatus.Enabled = False
            ElseIf TStatus = "C" Then
                labTStatus.Caption = "CANCELLED"
                cmdEdit.Enabled = False
                cmdPost.Enabled = False
                cmdUnPost.Enabled = False
                cmdPrint.Enabled = False
                cmdCancelCO.Enabled = False
                cmdUnPost.Enabled = False
                cmdDocumentCheckList.Enabled = False
                cmdUpdateStatus.Enabled = False
            Else
                labTStatus.Visible = False
                labTStatus.Caption = ""
                cmdEdit.Enabled = True
                cmdPost.Enabled = True
                cmdPrint.Enabled = True
                cmdCancelCO.Enabled = True
                cmdUnPost.Enabled = False
                cmdDocumentCheckList.Enabled = True
                cmdUpdateStatus.Enabled = True
            End If

            If lStatus = "C" Then
                cmdCancelCO.Enabled = True
            Else
                cmdCancelCO.Enabled = False
            End If
        End If
    Else
        If AddingLoan = False Then
            ShowNoRecord
            Select Case MsgBox("There are no Loan Application(s)." & vbCrLf & "  Do You Want To Add New Record", vbYesNo Or vbQuestion Or vbDefaultButton1, App.TITLE)
                Case vbYes
                    cmdAdd.Value = True
                Case vbNo
                    Unload Me
            End Select
        End If

    End If
End Sub

Private Sub Timer1_Timer()
    If labLStatus.Caption <> "" Then
        If labLStatus.Visible = True Then
            labLStatus.Visible = False
        Else
            labLStatus.Visible = True
        End If
    End If
End Sub

Private Sub txtAppInfo_App_BirthDate_Change()
    If IsDate(txtAppInfo_App_BirthDate.Text) = True Then
        txtAppInfo_App_Age = DateDiff("YYYY", txtAppInfo_App_BirthDate.Text, Now)
    End If
End Sub

Private Sub txtAppInfo_App_BirthDate_LostFocus()
    If IsDate(txtAppInfo_App_BirthDate) = False Then: Exit Sub
    txtAppInfo_App_BirthDate.Text = FormatDateTime(txtAppInfo_App_BirthDate, vbShortDate)
End Sub

Private Sub txtAppInfo_LengthOfStay_LostFocus()
    txtAppInfo_LengthOfStay = NumericVal(txtAppInfo_LengthOfStay.Text)
End Sub

Private Sub txtAppInfo_MonthlyRental_LostFocus()
    txtAppInfo_MonthlyRental = FormatNumber(NumericVal(txtAppInfo_MonthlyRental), 2)
End Sub

Private Sub txtAppInfo_Sps_BirthDate_Change()
    If IsDate(txtAppInfo_Sps_BirthDate.Text) = True Then
        txtAppInfo_Sps_Age = DateDiff("YYYY", txtAppInfo_Sps_BirthDate.Text, Now)
    End If

End Sub

Private Sub txtAppInfo_Sps_BirthDate_LostFocus()
    If IsDate(txtAppInfo_Sps_BirthDate) = False Then: Exit Sub
    txtAppInfo_Sps_BirthDate.Text = FormatDateTime(txtAppInfo_Sps_BirthDate, vbShortDate)
End Sub


Private Sub txtFindAPL_Change()
    Dim rsSeeSO                                                       As ADODB.Recordset
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsSeeSO = New ADODB.Recordset
    Set rsSeeSO = gconDMIS.Execute("select APL_No,DateApplied,Ind_Apl_LastName ,Ind_Apl_FirstName,Ind_Apl_MidName,AplCode, ProspectID, ID from SMIS_LoanIndiv where Ind_Apl_LastName like '" & ReplaceQuote(txtFindAPL) & "%' order by Ind_Apl_LastName asc")

    If Not (rsSeeSO.EOF And rsSeeSO.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsSeeSO
        lstCustomer.Refresh
    End If
End Sub

Private Sub txtInd_Apl_LengthOfStay_LostFocus()
    txtInd_Apl_LengthOfStay = NumericVal(txtInd_Apl_LengthOfStay)
End Sub

Private Sub txtInd_BA_Bal1_GotFocus()
    If NumericVal(txtInd_BA_Bal1.Text) <= 0 Then txtInd_BA_Bal1 = ""
End Sub

Private Sub txtInd_BA_Bal1_LostFocus()
    txtInd_BA_Bal1 = FormatNumber(NumericVal(txtInd_BA_Bal1), 2)
End Sub

Private Sub txtInd_BA_Bal2_GotFocus()
    If NumericVal(txtInd_BA_Bal2.Text) <= 0 Then txtInd_BA_Bal2 = ""
End Sub

Private Sub txtInd_BA_Bal2_LostFocus()
    txtInd_BA_Bal2 = FormatNumber(NumericVal(txtInd_BA_Bal2), 2)
End Sub

Private Sub txtInd_BA_Bal3_GotFocus()
    If NumericVal(txtInd_BA_Bal3.Text) <= 0 Then txtInd_BA_Bal3 = ""
End Sub

Private Sub txtInd_BA_Bal3_LostFocus()
    txtInd_BA_Bal3 = FormatNumber(NumericVal(txtInd_BA_Bal3), 2)
End Sub

Private Sub txtInd_BA_Bal4_GotFocus()
    If NumericVal(txtInd_BA_Bal4.Text) <= 0 Then txtInd_BA_Bal4 = ""
End Sub

Private Sub txtInd_BA_Bal4_LostFocus()
    txtInd_BA_Bal4 = FormatNumber(NumericVal(txtInd_BA_Bal4), 2)
End Sub

Private Sub txtLoan_AORPercentage_Change()
    On Error Resume Next
    txtLoan_FinBalAmount_Change
End Sub

Private Sub txtLoan_AORPercentage_GotFocus()
    If NumericVal(txtLoan_AORPercentage.Text) <= 0 Then txtLoan_AORPercentage = ""

End Sub

Private Sub txtLoan_AORPercentage_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtLoan_AORPercentage_LostFocus()
    If NumericVal(txtLoan_AORPercentage.Text) <= 0 Then txtLoan_AORPercentage = "0.00"
    txtLoan_AORPercentage = FormatNumber(txtLoan_AORPercentage)
End Sub

Private Sub txtLoan_BankTerms_Change()
    On Error Resume Next
    txtLoan_FinBalAmount_Change
End Sub

Private Sub txtLoan_BankTerms_GotFocus()
    If NumericVal(txtLoan_BankTerms.Text) <= 0 Then txtLoan_BankTerms = ""

End Sub

Private Sub txtLoan_BankTerms_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtLoan_BankTerms_LostFocus()
    If NumericVal(txtLoan_BankTerms.Text) <= 0 Then txtLoan_BankTerms = "0.00"
    txtLoan_BankTerms = FormatNumber(txtLoan_BankTerms)
End Sub

Private Sub txtLoan_Downpayment_Change()
    On Error Resume Next
    If ComputebyPert = False And AddorEdit <> "" Then
        txtLoan_DownpaymentPerct = FormatNumber((NumericVal(txtLoan_Downpayment) / NumericVal(txtLoan_UnitCost)) * 100, 2)
        UpdateAmountDetails
    End If
End Sub

Private Sub txtLoan_Downpayment_GotFocus()
    If NumericVal(txtLoan_Downpayment.Text) <= 0 Then txtLoan_Downpayment = ""

End Sub

Private Sub txtLoan_Downpayment_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtLoan_Downpayment_LostFocus()
    If NumericVal(txtLoan_Downpayment.Text) <= 0 Then txtLoan_Downpayment = "0.00"
    txtLoan_Downpayment = FormatNumber(txtLoan_Downpayment)
End Sub

Private Sub txtLoan_DownpaymentPerct_Change()
    On Error Resume Next
    If ComputebyPert = True And AddorEdit <> "" Then
        txtLoan_Downpayment = FormatNumber(NumericVal(txtLoan_UnitCost) * (NumericVal(txtLoan_DownpaymentPerct) / 100), 2)
        UpdateAmountDetails
    End If
End Sub

Private Sub txtLoan_DownpaymentPerct_GotFocus()
    If NumericVal(txtLoan_DownpaymentPerct.Text) <= 0 Then txtLoan_DownpaymentPerct = ""
    ComputebyPert = True

End Sub

Private Sub txtLoan_DownpaymentPerct_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtLoan_DownpaymentPerct_LostFocus()
    If NumericVal(txtLoan_DownpaymentPerct.Text) <= 0 Then txtLoan_DownpaymentPerct = "0.00"

    txtLoan_DownpaymentPerct = FormatNumber(txtLoan_DownpaymentPerct)
    ComputebyPert = False
End Sub

Private Sub txtLoan_FinBalAmount_Change()
    On Error Resume Next
    If AddorEdit = "" Then Exit Sub
    txtLoan_MonthlyAmortization = AORVALUE(NumericVal(txtLoan_FinBalAmount), NumericVal(txtLoan_AORPercentage), NumericVal(txtLoan_BankTerms))
End Sub

Private Sub txtLoan_PlaceOfUse_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtLoan_UnitCost_Change()
    On Error Resume Next
    If AddorEdit = "" Then Exit Sub
    txtLoan_Downpayment = FormatNumber(NumericVal(txtLoan_UnitCost) * (NumericVal(txtLoan_DownpaymentPerct) / 100), 2)
    UpdateAmountDetails
End Sub

Private Sub txtLoan_UnitCost_GotFocus()
    If NumericVal(txtLoan_UnitCost.Text) <= 0 Then txtLoan_UnitCost = ""

End Sub

Private Sub txtLoan_UnitCost_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtLoan_UnitCost_LostFocus()
    If NumericVal(txtLoan_UnitCost.Text) <= 0 Then txtLoan_UnitCost = "0.00"
    txtLoan_UnitCost = FormatNumber(txtLoan_UnitCost)
End Sub

Private Sub txtMonthlyIncome_Amort_GotFocus()
    If NumericVal(txtMonthlyIncome_Amort.Text) <= 0 Then txtMonthlyIncome_Amort = ""
End Sub

Private Sub txtMonthlyIncome_Amort_LostFocus()
    txtMonthlyIncome_Amort = FormatNumber(NumericVal(txtMonthlyIncome_Amort), 2)
End Sub

Private Sub txtMonthlyIncome_Applicant_GotFocus()
    If NumericVal(txtMonthlyIncome_Applicant.Text) <= 0 Then txtMonthlyIncome_Applicant = ""
End Sub

Private Sub txtMonthlyIncome_Applicant_LostFocus()
    txtMonthlyIncome_Applicant = FormatNumber(NumericVal(txtMonthlyIncome_Applicant), 2)
End Sub

Private Sub txtMonthlyIncome_LivingExpense_GotFocus()
    If NumericVal(txtMonthlyIncome_LivingExpense.Text) <= 0 Then txtMonthlyIncome_LivingExpense = ""
End Sub

Private Sub txtMonthlyIncome_LivingExpense_LostFocus()

    txtMonthlyIncome_LivingExpense = FormatNumber(NumericVal(txtMonthlyIncome_LivingExpense), 2)
End Sub

Private Sub txtMonthlyIncome_OtherIncome1_GotFocus()
    If NumericVal(txtMonthlyIncome_OtherIncome1.Text) <= 0 Then txtMonthlyIncome_OtherIncome1 = ""
End Sub

Private Sub txtMonthlyIncome_OtherIncome1_LostFocus()
    txtMonthlyIncome_OtherIncome1 = FormatNumber(NumericVal(txtMonthlyIncome_OtherIncome1), 2)
End Sub

Private Sub txtMonthlyIncome_OtherIncome2_GotFocus()
    If NumericVal(txtMonthlyIncome_OtherIncome2.Text) <= 0 Then txtMonthlyIncome_OtherIncome2 = ""
End Sub

Private Sub txtMonthlyIncome_OtherIncome2_LostFocus()
    txtMonthlyIncome_OtherIncome2 = FormatNumber(NumericVal(txtMonthlyIncome_OtherIncome2), 2)
End Sub

Private Sub txtMonthlyIncome_OtherIncome3_GotFocus()
    If NumericVal(txtMonthlyIncome_OtherIncome3.Text) <= 0 Then txtMonthlyIncome_OtherIncome3 = ""
End Sub

Private Sub txtMonthlyIncome_OtherIncome3_LostFocus()
    txtMonthlyIncome_OtherIncome3 = FormatNumber(NumericVal(txtMonthlyIncome_OtherIncome3), 2)
End Sub

Private Sub txtMonthlyIncome_Rental_GotFocus()

    If NumericVal(txtMonthlyIncome_LivingExpense.Text) <= 0 Then txtMonthlyIncome_LivingExpense = ""
End Sub

Private Sub txtMonthlyIncome_Rental_LostFocus()
    txtMonthlyIncome_Rental = FormatNumber(NumericVal(txtMonthlyIncome_Rental), 2)
End Sub

Private Sub txtMonthlyIncome_Spouse_GotFocus()
    If NumericVal(txtMonthlyIncome_Spouse.Text) <= 0 Then txtMonthlyIncome_Spouse = ""
End Sub

Private Sub txtMonthlyIncome_Spouse_LostFocus()
    txtMonthlyIncome_Spouse = FormatNumber(NumericVal(txtMonthlyIncome_Spouse), 2)
End Sub

Sub UpdateAmountDetails()
    If AddorEdit = "" Then Exit Sub
    Dim A, b, C, D, E
    A = NumericVal(txtLoan_UnitCost)
    b = NumericVal(txtLoan_Downpayment)
    C = NumericVal(txtLoan_BankTerms)
    D = NumericVal(txtLoan_FinBalAmount)
    E = NumericVal(txtLoan_AORPercentage)
    txtLoan_FinBalAmount = FormatNumber((A - b), 2)

End Sub


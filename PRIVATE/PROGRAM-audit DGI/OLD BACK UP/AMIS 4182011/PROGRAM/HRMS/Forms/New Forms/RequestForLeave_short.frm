VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMS_RequestForLeave_short 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REQUEST FOR LEAVE & OVERTIME"
   ClientHeight    =   5880
   ClientLeft      =   1110
   ClientTop       =   2625
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "RequestForLeave_short.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8430
   Begin Crystal.CrystalReport rpt 
      Left            =   180
      Top             =   5130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   -150
      ScaleHeight     =   1005
      ScaleWidth      =   8520
      TabIndex        =   43
      Top             =   4770
      Width           =   8520
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
         Left            =   6420
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RequestForLeave_short.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Unpost this Transaction"
         Top             =   120
         Width           =   705
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Disapprove"
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
         Left            =   5730
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RequestForLeave_short.frx":0A21
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":0B73
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Unpost this Transaction"
         Top             =   120
         Width           =   705
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Approve"
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
         Left            =   5040
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RequestForLeave_short.frx":0EB8
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":100A
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Post this Transaction"
         Top             =   120
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
         Left            =   7800
         MouseIcon       =   "RequestForLeave_short.frx":132F
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":1481
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Exit Window"
         Top             =   120
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
         Left            =   7110
         MouseIcon       =   "RequestForLeave_short.frx":17E7
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":1939
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Print this Record"
         Top             =   120
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
         Left            =   4350
         MouseIcon       =   "RequestForLeave_short.frx":1C9F
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":1DF1
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Delete Selected Record"
         Top             =   120
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
         Left            =   3660
         MouseIcon       =   "RequestForLeave_short.frx":211C
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":226E
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Edit Selected Record"
         Top             =   120
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
         Left            =   2970
         MouseIcon       =   "RequestForLeave_short.frx":25CA
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":271C
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Add Record"
         Top             =   120
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
         Left            =   2280
         MouseIcon       =   "RequestForLeave_short.frx":2A2F
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":2B81
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Find a Record"
         Top             =   120
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
         Left            =   1590
         MouseIcon       =   "RequestForLeave_short.frx":2E7B
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":2FCD
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Move to Next Record"
         Top             =   120
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
         Left            =   900
         MouseIcon       =   "RequestForLeave_short.frx":3325
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":3477
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Move to Previous Record"
         Top             =   120
         Width           =   705
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5310
      ScaleHeight     =   855
      ScaleWidth      =   3060
      TabIndex        =   40
      Top             =   4860
      Width           =   3060
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
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
         Left            =   2340
         MouseIcon       =   "RequestForLeave_short.frx":37D6
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":3928
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Cancel"
         Top             =   30
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
         Left            =   1650
         MouseIcon       =   "RequestForLeave_short.frx":3C66
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave_short.frx":3DB8
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   4845
      Left            =   30
      ScaleHeight     =   4845
      ScaleWidth      =   11115
      TabIndex        =   0
      Top             =   0
      Width           =   11115
      Begin VB.Frame Frame4 
         Caption         =   "Employee Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   11055
         Begin VB.ComboBox cboEmployeeNumber 
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
            Height          =   345
            Left            =   150
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   5
            Text            =   "Combo2"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txtEmployeeName 
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
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   420
            Width           =   3945
         End
         Begin MSComCtl2.DTPicker dtFiling 
            Height          =   375
            Left            =   6060
            TabIndex        =   7
            Top             =   390
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39513
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Name"
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
            Index           =   0
            Left            =   2070
            TabIndex        =   2
            Top             =   210
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "No"
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
            Index           =   2
            Left            =   180
            TabIndex        =   3
            Top             =   210
            Width           =   210
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Filing Date"
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
            Left            =   6090
            TabIndex        =   4
            Top             =   180
            Width           =   840
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Reason For Request"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3795
         Left            =   6030
         TabIndex        =   38
         Top             =   930
         Width           =   2265
         Begin RichTextLib.RichTextBox txtReason 
            Height          =   3555
            Left            =   60
            TabIndex        =   24
            Top             =   180
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   6271
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"RequestForLeave_short.frx":4108
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Request Dates && Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3795
         Left            =   2940
         TabIndex        =   31
         Top             =   930
         Width           =   3045
         Begin VB.ComboBox cboCutOff 
            Height          =   345
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   360
            Width           =   2835
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1095
            Left            =   60
            ScaleHeight     =   1095
            ScaleWidth      =   2955
            TabIndex        =   37
            Top             =   2550
            Width           =   2955
            Begin VB.TextBox txtReasonInvalid 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   0
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   20
               Top             =   0
               Width           =   2895
            End
         End
         Begin MSComCtl2.DTPicker dtFromDate 
            Height          =   345
            Left            =   120
            TabIndex        =   18
            Top             =   930
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   609
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker dtToDate 
            Height          =   345
            Left            =   1560
            TabIndex        =   19
            Top             =   930
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker dtReportingDate 
            Height          =   345
            Left            =   120
            TabIndex        =   23
            Top             =   2130
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker dtFromTime 
            Height          =   345
            Left            =   120
            TabIndex        =   21
            Top             =   1530
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            Format          =   20709378
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker dtToTime 
            Height          =   345
            Left            =   1560
            TabIndex        =   22
            Top             =   1530
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            Format          =   20709378
            CurrentDate     =   39513
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cut-Off"
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
            Left            =   120
            TabIndex        =   62
            Top             =   180
            Width           =   585
         End
         Begin VB.Label LABOTID 
            AutoSize        =   -1  'True
            Caption         =   "Label3"
            Height          =   225
            Left            =   1620
            TabIndex        =   61
            Top             =   2280
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label LABLEAVEID 
            AutoSize        =   -1  'True
            Caption         =   "Label6"
            Height          =   225
            Left            =   2340
            TabIndex        =   60
            Top             =   2280
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label labRTDesc1 
            AutoSize        =   -1  'True
            Caption         =   "From"
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
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   435
         End
         Begin VB.Label labRTDesc2 
            AutoSize        =   -1  'True
            Caption         =   "To"
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
            Left            =   1560
            TabIndex        =   33
            Top             =   720
            Width           =   210
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Reporting Date"
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
            Left            =   120
            TabIndex        =   36
            Top             =   1890
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Time From"
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
            Left            =   120
            TabIndex        =   34
            Top             =   1320
            Width           =   900
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Time To"
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
            Left            =   1560
            TabIndex        =   35
            Top             =   1320
            Width           =   675
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Applicable To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3795
         Left            =   0
         TabIndex        =   8
         Top             =   930
         Width           =   2925
         Begin VB.Frame Frame1 
            Caption         =   "Status Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2445
            Left            =   60
            TabIndex        =   27
            Top             =   1290
            Width           =   2805
            Begin VB.Label labREQNO 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   90
               TabIndex        =   16
               Top             =   1980
               Width           =   2625
            End
            Begin VB.Label Label1 
               Caption         =   "Approved By"
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
               Index           =   0
               Left            =   90
               TabIndex        =   28
               Top             =   570
               Width           =   2595
            End
            Begin VB.Label Label1 
               Caption         =   "Date Approved"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   29
               Top             =   1185
               Width           =   2595
            End
            Begin VB.Label labApprovedBy 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   90
               TabIndex        =   14
               Top             =   795
               Width           =   2625
            End
            Begin VB.Label labDateApproved 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   90
               TabIndex        =   15
               Top             =   1410
               Width           =   2625
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Application No"
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
               Left            =   90
               TabIndex        =   30
               Top             =   1740
               Width           =   2595
            End
            Begin VB.Label labStatus 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label3"
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
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   2625
            End
         End
         Begin VB.OptionButton optLeave 
            Caption         =   "&Leave"
            Height          =   225
            Left            =   270
            TabIndex        =   9
            Top             =   300
            Width           =   795
         End
         Begin VB.OptionButton optOverTime 
            Caption         =   "&Overtime"
            Height          =   225
            Left            =   1140
            TabIndex        =   10
            Top             =   300
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Frame fraOvertime 
            Caption         =   "Select Your Overtime Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   60
            TabIndex        =   26
            Top             =   600
            Width           =   2805
            Begin VB.ComboBox cboOvertime 
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
               Left            =   90
               TabIndex        =   12
               Text            =   "cboOvertime"
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.Frame fraLeave 
            Caption         =   "Leave Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   60
            TabIndex        =   25
            Top             =   600
            Width           =   2805
            Begin VB.ComboBox cboLeaveType 
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
               Left            =   150
               TabIndex        =   11
               Top             =   240
               Width           =   2595
            End
         End
      End
      Begin VB.Label LABID 
         Caption         =   "0"
         Height          =   345
         Left            =   3810
         TabIndex        =   39
         Top             =   4950
         Width           =   1125
      End
   End
   Begin VB.PictureBox picSearch 
      BorderStyle     =   0  'None
      Height          =   6165
      Left            =   30
      ScaleHeight     =   6165
      ScaleWidth      =   11145
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   11145
      Begin MSComctlLib.ListView ListView1 
         Height          =   5115
         Left            =   -30
         TabIndex        =   59
         Top             =   600
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   9022
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date Filed"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Req No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "EMP No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Action Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Leave Dates"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         CausesValidation=   0   'False
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
         Left            =   10590
         TabIndex        =   58
         Top             =   60
         Width           =   375
      End
      Begin VB.TextBox TXTSEARCH 
         Height          =   405
         Left            =   1590
         TabIndex        =   56
         Top             =   60
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Employee  Name"
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox picStatus1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   2100
      ScaleHeight     =   4485
      ScaleWidth      =   4245
      TabIndex        =   63
      Top             =   180
      Visible         =   0   'False
      Width           =   4275
      Begin VB.PictureBox picML 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   30
         ScaleHeight     =   735
         ScaleWidth      =   4185
         TabIndex        =   65
         Top             =   1290
         Visible         =   0   'False
         Width           =   4185
         Begin VB.ComboBox cboML 
            Height          =   345
            ItemData        =   "RequestForLeave_short.frx":4189
            Left            =   30
            List            =   "RequestForLeave_short.frx":4193
            TabIndex        =   71
            Top             =   270
            Width           =   4095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Maternity Leave Type"
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
            Index           =   1
            Left            =   30
            TabIndex        =   66
            Top             =   30
            Width           =   1770
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   2730
         ScaleHeight     =   885
         ScaleWidth      =   1440
         TabIndex        =   64
         Top             =   3600
         Width           =   1440
         Begin VB.CommandButton cmdStatusCancel 
            Caption         =   "&Cancel"
            CausesValidation=   0   'False
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
            Left            =   720
            MouseIcon       =   "RequestForLeave_short.frx":41A9
            MousePointer    =   99  'Custom
            Picture         =   "RequestForLeave_short.frx":42FB
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Cancel"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdStatusOK 
            Caption         =   "&Ok"
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
            MouseIcon       =   "RequestForLeave_short.frx":4639
            MousePointer    =   99  'Custom
            Picture         =   "RequestForLeave_short.frx":478B
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Save this Record"
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.ComboBox cboApprovedBy 
         Height          =   345
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   900
         Width           =   4095
      End
      Begin VB.TextBox txtNotes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   73
         Text            =   "RequestForLeave_short.frx":4ADB
         Top             =   2220
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtApproved 
         Height          =   345
         Left            =   1470
         TabIndex        =   67
         Top             =   390
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   609
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   39513
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Notes"
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
         Index           =   6
         Left            =   60
         TabIndex        =   74
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Action Date"
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
         Index           =   5
         Left            =   90
         TabIndex        =   72
         Top             =   420
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Action Person"
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
         Index           =   4
         Left            =   60
         TabIndex        =   70
         Top             =   660
         Width           =   1170
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   -30
         TabIndex        =   68
         Top             =   0
         Width           =   4245
         _Version        =   655364
         _ExtentX        =   7488
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "Status"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   0
         ForeColor       =   16777215
      End
   End
End
Attribute VB_Name = "frmHRMS_RequestForLeave_short"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRequest                                                         As ADODB.Recordset
Dim ADDOREDIT                                                         As String
Dim STATUSX                                                           As String
Dim XEMPNO As String
Public Event SelectionMade()
Public Sub SelectSQl(XXX As String, xxempno As String)
    
    Set rsRequest = New ADODB.Recordset
    rsRequest.Open XXX, gconDMIS, adOpenForwardOnly, adLockReadOnly
    XEMPNO = xxempno
             
End Sub

Function GenerateCode() As String
    Dim rsID                                                          As ADODB.Recordset
    Set rsID = gconDMIS.Execute("SELECT MAX(ISNULL(REQNO,0)) AS HRMS_REQUESTLEAVE_OT FROM HRMS_REQUESTLEAVE_OT ")
    If rsID.FIELDS(0).Value = 0 Then
        GenerateCode = Format(1, "000000")
    Else
        GenerateCode = Format(val(N2Str2Zero(rsID(0))) + 1, "000000")
    End If
    Set rsID = Nothing
End Function
Function GetDayNo(YY As String) As String
    Dim rsDAYSNO As New ADODB.Recordset
    Dim xdays As Integer
    
    Set rsDAYSNO = gconDMIS.Execute("Select days_no from hrms_leavemaster where leave_code = '" & YY & "'")

    If Not (rsDAYSNO.BOF And rsDAYSNO.EOF) Then
        'xdays = Trim(rsDAYSNO!DAYS_NO)
        xdays = N2Str2Zero(rsDAYSNO!DAYS_NO)
    End If
    
    GetDayNo = xdays
    
End Function
Function GetLeaveCode(XXX As String)
    Dim rsLeaveLook                                                   As ADODB.Recordset
    Set rsLeaveLook = gconDMIS.Execute("SELECT LEAVE_CODE FROM HRMS_LEAVEMASTER WHERE ID=" & XXX)
    If Not (rsLeaveLook.BOF Or rsLeaveLook.EOF) Then
        GetLeaveCode = LTrim(RTrim(Null2String(rsLeaveLook!LEAVE_CODE)))
    End If
    Set rsLeaveLook = Nothing
End Function

Function GetLeaveDescription(XXX As String)
    Dim rsLeaveLook                                                   As ADODB.Recordset
    Set rsLeaveLook = gconDMIS.Execute("SELECT * FROM HRMS_LEAVEMASTER WHERE LEAVE_CODE='" & Repleys(XXX) & "'")
    If Not (rsLeaveLook.BOF Or rsLeaveLook.EOF) Then
        GetLeaveDescription = LTrim(RTrim(Null2String(rsLeaveLook!LEAVE_desc)))
    End If
    Set rsLeaveLook = Nothing
End Function

Function GetOverTimeCode(XXX As String)
    Dim rsOTLook                                                      As ADODB.Recordset
    Set rsOTLook = gconDMIS.Execute("SELECT PAY_CODE FROM HRMS_OTCODES WHERE ID=" & XXX)
    If Not (rsOTLook.BOF Or rsOTLook.EOF) Then
        GetOverTimeCode = LTrim(RTrim(Null2String(rsOTLook!PAY_CODE)))
    End If
    Set rsOTLook = Nothing
End Function

Function GetOverTimeDescription(XXX As String)
    Dim rsOTLook                                                      As ADODB.Recordset
    Set rsOTLook = gconDMIS.Execute("SELECT * FROM HRMS_OTCODES WHERE PAY_CODE='" & Repleys(XXX) & "'")
    If Not (rsOTLook.BOF Or rsOTLook.EOF) Then
        GetOverTimeDescription = LTrim(RTrim(Null2String(rsOTLook!PAY_DESC)))
    End If
    Set rsOTLook = Nothing
End Function

Function SelectCombo(C As ComboBox) As Integer
    If C.ListCount = 0 Then
        SelectCombo = -1
        C.Text = ""
        Exit Function
    End If
    Dim I                                                             As Long
    For I = 0 To C.ListCount - 1
        If UCase(C.list(I)) = UCase(Trim(C.Text)) Then
            SelectCombo = I
            Exit Function
        End If
    Next
    SelectCombo = -1
    C.Text = ""
End Function

Sub ResizeGrid()
    With ListView1
        .ColumnHeaders(1).Width = .Width * 0.11
        .ColumnHeaders(2).Width = .Width * 0.08
        .ColumnHeaders(3).Width = .Width * 0.08
        .ColumnHeaders(4).Width = .Width * 0.25
        .ColumnHeaders(5).Width = .Width * 0.08
        .ColumnHeaders(6).Width = .Width * 0.2
        .ColumnHeaders(7).Width = .Width * 0.07
        .ColumnHeaders(8).Width = .Width * 0.12
    End With

End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSTMP As New ADODB.Recordset
    Dim ITEM As ListItem
    
    If XXX = "" Then
        'Listview_Loadval ListView1.ListItems, gconDMIS.Execute("SELECT R.DTE_FILING, R.REQNO, R.EMPNO, E.Lastname + ', ' +  E.Firstname, R.REQTYPE, R.REQDESC, R.STATUS,R.DTE_ACTION, R.ID, R.DTE_FROM  FROM HRMS_EMPINFO E INNER JOIN HRMS_REQUESTLEAVE_OT  R ON E.EMPNO = R.EMPNO ORDER BY R.DTE_FILING ")
        Set RSTMP = gconDMIS.Execute("SELECT R.DTE_FILING as a1, R.REQNO as a2, R.EMPNO as a3, E.Lastname + ', ' +  E.Firstname as fname, R.REQTYPE as a4, R.REQDESC as a5, R.STATUS as a6, R.DTE_ACTION as a7, R.ID as a8, R.DTE_FROM as a9, r.dte_to as a10  FROM HRMS_EMPINFO E INNER JOIN HRMS_REQUESTLEAVE_OT  R ON E.EMPNO = R.EMPNO ORDER BY R.DTE_FILING ")
    Else
        'Listview_Loadval ListView1.ListItems, gconDMIS.Execute("SELECT R.DTE_FILING, R.REQNO, R.EMPNO, E.Lastname + ', ' +  E.Firstname, R.REQTYPE, R.REQDESC, R.STATUS,R.DTE_ACTION, R.ID, R.DTE_FROM  FROM HRMS_EMPINFO E INNER JOIN HRMS_REQUESTLEAVE_OT  R ON E.EMPNO = R.EMPNO WHERE E.LASTNAME + ' ' +  E.FIRSTNAME LIKE'" & Repleys(XXX) & "%' ORDER BY DTE_FILING")
        Set RSTMP = gconDMIS.Execute("SELECT R.DTE_FILING as a1, R.REQNO as a2, R.EMPNO as a3, E.Lastname + ', ' +  E.Firstname as fname, R.REQTYPE as a4, R.REQDESC as a5, R.STATUS as a6, R.DTE_ACTION as a7, R.ID as a8, R.DTE_FROM as a9, r.dte_to as a10  FROM HRMS_EMPINFO E INNER JOIN HRMS_REQUESTLEAVE_OT  R ON E.EMPNO = R.EMPNO WHERE E.LASTNAME + ' ' +  E.FIRSTNAME LIKE'" & Repleys(XXX) & "%' ORDER BY DTE_FILING")
    End If
    ListView1.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = ListView1.ListItems.Add(, , Null2String(RSTMP!A1))
            ITEM.SubItems(1) = Null2String(RSTMP!a2)
            ITEM.SubItems(2) = Null2String(RSTMP!a3)
            ITEM.SubItems(3) = Null2String(RSTMP!FName)
            ITEM.SubItems(4) = Null2String(RSTMP!a4)
            ITEM.SubItems(5) = Null2String(RSTMP!a5)
            ITEM.SubItems(6) = Null2String(RSTMP!a6)
            ITEM.SubItems(7) = Null2String(RSTMP!a7)
            ITEM.SubItems(8) = Null2String(RSTMP!a8)
            ITEM.SubItems(9) = Null2String(RSTMP!a9) & "-" & Null2String(RSTMP!a10)
        
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub GetEmployeeDetails()
    txtEmployeeName = ""
    Dim rsEmployee                                                    As ADODB.Recordset
    Set rsEmployee = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE EMPNO='" & Repleys(cboEmployeeNumber) & "'")
    If Not (rsEmployee.BOF Or rsEmployee.EOF) Then
        txtEmployeeName = Null2String(rsEmployee!lastname) & " ," & Null2String(rsEmployee!FIRSTNAME)
    End If
    Set rsEmployee = Nothing
End Sub

Sub INITCBO()
'    cboCUTOFF.Clear
'    cboCUTOFF.AddItem "1st Cut-Off"
'    cboCUTOFF.AddItem "2nd Cut-Off"
'    Combo_Loadval cboEmployeeNumber, gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A'")
'
'    Dim rsot                                                          As ADODB.Recordset
'    Set rsot = gconDMIS.Execute("SELECT LTRIM(RTRIM(PAY_DESC)) AS OTCODE, ID FROM HRMS_OTCODES ORDER BY PAY_DESC ASC")
'    While Not rsot.EOF
'        cboOvertime.AddItem (Null2String(rsot!OTCODE))
'        cboOvertime.ItemData(cboOvertime.NewIndex) = rsot!ID
'        rsot.MoveNext
'    Wend
'    Set rsot = gconDMIS.Execute("SELECT LEAVE_DESC, ID FROM HRMS_LEAVEMASTER ORDER BY 1 ASC")
'    While Not rsot.EOF
'        cboLeaveType.AddItem (Null2String(rsot!LEAVE_desc))
'        cboLeaveType.ItemData(cboLeaveType.NewIndex) = rsot!ID
'        rsot.MoveNext
'    Wend
'    Combo_Loadval cboApprovedBy, gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME  FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A'")
'    Set rsot = Nothing



 Dim rsot                                                          As ADODB.Recordset
    Dim SEX As String
        
    SEX = GETEMPSEXCODE(cboEmployeeNumber)
    
    Set rsot = gconDMIS.Execute("SELECT LTRIM(RTRIM(PAY_DESC)) AS OTCODE, ID FROM HRMS_OTCODES ORDER BY PAY_DESC ASC")
    While Not rsot.EOF
        cboOvertime.AddItem (Null2String(rsot!OTCODE))
        cboOvertime.ItemData(cboOvertime.NewIndex) = rsot!ID
        rsot.MoveNext
    Wend
    Set rsot = gconDMIS.Execute("SELECT LEAVE_DESC, ID FROM HRMS_LEAVEMASTER ORDER BY 1 ASC")
    While Not rsot.EOF
        
        If SEX = "M" Then
            If Trim(rsot!LEAVE_desc) <> "MATERNITY LEAVE" Then
                cboLeaveType.AddItem (Null2String(rsot!LEAVE_desc))
                cboLeaveType.ItemData(cboLeaveType.NewIndex) = rsot!ID
            Else
                
            'do nothing
            End If
        Else
            If Trim(rsot!LEAVE_desc) <> "PATERNITY LEAVE" Then
                cboLeaveType.AddItem (Null2String(rsot!LEAVE_desc))
                cboLeaveType.ItemData(cboLeaveType.NewIndex) = rsot!ID
            Else
                    ' do nothing
            End If
        End If
        ' -----------------
        rsot.MoveNext
    Wend
    
    
    Combo_Loadval cboApprovedBy, gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME  FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A'")
    Set rsot = Nothing


End Sub
Function GETEMPSEXCODE(EMP As String) As String
    Dim sqltxt As String
    Dim rsEmpSEx As New ADODB.Recordset
    Dim XSEX As String
    
    sqltxt = "Select SEX from hrms_empinfo where empno = '" & EMP & "'"
    Set rsEmpSEx = gconDMIS.Execute(sqltxt)
    If Not (rsEmpSEx.BOF And rsEmpSEx.EOF) Then
        XSEX = Trim(rsEmpSEx!SEX)
    Else
        XSEX = "F" 'default
    End If
    
    GETEMPSEXCODE = XSEX

End Function

Sub rsrefresh()
    Set rsRequest = New ADODB.Recordset
    rsRequest.Open "SELECT * FROM HRMS_REQUESTLEAVE_OT ORDER BY REQNO DESC", gconDMIS, adOpenKeyset
End Sub

Sub StoreMemVars()
    If Not rsRequest.EOF Or Not rsRequest.BOF Then
        labID = rsRequest!ID
'        If NumericVal(rsRequest!CUT_OFF) = 1 Then
'            cboCutOff.ListIndex = 0
'        Else
'            cboCutOff.ListIndex = 1
'        End If
'
        cboEmployeeNumber = Null2String(rsRequest!EMPNO)
        labREQNO = rsRequest!REQNO
        
        If Null2String(rsRequest!REQTYPE) = "O" Then
            optOverTime.Value = True
            cboOvertime = Null2String(rsRequest!reqdesc)
        Else
            optLeave.Value = True
            cboLeaveType.Text = Null2String(rsRequest!reqdesc)
        End If
        
        GetEmployeeDetails
        dtFromDate = Null2String(rsRequest!DTE_FROM)
        dtFromTime = Null2String(rsRequest!OT_FROM)
        dtToDate = Null2String(rsRequest!dte_to)
        dtFiling = Null2String(rsRequest!DTE_FILING)
        
        If IsDate(rsRequest!OT_TO) = True Then
            dtToTime = TimeValue(rsRequest!OT_TO)
        End If

        
        
        If IsDate(rsRequest!DTE_REPORTING) = True Then
            dtReportingDate = DateValue(rsRequest!DTE_REPORTING)
        End If

        txtReason = Null2String(rsRequest!REASON_REQ)
        txtReasonInvalid = Null2String(rsRequest!NOTES)
        labApprovedBy = Null2String(rsRequest!ApprovedBy)
        labDateApproved = Null2String(rsRequest!DTE_ACTION)
        cmdPrint.Enabled = True

        If Null2String(rsRequest!STATUS) = "A" Then
            labStatus = "APPROVED"
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdPrint.Enabled = True
            cmdPost.Enabled = False
            cmdUnPost.Enabled = True
            cmdCancelCO.Enabled = True
            cmdDelete.Enabled = False
        ElseIf Null2String(rsRequest!STATUS) = "D" Then
            labStatus = "DISAPPROVED"
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdPrint.Enabled = True
            cmdPost.Enabled = True
            cmdCancelCO.Enabled = True
            cmdUnPost.Enabled = False
        
        ElseIf Null2String(rsRequest!STATUS) = "C" Then
            labStatus = "CANCELLED"
            cmdEdit.Enabled = False
            cmdUnPost.Enabled = False
            cmdPost.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPrint.Enabled = True
            cmdDelete.Enabled = False
    
        Else
            labStatus = "NOT PROCESSED"
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            cmdPrint.Enabled = True
        End If
        
        
        If rsRequest!CUT_OFF = 1 Then cboCutOff.Text = "1st Cut-Off"
        If rsRequest!CUT_OFF = 2 Then cboCutOff.Text = "2nd Cut-Off"
    Else
        ShowNoRecord
        cmdAdd.Value = True
        cboEmployeeNumber = XEMPNO
    End If
End Sub

Private Sub cboApprovedBy_Change()
On Error Resume Next
 cboApprovedBy.ListIndex = SelectCombo(cboApprovedBy)
End Sub

Private Sub cboApprovedBy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdStatusOK_Click
End If
End Sub

Private Sub cboEmployeeNumber_Change()
    GetEmployeeDetails
    cboLeaveType.Clear
    INITCBO
    
End Sub

Private Sub cboEmployeeNumber_Click()
    cboEmployeeNumber_Change
    'Command2.Visible = True
End Sub

Private Sub cboEmployeeNumber_LostFocus()
    cboEmployeeNumber.ListIndex = SelectCombo(cboEmployeeNumber)
End Sub

Private Sub cboLeaveType_Change()
    LABLEAVEID = 0
    If cboLeaveType.ListIndex <> -1 Then
        LABLEAVEID = cboLeaveType.ItemData(cboLeaveType.ListIndex)
    End If
End Sub

Private Sub cboLeaveType_Click()
    cboLeaveType_Change
End Sub

Private Sub cboLeaveType_LostFocus()
    cboLeaveType.ListIndex = SelectCombo(cboLeaveType)
    cboLeaveType_Change
End Sub

Private Sub cboOvertime_Change()
    LABOTID = 0
    If cboOvertime.ListIndex <> -1 Then
        LABOTID = cboOvertime.ItemData(cboOvertime.ListIndex)
    End If
End Sub

Private Sub cboOvertime_Click()
    cboOvertime_Change
End Sub

Private Sub cboOvertime_LostFocus()
    LABOTID = 0
    cboOvertime.ListIndex = SelectCombo(cboOvertime)
    cboOvertime_Change
End Sub

Private Sub cmdAdd_Click()
    ADDOREDIT = "ADD"
    InitMemvars
   
    picAdds.Visible = False
    picSaves.Visible = True
    picMain.Enabled = True

    cboEmployeeNumber.Text = XEMPNO
    cboEmployeeNumber.Enabled = False
    dtFiling.Value = Date
    dtFromTime.Value = "08:00:00 AM":   dtToTime.Value = "05:00:00 PM"
    labREQNO = GenerateCode
End Sub

Private Sub cmdCancel_Click()
    ADDOREDIT = ""
    picAdds.Visible = True
    picSaves.Visible = False
    picMain.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdCancelCO_Click()
STATUSX = "C": picStatus1.Visible = True: picStatus1.ZOrder 0
End Sub

Private Sub cmdDelete_Click()
    If ShowConfirmDelete = True Then
        gconDMIS.Execute ("DELETE FROM HRMS_REQUESTLEAVE_OT WHERE ID = " & labID)
        rsrefresh
        StoreMemVars
    End If
End Sub

Private Sub cmdEdit_Click()
    ADDOREDIT = "EDIT"
    cboEmployeeNumber.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    picMain.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    picSearch.ZOrder 0
    picSearch.Visible = True
    FillSearchGrid ""
End Sub

Private Sub cmdNext_Click()
    rsRequest.MoveNext
    If rsRequest.EOF Then
        rsRequest.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPost_Click()
    'picStatus1.Visible = True
    'picStatus1.ZOrder 0

    STATUSX = "A"
    picStatus1.Visible = True
    picStatus1.ZOrder 0
    cboApprovedBy.SetFocus
    dtApproved.Value = Date
    txtNOTES.Text = ""
    
    If cboLeaveType.Text = UCase("Maternity Leave") Then
        cboML.ListIndex = 0
        picML.Visible = True
    Else
        picML.Visible = False
    End If

End Sub

Private Sub cmdPrevious_Click()
    rsRequest.MovePrevious
    If rsRequest.BOF Then
        rsRequest.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub
Function ISHOLIDAY(xdate) As Boolean
    Dim sqltxt As String
    Dim RSHOLIDAY As New ADODB.Recordset
    Dim XMONTH, xDay As Integer
    
    xDay = CInt(Day(xdate))
    XMONTH = CInt(MONTH(xdate))
    
    sqltxt = "Select manth,deyt,description from hrms_holiday_list"
    sqltxt = sqltxt & " where manth = '" & XMONTH & "' and deyt = '" & xDay & "'"
    Set RSHOLIDAY = gconDMIS.Execute(sqltxt)
    If Not (RSHOLIDAY.EOF And RSHOLIDAY.BOF) Then
        ISHOLIDAY = True
    Else
        ISHOLIDAY = False
        
    End If
End Function


Private Sub cmdPrint_Click()
    If optLeave.Value = True Then
        rpt.Formulas(0) = "CompanyName='" & COMPANY_NAME & "'"
        rpt.Formulas(1) = "CompanyAddress='" & COMPANY_ADDRESS & "'"
        rpt.Formulas(2) = "DEPARTMENTHEAD='" & APPROVED_BY & "'"
        rpt.Formulas(3) = "NOTEDBY='" & NOTED_BY & "'"
        rpt.Formulas(4) = "REQUESTEDBY='" & txtEmployeeName & "'"
        PrintSQLReport rpt, HRMS_REPORT_PATH & "RequestForLeave.rpt", "{LO.REQNO}=" & N2Str2Null(labREQNO), DMIS_REPORT_Connection, 1
       
    Else
        rpt.Formulas(0) = "CompanyName='" & COMPANY_NAME & "'"
        rpt.Formulas(1) = "CompanyAddress='" & COMPANY_ADDRESS & "'"
        rpt.Formulas(2) = "DEPARTMENTHEAD='" & APPROVED_BY & "'"
        rpt.Formulas(3) = "NOTEDBY='" & NOTED_BY & "'"
        rpt.Formulas(4) = "REQUESTEDBY='" & txtEmployeeName & "'"
        PrintSQLReport rpt, HRMS_REPORT_PATH & "RequestForOT.rpt", "{LO.REQNO}=" & N2Str2Null(labREQNO), DMIS_REPORT_Connection, 1
    End If
End Sub

Private Sub cmdSave_Click()
    Dim VREQNO, VREQTYPE, VREQCODE, VEMPNO, VDTE_FROM, VDTE_TO, VOT_FROM, VOT_TO, VDTE_REPORTING, VDTE_FILING, VREQ_BY, VREASON_REQ, VUSERCODE, VLASTUPDATED, VREQDESC
    Dim SQL                                                           As String
    Dim vCUTOFF                                                       As Integer
    Dim MM                                                            As Integer
    Dim YY                                                            As Integer
    Dim checkrange                                                    As ADODB.Recordset
    Dim days                                                          As Integer
    Dim getrange                                                      As Integer
    Dim xCODE                                                         As String
    Dim sqltxt                                                        As String
    Dim CHECK                                                         As Boolean
    Dim xAVE                                                          As Integer
    Dim xdate                                                         As String
    Dim xNAME                                                         As String
    Dim Msg As Integer
     
    If cboCutOff.Text = "" Then
        ShowIsRequiredMsg "Choose a Cut-Off"
        cboCutOff.SetFocus
        Exit Sub
    End If

    If txtReason.Text = "" Then
        MsgBox "Pls specify the reason for leave"
        txtReason.SetFocus
        Exit Sub
    End If
    
    If txtEmployeeName.Text = "" Then
        MsgBox "Pls.. Select Employee ", vbInformation
        cboEmployeeNumber.SetFocus
        Exit Sub
    End If
    
    If dtFromDate > dtToDate Then
        MsgBox "Invalid Setup of Date ", vbInformation
        Exit Sub
    End If
    
    If dtToDate > dtReportingDate Then
        MsgBox "Invalid Reporting Date ", vbInformation
        Exit Sub
    End If
    
   If ISHOLIDAY(dtFromDate) = True Then
        MsgBox "Selected (From Date) is Holiday", vbInformation, "HRMS"
        dtFromDate.SetFocus
        Exit Sub
   End If
  
  If Weekday(dtFromDate) = vbSunday Then
        MsgBox "Selected (From Date) is Sunday", vbInformation, "HRMS"
        dtFromDate.SetFocus
        Exit Sub
  End If
     
  If Weekday(dtFromDate) = vbSaturday Then
        MsgBox "Selected (From Date) is Saturday", vbInformation, "HRMS"
        dtFromDate.SetFocus
        Exit Sub
  End If
    
  If ISHOLIDAY(dtToDate) = True Then
        MsgBox "Selected (To Date) is Holiday", vbInformation, "HRMS"
        dtToDate.SetFocus
        Exit Sub
  End If
  
  If Weekday(dtToDate) = vbSunday Then
        MsgBox "Selected (To Date) is Sunday", vbInformation, "HRMS"
        dtToDate.SetFocus
        Exit Sub
  End If
     
  If Weekday(dtToDate) = vbSaturday Then
        MsgBox "Selected (To Date) is Saturday", vbInformation, "HRMS"
        dtToDate.SetFocus
        Exit Sub
  End If
  
  
  
  
  If optLeave.Value = True Then
  
  
          Set checkrange = gconDMIS.Execute("select leave_code from hrms_leavemaster where leave_desc = '" & Trim(cboLeaveType.Text) & "'")
                If Not (checkrange.BOF And checkrange.EOF) Then
                    xCODE = Trim(checkrange!LEAVE_CODE)
        
                End If
            
             
          xAVE = GETAVTYPE(cboEmployeeNumber, GETTYPE(cboLeaveType))
          CHECK = GotValidate(xCODE, cboEmployeeNumber)
          xNAME = GETEMPNAME(cboEmployeeNumber)
            
          If CHECK = True Then
        
                If MsgBox(xNAME & " has only have " & xAVE & " of MAXIMUM " & cboLeaveType & " EXCEED LIMIT" & vbCrLf & _
                "Continue this process will result to deduction. Proceed?", vbQuestion + vbYesNo, "Error") = vbYes Then
                   'proceed
                Else
                    Call cmdCancel_Click
                    Exit Sub
                End If
          End If
    
End If
   
   
   If cboCutOff.Text = "1st Cut-Off" Then vCUTOFF = 1
   If cboCutOff.Text = "2nd Cut-Off" Then vCUTOFF = 2
    
    MM = MONTH(dtFromDate)
    YY = YEAR(dtFromDate)

    VREQNO = N2Str2Null(labREQNO)
    If optLeave.Value = True Then
        VREQTYPE = N2Str2Null("L")
        VREQCODE = N2Str2Null(GetLeaveCode(LABLEAVEID))
        VREQDESC = N2Str2Null(cboLeaveType)
    Else
        
        VREQTYPE = N2Str2Null("O")
        VREQCODE = N2Str2Null(GetOverTimeCode(LABOTID))
        VREQDESC = N2Str2Null(cboOvertime)
    End If

    VEMPNO = N2Str2Null(cboEmployeeNumber)
    VDTE_FROM = N2Str2Null(dtFromDate.Value)
    VDTE_TO = N2Str2Null(dtToDate.Value)
    VOT_FROM = N2Str2Null(TimeValue(dtFromTime.Value))
    VOT_TO = N2Str2Null(TimeValue(dtToTime.Value))
    VDTE_REPORTING = N2Str2Null(dtReportingDate.Value)
    VDTE_FILING = N2Str2Null(dtFiling.Value)
    VREQ_BY = N2Str2Null(cboEmployeeNumber)
    VREASON_REQ = N2Str2Null(txtReason.Text)
    VUSERCODE = N2Str2Null(LOGCODE)
    VLASTUPDATED = N2Str2Null(Now)

    If ADDOREDIT = "ADD" Then
        gconDMIS.Execute ("INSERT INTO HRMS_REQUESTLEAVE_OT( REQNO, REQTYPE, REQCODE, REQDESC, EMPNO, DTE_FROM, DTE_TO, OT_FROM, OT_TO, DTE_REPORTING, DTE_FILING, REQ_BY, REASON_REQ, USERCODE, LASTUPDATED, CUT_OFF, PAY_MONTH, PAY_YEAR)VALUES (" & _
                          VREQNO & _
                          "," & VREQTYPE & _
                          "," & VREQCODE & _
                          "," & VREQDESC & _
                          "," & VEMPNO & _
                          "," & VDTE_FROM & _
                          "," & VDTE_TO & _
                          "," & VOT_FROM & _
                          "," & VOT_TO & _
                          "," & VDTE_REPORTING & _
                          "," & VDTE_FILING & _
                          "," & VREQ_BY & _
                          "," & VREASON_REQ & _
                          "," & VUSERCODE & _
                          "," & VLASTUPDATED & _
                          "," & vCUTOFF & _
                          "," & MM & _
                          "," & YY & ")")
    Else
        gconDMIS.Execute ("UPDATE HRMS_REQUESTLEAVE_OT  SET " & _
                        " REQNO = " & VREQNO & "," & _
                        " REQTYPE = " & VREQTYPE & "," & _
                        " REQDESC = " & VREQDESC & "," & _
                        " REQCODE = " & VREQCODE & "," & _
                        " EMPNO = " & VEMPNO & "," & _
                        " DTE_FROM = " & VDTE_FROM & "," & _
                        " DTE_TO = " & VDTE_TO & "," & _
                        " OT_FROM = " & VOT_FROM & "," & _
                        " OT_TO = " & VOT_TO & "," & _
                        " DTE_REPORTING = " & VDTE_REPORTING & "," & _
                        " DTE_FILING = " & VDTE_FILING & "," & _
                        " REQ_BY = " & VREQ_BY & "," & _
                        " REASON_REQ = " & VREASON_REQ & "," & _
                        " USERCODE = " & VUSERCODE & "," & _
                        " LASTUPDATED = " & VLASTUPDATED & "," & _
                        " CUT_OFF = " & vCUTOFF & "," & _
                        " PAY_MONTH = " & MM & "," & _
                        " PAY_YEAR = " & YY & _
                        " WHERE ID=" & labID)

    End If

    rsrefresh
    rsRequest.Find ("REQNO='" & labREQNO & "'")
    CmdCancel.Value = True
    cmdCancelCO.Enabled = True
    cmdPost.Enabled = True
    cmdUnPost.Enabled = True
End Sub
Function GETEMPNAME(XEMPNO As String) As String
    Dim sqltxt As String
    Dim rsEMPNAME As New ADODB.Recordset
    
    sqltxt = "Select FIRSTNAME + ' ' +  LASTNAME as FNAME from hrms_empinfo"
    sqltxt = sqltxt & " where empno = '" & XEMPNO & "'"
    
    Set rsEMPNAME = gconDMIS.Execute(sqltxt)
    If Not (rsEMPNAME.BOF And rsEMPNAME.EOF) Then
        GETEMPNAME = Trim(rsEMPNAME!FName)
    End If
    
    Set rsEMPNAME = Nothing
End Function
Function GETAVTYPE(XEMPNO As String, TYPE_LEAVE As String) As Integer
    Dim RSTMP As New ADODB.Recordset
    Dim sqltxt As String
        
    sqltxt = "Select available from HRMS_LEAVE where EMPLNO = '" & XEMPNO & "'"
    sqltxt = sqltxt & " and [type] = '" & TYPE_LEAVE & "'"
    Set RSTMP = gconDMIS.Execute(sqltxt)
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GETAVTYPE = Trim(RSTMP!Available)
    
    Else
        GETAVTYPE = GetDayNo(TYPE_LEAVE)
    End If
    
    Set RSTMP = Nothing
End Function
Function GETTYPE(XDESC) As String
    Dim sqltxt As String
    Dim rsType As New ADODB.Recordset
    
    sqltxt = "Select leave_code from hrms_leavemaster where leave_desc = '" & XDESC & "' "
    Set rsType = gconDMIS.Execute(sqltxt)
    If Not (rsType.EOF And rsType.BOF) Then
        GETTYPE = Trim(rsType!LEAVE_CODE)
    End If

    Set rsType = Nothing
End Function

Function GotValidate(XXX As String, EMP As String) As Boolean
    Dim sqltxt As String
    Dim RSTMP As New ADODB.Recordset
    Dim var1, var2 As Integer
    Dim xydatediff As Integer
    
    var1 = 0: var2 = 0
    xdatediff = DateDiff("D", dtFromDate, dtToDate) + 1
    sqltxt = "Select available from hrms_leave where [type] = '" & XXX & "'"
    sqltxt = sqltxt & " and emplno = '" & EMP & "' "
    
    Set RSTMP = gconDMIS.Execute(sqltxt)
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        var1 = Trim(RSTMP!Available)
    Else
        var2 = GetDayNo(XXX)
    End If
    
    If xdatediff > CInt(var1) And xdatediff > CInt(var2) Then
       GotValidate = True
    Else
       GotValidate = False
    End If
    
End Function

Function GETTOTALDAYS(XXX As String, EMP As String) As Integer
    Dim sqltxt As String
    Dim RSTMP As New ADODB.Recordset
    Dim var1, var2 As Integer
    Dim xydatediff As Integer
    
    var1 = 0: var2 = 0
    xdatediff = DateDiff("D", dtFromDate, dtToDate) + 1
    sqltxt = "Select available from hrms_leave where [type] = '" & XXX & "'"
    sqltxt = sqltxt & " and emplno = '" & EMP & "' "
    
    Set RSTMP = gconDMIS.Execute(sqltxt)
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        var1 = Null2String(RSTMP!Available)
        GETTOTALDAYS = var1
        
    Else
        
        var2 = Null2String(GetDayNo(XXX))
        GETTOTALDAYS = var2
    End If
    
    
    
End Function

Function remaining(XXX As String, EMP As String) As Boolean
    Dim sqltxt As String
    Dim RSTMP As New ADODB.Recordset
    Dim var1, var2 As Integer
    Dim xydatediff As Integer
    
    var1 = 0: var2 = 0
    xdatediff = DateDiff("D", dtFromDate, dtToDate) + 1
    sqltxt = "Select available from hrms_leave where [type] = '" & XXX & "'"
    sqltxt = sqltxt & " and emplno = '" & EMP & "' "
    
    Set RSTMP = gconDMIS.Execute(sqltxt)
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        var1 = Trim(RSTMP!Available)
    Else
        var1 = GetDayNo(XXX)
    End If
    
    
End Function
Private Sub cmdStatusCancel_Click()
picStatus1.Visible = False: picStatus1.ZOrder 1: STATUSX = ""
End Sub

Private Sub cmdStatusOK_Click()


    Dim rsLeaveCodes As New ADODB.Recordset
    Dim sqltxt, codedesc As String
    Dim xCODE As String
    Dim bool As Boolean
    Dim Y As String
    
    On Error GoTo Errorcode
    sqltxt = "select reqcode,reqdesc,empno from hrms_requestleave_ot where "
    sqltxt = sqltxt & "reqcode IN(select leave_code from hrms_leavemaster) "
    sqltxt = sqltxt & "and req_by = '" & Trim(cboEmployeeNumber.Text) & "' "
    sqltxt = sqltxt & "and reqdesc = '" & Trim(cboLeaveType.Text) & "'"
    
    Set rsLeaveCodes = gconDMIS.Execute(sqltxt)
    If Not (rsLeaveCodes.BOF And rsLeaveCodes.EOF) Then
            xCODE = Trim(rsLeaveCodes!reqcode)
    End If

'    msg = GetDayNo(xCODE)
'
'
'    If optLeave.Value = True Then
'        If msg = 0 Then
'            MsgBox "Please Setup first Leave Codes in Table tab", vbInformation
'        Exit Sub
'        End If
'    End If

    If cboApprovedBy.ListIndex = -1 Then
        MsgBox "Please Select Proper Approver From The List!", vbInformation
        Exit Sub
    End If


    If STATUSX = "A" Then
        If MsgBox("Are You Sure You Want to Approve This Application", vbInformation + vbYesNo) = vbNo Then Exit Sub
        vDTE_APPROVED = N2Str2Null(dtApproved)
        vAPPROVEDBY = N2Str2Null(cboApprovedBy)
        vNotes = N2Str2Null(txtNOTES)
        gconDMIS.Execute ("UPDATE HRMS_REQUESTLEAVE_OT SET DTE_ACTION=" & vDTE_APPROVED & " ,STATUS='A' , APPROVEDBY=" & vAPPROVEDBY & " ,NOTES=" & vNotes & ", ML_TYPE = " & N2Str2Null(cboML) & " where id=" & labID)
        rsrefresh
        rsRequest.Find ("REQNO='" & labREQNO & "'")
        StoreMemVars
        
        'UPDATE BY   : MJP 04142009 0506PM-------------------------------------------------
        'DESCRIPTION : LEAVE MAINTENANCE MODULE
            'Call DeductToLeave(STATUSX)
        'DESCRIPTION : LEAVE MAINTENANCE MODULE
        'UPDATE BY   : MJP 04142009 0506PM-------------------------------------------------
        
        cmdStatusCancel_Click
        cmdPost.Enabled = False
        If optOverTime.Value = True Then: APPROVEOT
        If optLeave.Value = True Then: APPROVEDLEAVE
    ElseIf STATUSX = "D" Then
        If MsgBox("Are You Sure You Want to Dis-Approve This Application", vbInformation + vbYesNo) = vbNo Then Exit Sub
        vDTE_APPROVED = N2Str2Null(dtApproved)
        vAPPROVEDBY = N2Str2Null(cboApprovedBy)
        vNotes = N2Str2Null(txtNOTES)
        gconDMIS.Execute ("UPDATE HRMS_REQUESTLEAVE_OT SET DTE_ACTION=" & vDTE_APPROVED & " ,STATUS='D' , APPROVEDBY=" & vAPPROVEDBY & " ,NOTES=" & vNotes & " where id=" & labID)
        rsrefresh
        rsRequest.Find ("REQNO='" & labREQNO & "'")
        StoreMemVars
        
        'UPDATE BY   : MJP 04142009 0506PM-------------------------------------------------
        'DESCRIPTION : LEAVE MAINTENANCE MODULE
            'Call DeductToLeave(STATUSX)
        'DESCRIPTION : LEAVE MAINTENANCE MODULE
        'UPDATE BY   : MJP 04142009 0506PM-------------------------------------------------
        
        cmdStatusCancel_Click
        If optOverTime.Value = True Then: DISAPPROVEOT
        If optLeave.Value = True Then: Call DISAPPROVEDLEAVE(xCODE)
        'cmdUnPost.Enabled = False
        
    Else
        If MsgBox("Are You Sure You Want to Cancel This Application", vbInformation + vbYesNo) = vbNo Then Exit Sub
        vDTE_APPROVED = N2Str2Null(dtApproved)
        vAPPROVEDBY = N2Str2Null(cboApprovedBy)
        vNotes = N2Str2Null(txtNOTES)
        gconDMIS.Execute ("UPDATE HRMS_REQUESTLEAVE_OT SET DTE_ACTION=" & vDTE_APPROVED & " ,STATUS='C' , APPROVEDBY=" & vAPPROVEDBY & " ,NOTES=" & vNotes & " where id=" & labID)
        rsrefresh
        rsRequest.Find ("REQNO='" & labREQNO & "'")
        StoreMemVars
        cmdStatusCancel_Click
        If optOverTime.Value = True Then: DISAPPROVEOT
        If optLeave.Value = True Then: Call CANCELLEAVE(xCODE)
        'cmdCancelCO.Enabled = False
    End If
Errorcode:

End Sub

Private Sub cmdStatusOK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdStatusOK_Click
End If
End Sub

Private Sub cmdUnPost_Click()
STATUSX = "D": picStatus1.Visible = True: picStatus1.ZOrder 0
txtNOTES.Text = ""
End Sub

Private Sub Command1_Click()
    picSearch.ZOrder 1
    picSearch.Visible = False
End Sub

Private Sub dtFromDate_LostFocus()
  
  If ISHOLIDAY(dtFromDate) = True Then
        MsgBox "Selected (From Date) is Holiday", vbInformation, "HRMS"
        dtFromDate.SetFocus
        Exit Sub
  End If
  
  If Weekday(dtFromDate) = vbSunday Then
        MsgBox "Selected (From Date) is Sunday", vbInformation, "HRMS"
        dtFromDate.SetFocus
        Exit Sub
  End If
     
  If Weekday(dtFromDate) = vbSaturday Then
        MsgBox "Selected (From Date) is Saturday", vbInformation, "HRMS"
        dtFromDate.SetFocus
        Exit Sub
  End If
End Sub


Private Sub dtReportingDate_LostFocus()
'    If dtReportingDate < dtToDate Then
'        MsgBox "Invalid Reporting Date", vbInformation
'        dtReportingDate.SetFocus
'    End If
    
    If ISHOLIDAY(dtReportingDate) = True Then
        MsgBox "Selected Reporting Date is Holiday", vbInformation, "HRMS"
        dtReportingDate.SetFocus
        Exit Sub
    End If
    
    If Weekday(dtFromDate) = vbSunday Then
        MsgBox "Date selected is Sunday", vbInformation, "HRMS"
        dtReportingDate.SetFocus
        Exit Sub
    End If
     
    If Weekday(dtFromDate) = vbSaturday Then
        MsgBox "Date selected is Saturday", vbInformation, "HRMS"
        dtReportingDate.SetFocus
        Exit Sub
    End If
    
    
End Sub

Private Sub dtToDate_LostFocus()
  If dtFromDate > dtToDate Then
        MsgBox "Invalid Date ", vbInformation
  End If
    
    
  If ISHOLIDAY(dtToDate) = True Then
        MsgBox "Selected (To Date) is Holiday", vbInformation, "HRMS"
        dtToDate.SetFocus
        Exit Sub
  End If
  
  If Weekday(dtToDate) = vbSunday Then
        MsgBox "Selected (To Date) is Sunday", vbInformation, "HRMS"
        dtToDate.SetFocus
        Exit Sub
  End If
     
  If Weekday(dtToDate) = vbSaturday Then
        MsgBox "Selected (To Date) is Saturday", vbInformation, "HRMS"
        dtToDate.SetFocus
        Exit Sub
  End If
    
    
End Sub

Private Sub dtToTime_LostFocus()
    If dtFromTime > dtToTime Then
        MsgBox "Invalid Time ", vbInformation
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitMemvars
    'INITCBO
    Call INITCBO
    Call INITEMP
    optLeave.Value = True
    picAdds.Visible = True
    picSaves.Visible = False
    picMain.Enabled = False
    'rsrefresh
    
    Set rsRequest = New ADODB.Recordset
    rsRequest.Open "SELECT * FROM HRMS_REQUESTLEAVE_OT WHERE EMPNO  = '" & XEMPNO & "'", gconDMIS, adOpenKeyset

    
    StoreMemVars
    ResizeGrid
    'cmdCancelCO.Enabled = True
    'cmdPost.Enabled = True
    'cmdUnPost.Enabled = True
     Screen.MousePointer = 0
    
End Sub
Sub INITEMP()
    Combo_Loadval cboEmployeeNumber, gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A'")
End Sub

Private Sub InitMemvars()
    cboCutOff.Clear
    cboCutOff.AddItem "1st Cut-Off"
    cboCutOff.AddItem "2nd Cut-Off"

    cboOvertime.ListIndex = -1
    cboLeaveType.ListIndex = -1
    cboEmployeeNumber = ""
    txtDepartment = ""
    txtEmployeeName = ""
    txtReason = ""
    txtReasonInvalid.Text = ""
    dtFromDate = LOGDATE
    dtFromTime = LOGTIME
    dtToDate = LOGDATE
    dtToTime = LOGTIME
    dtFiling = LOGDATE
    dtReportingDate = LOGDATE
    '    dtReportingDate.MinDate = LOGDATE
    optLeave_Click
    LABOTID = 0
    'LABLEAVEID = 0
    labApprovedBy = ""
    labDateApproved = ""
    labStatus = ""
    labREQNO = ""
End Sub



Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then
        Exit Sub
    End If
    On Error Resume Next
    rsRequest.MoveFirst
    rsRequest.Find ("ID = " & ListView1.SelectedItem.ListSubItems(8).Text)
    StoreMemVars
    picSearch.ZOrder 1
    picSearch.Visible = False
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ListView1_DblClick
End Sub

Private Sub optLeave_Click()
    If optLeave.Value = True Then
        fraLeave.Visible = True
        fraOvertime.Visible = False
        If cboLeaveType.ListCount > 0 Then
            cboLeaveType.ListIndex = 1
        End If
        labRTDesc1 = "From"
        dtToDate.Enabled = True
        dtReportingDate.Enabled = True
        dtFromTime.Value = "08:00:00 AM":   dtToTime.Value = "05:00:00 PM"
    End If
End Sub

Private Sub optOverTime_Click()
    If optOverTime.Value = True Then
        fraLeave.Visible = False
        fraOvertime.Visible = True
        If cboOvertime.ListCount > 0 Then
            cboOvertime.ListIndex = 1
        End If
        labRTDesc1 = "Over Time Date"
        dtToDate.Enabled = False
        dtReportingDate.Enabled = False
        dtFromTime.Value = "06:00:00 PM":   dtToTime.Value = "09:00:00 PM"
    End If
End Sub



Private Sub txtNotes_Change()
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtsearch_Change()
    FillSearchGrid txtSearch
End Sub
Sub DeductToLeave(XSTAT As String)
    Dim DAY_CNT As Integer
    Dim RSTMP As New ADODB.Recordset
    Dim XUSED As Integer
    Dim XAVAIL As Integer
    DAY_CNT = DateDiff("d", dtFromDate, dtToDate) + 1
    
    If XSTAT = "A" Then
        Set RSTMP = gconDMIS.Execute("SELECT (AVAILABLE - USED) AS REMAIN, USED, AVAILABLE FROM HRMS_LEAVE WHERE EMPLNO = " & cboEmployeeNumber.Text & _
            " AND TYPE = " & N2Str2Null(GetLEaveType(cboLeaveType)) & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            XUSED = NumericVal(RSTMP!used)
            If RSTMP!REMAIN < DAY_CNT Then
                MsgBox "Employee dont have sufficient leave available", vbInformation, "HRMS"
                Exit Sub
            End If
        End If
        Set RSTMP = Nothing
        
        gconDMIS.Execute ("UPDATE HRMS_LEAVE SET USED = " & XUSED + DAY_CNT & _
            " WHERE EMPLNO = " & cboEmployeeNumber & _
            " AND TYPE = " & N2Str2Null(GetLEaveType(cboLeaveType)) & "")
    Else
        Set RSTMP = gconDMIS.Execute("SELECT (AVAILABLE - USED) AS REMAIN, USED, AVAILABLE FROM HRMS_LEAVE WHERE EMPLNO = " & cboEmployeeNumber & _
            " AND TYPE = " & N2Str2Null(GetLEaveType(cboLeaveType)) & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            XUSED = NumericVal(RSTMP!used)
        End If
        Set RSTMP = Nothing
        
        gconDMIS.Execute ("UPDATE HRMS_LEAVE SET USED = " & XUSED - DAY_CNT & _
            " WHERE EMPLNO = " & cboEmployeeNumber & _
            " AND TYPE = " & N2Str2Null(GetLEaveType(cboLeaveType)) & "")
    End If
End Sub
Sub APPROVEOT()
    Dim rsot                                                          As ADODB.Recordset
    Dim RSEMPINOF                                                     As ADODB.Recordset
    Dim TaymWork                                                      As Double
    Dim RperHar                                                       As Double
    Dim DailyRate                                                     As Double
    Dim OTRATE                                                        As Double
    Dim rsotrate                                                      As ADODB.Recordset
    Dim ot_rate                                                       As Double
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim vCUTOFF                                                       As Integer
    Dim MM                                                            As Integer
    Dim YY                                                            As Integer

    Set RSEMPINOF = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE EMPNO=" & N2Str2Null(cboEmployeeNumber))
    If Not (RSEMPINOF.BOF And RSEMPINOF.EOF) Then
        BasicPay = N2Str2Zero(RSEMPINOF!BASICSALARY)
        EMPLIVIL = Null2String(RSEMPINOF!EMPLEVEL)
    Else
        BasicPay = 0
    End If

    vHOURS = NumericVal(DateDiff("n", dtFromTime.Value, dtToTime.Value))
    If vHOURS > 0 Then
        vHOURS = vHOURS / 60
    End If
    If optLeave.Value = True Then
        vCODE = N2Str2Null(GetLeaveCode(LABLEAVEID))
    Else
        Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_REQUESTLEAVE_OT WHERE REQNO = '" & labREQNO.Caption & "'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            vCODE = N2Str2Null(RSTMP!reqcode)
            vCUTOFF = Null2String(RSTMP!CUT_OFF)
            MM = Null2String(RSTMP!PAY_MONTH)
            YY = Null2String(RSTMP!PAY_YEAR)
        End If
        'vCODE = N2Str2Null(GetOverTimeCode(LABOTID))
    End If
    Set rsotrate = gconDMIS.Execute("select pay_rate  from HRMS_OTCodes where pay_code=" & vCODE)
    If Not rsotrate.EOF Or Not rsotrate.BOF Then
        ot_rate = NumericVal(rsotrate!pay_rate)
    End If

    DailyRate = (BasicPay * 12) / 314
    RperHar = DailyRate / 8
    VRATE = ((vHOURS * RperHar) * CDbl(ot_rate))


    Set rsot = gconDMIS.Execute("SELECT REQ_CODE FROM HRMS_OVERTIME WHERE REQ_CODE='" & labREQNO & "'")
    If rsot.EOF Or rsot.BOF Then
        gconDMIS.Execute "Insert into HRMS_Overtime " & _
                         "(EMPLEVEL, Empno, Ocode, Deyt, Deyt2, Totalhr, Amount, Holiday, Justification, REQ_CODE, CUT_OFF, PAY_MONTH, PAY_YEAR) " & _
                         "values ('" & EMPLIVIL & _
                         "'," & N2Str2Null(cboEmployeeNumber) & _
                         "," & vCODE & ", " & N2Date2Null(dtFromDate) & _
                         "," & N2Date2Null(dtFromDate) & _
                         "," & vHOURS & _
                         "," & VRATE & _
                         "," & VRATE & _
                         "," & N2Str2Null(txtNOTES) & _
                         ",'" & labREQNO & _
                         "'," & vCUTOFF & _
                         "," & MM & _
                         "," & YY & ")"
    Else
        gconDMIS.Execute "update HRMS_Overtime set" & _
                       " EMPLEVEL = '" & EMPLIVIL & "'," & _
                       " empno = " & N2Str2Null(cboEmployeeNumber) & "," & _
                       " ocode = " & vCODE & "," & _
                       " deyt = " & N2Date2Null(dtFromDate) & "," & _
                       " deyt2 = " & N2Date2Null(dtFromDate) & "," & _
                       " totalhr = " & vHOURS & "," & _
                       " amount = " & VRATE & "," & _
                       " Holiday  = " & VRATE & "," & _
                       " REQ_CODE  = '" & labREQNO & "'," & _
                       " Justification = " & N2Str2Null(txtNOTES) & "," & _
                       " CUT_OFF = " & vCUTOFF & "," & _
                       " PAY_MONTH = " & MM & "," & _
                       " PAY_YEAR = " & YY & _
                       " where REQ_CODE = '" & labREQNO & "'"

    End If
End Sub
Sub DISAPPROVEOT()
    gconDMIS.Execute ("DELETE FROM HRMS_OVERTIME WHERE REQ_CODE='" & labREQNO & "'")
End Sub

Public Sub CANCELLEAVE(X As String)

    Dim sqltxt As String
    Dim xminus As Integer
    Dim rsCancel As New ADODB.Recordset
    Dim xAV As Integer
    Dim XUSED As Integer
    Dim xGetA, xGetU As Integer
    
    xminus = DateDiff("D", dtFromDate, dtToDate) + 1
    xAV = 0: XUSED = 0
    Set rsCancel = gconDMIS.Execute("Select available,used from hrms_leave where emplno = '" & cboEmployeeNumber & "' and [type] = '" & X & "' ")
    If Not (rsCancel.BOF And rsCancel.EOF) Then
        xAV = Trim(rsCancel!Available)
        XUSED = Trim(rsCancel!used)
    End If
    
    xGetA = xAV + xminus
    xGetU = XUSED - xminus
    
    '--- update hrms_leave
    sqltxt = "Update hrms_leave set available = '" & xGetA & "', used = '" & xGetU & "'"
    sqltxt = sqltxt & " where emplno = '" & cboEmployeeNumber & "' and [type] = '" & X & "'"
    
    gconDMIS.Execute (sqltxt)
    
    Set rsCancel = Nothing
End Sub

Function GetLEaveType(XDESC As String) As String
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT LEAVE_CODE FROM HRMS_LEAVEMASTER WHERE LEAVE_DESC = " & N2Str2Null(XDESC) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetLEaveType = Null2String(RSTMP!LEAVE_CODE)
    End If
    Set RSTMP = Nothing
End Function

Function TRACKDATE(dtefrom As Date, dteto As Date, dayz As Integer) As Date
    Dim I As Integer
    
    I = 0
    For I = 1 To dayz
        dtefrom = dtefrom + 1
    Next
    
    TRACKDATE = dtefrom
End Function

Private Sub DISAPPROVEDLEAVE(xCODE)
'nothing to do here
End Sub
Public Sub APPROVEDLEAVE()

 On Error GoTo Errorcode
    
    Dim rsLeaveEntry                                    As New ADODB.Recordset
    Dim rsGetType                                       As New ADODB.Recordset
    Dim RSGetRemLeave                                   As New ADODB.Recordset
    Dim rsSETUPDEDUCTION                                As New ADODB.Recordset
    Dim rsgetlevel                                      As New ADODB.Recordset
    Dim rsvalid                                         As New ADODB.Recordset

    Dim sqlInsert                                       As String
    Dim sqltxt As String
    Dim yemplevel As String
    Dim newAve As Integer
    Dim newUsed As Integer
    Dim yEmpIno As String
    Dim yType As String
    Dim yAve, yUsed As Integer
    Dim date_from As String
    Dim date_to As String
    Dim GetUsed As Integer
    Dim rperday As Integer
    Dim maxleave As Integer
    Dim val As Integer
    
    Dim CHECK As Boolean
    Dim checkrange                                      As New ADODB.Recordset
    Dim CXDATE                                          As Integer
    Dim CYDATE                                          As Integer
    Dim CZDATE                                          As Integer
    Dim MM                                              As String
    Dim YY                                              As String
    Dim VCUT_OFF                                        As Integer
    Dim xCODES As String
    Dim xAVE As Integer
    Dim xdate As String
    Dim xNAME As String
      
    'GET NUMBER OF LEAVE- ----
   
   Dim CNT As Integer
   Dim avail As Integer
   Dim datefroms As String
   Dim datetos As String
   Dim SQL As String
   Dim RSTMP As New ADODB.Recordset
   Dim var1 As Integer
   Dim var2 As Integer
   Dim xydatediff As Integer
   var1 = 0
   CNT = 1
   avail = 0
    
    Dim MAXIMUMLIMIT As Integer
    Dim FORDEDUCT As Integer
    Dim BILANG As Integer
    Dim DEDUCTDATE As Date
    
'    MAXIMUMLIMIT = GETTOTALDAYS(GETTYPE(cboLeaveType), cboEmployeeNumber.Text)
'    DEDUCTDATE = TRACKDATE(CDate(dtFromDate), CDate(dtToDate), MAXIMUMLIMIT)
'
'    'put the deduction hir
'    FORDEDUCT = DateDiff("D", DEDUCTDATE, dtToDate) + 1
'    For BILANG = 1 To FORDEDUCT
'        DEDUCTDATE = DEDUCTDATE + 1
'        'put deduction hir
'    Next

    Set rsGetType = New ADODB.Recordset
    Set RSGetRemLeave = New ADODB.Recordset
    rsGetType.Open ("Select reqcode,dte_from,dte_to from HRMS_REQUESTLEAVE_OT where empno = '" & Trim(cboEmployeeNumber.Text) & "' and reqdesc = '" & cboLeaveType.Text & "' and reqno = '" & labREQNO & "' and status = 'A'"), gconDMIS

    If Not (rsGetType.BOF And rsGetType.EOF) Then
        'Get date diffirence
        date_from = Trim(rsGetType!DTE_FROM)
        date_to = Trim(rsGetType!dte_to)
        
    End If
    
    GetUsed = (DateDiff("D", date_from, date_to) + 1) - SEARCHSUNDAY(dtFromDate, dtToDate) - SEARCHHOLIDAY(dtFromDate, dtToDate) - SEARCHSATURDAY(dtFromDate, dtToDate)

    yEmpIno = cboEmployeeNumber.Text
    yType = Trim(rsGetType!reqcode)
    yUsed = GetUsed
    
    
    Set checkrange = gconDMIS.Execute("select leave_code from hrms_leavemaster where leave_desc = '" & Trim(cboLeaveType.Text) & "'")
        If Not (checkrange.BOF And checkrange.EOF) Then
            xCODES = Trim(checkrange!LEAVE_CODE)

    End If
    
    RSGetRemLeave.Open ("Select days_no from hrms_leavemaster where leave_code = '" & Trim(yType) & "'"), gconDMIS
    If Not (RSGetRemLeave.BOF And RSGetRemLeave.EOF) Then

    yAve = Trim(RSGetRemLeave!DAYS_NO)
    maxleave = Trim(RSGetRemLeave!DAYS_NO)
    End If
    
    xAVE = GETAVTYPE(cboEmployeeNumber, GETTYPE(cboLeaveType))
    CHECK = GotValidate(xCODES, cboEmployeeNumber)
    xNAME = GETEMPNAME(cboEmployeeNumber)
'   datefroms = Trim(dtFromDate)
'   datetos = Trim(dtToDate)


   
        If CHECK = True Then
                
             MAXIMUMLIMIT = GETTOTALDAYS(GETTYPE(cboLeaveType), cboEmployeeNumber.Text)
             
             DEDUCTDATE = TRACKDATE(CDate(dtFromDate), CDate(dtToDate), MAXIMUMLIMIT)

            'put the deduction here
            FORDEDUCT = DateDiff("D", DEDUCTDATE, dtToDate) + 1
            For BILANG = 1 To FORDEDUCT
            
                        
              'DEDUCTDATE = DEDUCTDATE
              'put deduction here
            
            
            If cboCutOff.Text = "1st Cut-Off" Then
                         VCUT_OFF = 1
            End If
            If cboCutOff.Text = "2nd Cut-Off" Then
                         VCUT_OFF = 2
            End If


            Set rsSETUPDEDUCTION = gconDMIS.Execute("SELECT WORKING_DAY FROM HRMS_SETUPDEDUCTION")
                     If Not (rsSETUPDEDUCTION.EOF And rsSETUPDEDUCTION.BOF) Then
                         DAYS_OF_WORK = N2Str2Zero(rsSETUPDEDUCTION!WORKING_DAY)
                     End If
            Set rsgetlevel = gconDMIS.Execute("SELECT emplevel FROM HRMS_EMPINFO where empno = '" & Trim(cboEmployeeNumber.Text) & "'")
                              yemplevel = Trim(rsgetlevel!EMPLEVEL)
            rperday = Round(((GetBasicPay(Trim(cboEmployeeNumber.Text)) * 12) / DAYS_OF_WORK), 2)

            
            CXDATE = CInt(MONTH(DEDUCTDATE))
            CYDATE = CInt(Day(DEDUCTDATE))
            CZDATE = CInt(YEAR(DEDUCTDATE))

            Diyt = DateSerial(CZDATE, CXDATE, CYDATE)

            sqlInsert = "INSERT INTO HRMS_DEDUCTIONS (EMPLEVEL, EMPNO, DEYT, PARTICULAR, AMOUNT, NOMIN, CUT_OFF, PAY_MONTH, PAY_YEAR, MANUAL)"
            sqlInsert = sqlInsert & " Values('" & yemplevel & "'," & N2Str2Null(yEmpIno) & "," & N2Date2Null(Diyt) & ",'WD','" & rperday & "','0','" & VCUT_OFF & "','" & CXDATE & "','" & CZDATE & "','Y')"
            gconDMIS.Execute (sqlInsert)
            
            DEDUCTDATE = DEDUCTDATE + 1
            Next
        
        
        Else
                yAve = yAve - yUsed
                newUsed = 0: newAve = 0
                Set rsLeaveEntry = New ADODB.Recordset
                rsLeaveEntry.Open ("Select Emplno,[type],available,used from hrms_leave where emplno = '" & Trim(cboEmployeeNumber.Text) & "' and [type] = '" & yType & "' "), gconDMIS
                If Not (rsLeaveEntry.BOF And rsLeaveEntry.EOF) Then
                    newAve = Trim(rsLeaveEntry!Available)
                    newUsed = Trim(rsLeaveEntry!used)
        
                    newUsed = newUsed + yUsed
                    newAve = newAve - yUsed
        
                    sqlInsert = "Update hrms_leave set available = '" & newAve & "',used = '" & newUsed & "'"
                    sqlInsert = sqlInsert & " where emplno = '" & Trim(cboEmployeeNumber.Text) & "' and  [type] = '" & yType & "'"
                Else
                    sqlInsert = "Insert into hrms_leave (EmpLno,Type,Available,Used,Dateasof)"
                    sqlInsert = sqlInsert & " Values('" & yEmpIno & "','" & yType & "','" & yAve & "','" & yUsed & "','" & Trim(dtFiling) & "')"
        
                End If
                gconDMIS.Execute (sqlInsert)
        
              'Update HRMS leave
                Select Case yType
                    Case "VL"
                        sqltxt = "update hrms_leave set maxVL = '" & Trim(maxleave) & "'"
                        sqltxt = sqltxt & " where emplno = '" & yEmpIno & "' and [type] = 'VL'"
        
                    Case "SL"
                        sqltxt = "update hrms_leave set maxSL = '" & Trim(maxleave) & "'"
                        sqltxt = sqltxt & " where emplno = '" & yEmpIno & "' and [type] = 'SL'"
        
                    Case "PL"
                        sqltxt = "update hrms_leave set maxPL = '" & Trim(maxleave) & "'"
                        sqltxt = sqltxt & " where emplno = '" & yEmpIno & "' and [type] = 'PL'"
        
                    Case "ML"
                        sqltxt = "update hrms_leave set maxML = '" & Trim(maxleave) & "'"
                        sqltxt = sqltxt & " where emplno = '" & yEmpIno & "' and [type] = 'ML'"
        
                    Case "EL"
                        sqltxt = "update hrms_leave set maxEL = '" & Trim(maxleave) & "'"
                        sqltxt = sqltxt & " where emplno = '" & yEmpIno & "' and [type] = 'EL'"
        
                    Case Else
        
                End Select
                gconDMIS.Execute (sqltxt)
        
                Set rsLeaveEntry = Nothing
                Set RSGetRemLeave = Nothing
                Set rsGetType = Nothing
           
        End If



Errorcode:
'        MsgBox Err.Description
'        Exit Sub
End Sub
Function SEARCHSUNDAY(FromDate As Date, ToDate As Date) As Integer
    Dim count As Integer
    Dim CYDATE As Date
    Dim CXDATE As Date
    Dim SUN As Integer
    
    
    CXDATE = Trim(ToDate)
    CYDATE = Trim(FromDate)
     
    Do While CYDATE <> CXDATE
        If Weekday(CXDATE) = vbSunday Then
            SUN = SUN + 1
            
        End If
        CXDATE = CDate(CXDATE) - 1
    
    Loop
    
    SEARCHSUNDAY = SUN
    
End Function

Function SEARCHSATURDAY(FromDate As Date, ToDate As Date) As Integer
    Dim count As Integer
    Dim CYDATE As Date
    Dim CXDATE As Date
    Dim SAT As Integer
    
    
    CXDATE = Trim(ToDate)
    CYDATE = Trim(FromDate)
     
    Do While CYDATE <> CXDATE
        If Weekday(CXDATE) = vbSaturday Then
            SAT = SAT + 1
            
        End If
        CXDATE = CDate(CXDATE) - 1
    
    Loop
    
    SEARCHSATURDAY = SAT
    
End Function


'compute for how many holidays
'By: NVB
Function SEARCHHOLIDAY(FromDate As Date, ToDate As Date) As Integer
    Dim count As Integer
    Dim xFROM As Integer
    Dim xTO As Integer
    Dim CXDATE, CYDATE, CZDATE As Integer
    Dim HOLIDAY As Integer
    Dim sqltxt As String
    Dim RSHOLIDAY As New ADODB.Recordset
    Dim XMANTH, XDEYT, XYEAR As Integer
    Dim XCOMBINE As Date
    Dim YCOMBINE As Date
    Dim ZFROMDATE As Date
    
    'xTO = CInt(Day(ToDate))
    'xFROM = CInt(Day(FromDate))
    ZFROMDATE = FromDate
    
    CXDATE = CInt(MONTH(ToDate))
    CYDATE = CInt(Day(ToDate))
    CZDATE = CInt(YEAR(ToDate))
    
    YCOMBINE = CDate((CStr(CXDATE & "/" & CYDATE & "/" & CZDATE)))

    sqltxt = "SELECT MANTH,DEYT FROM HRMS_HOLIDAY_LIST"
    Set RSHOLIDAY = gconDMIS.Execute(sqltxt)
    If Not (RSHOLIDAY.EOF And RSHOLIDAY.BOF) Then
    
    End If
        
    RSHOLIDAY.MoveFirst
    Do While Not RSHOLIDAY.EOF
    
        XMANTH = CInt(Trim(RSHOLIDAY!MANTH))
        XDEYT = CInt(Trim(RSHOLIDAY!DEYT))
        XYEAR = CInt(YEAR(Trim(ToDate)))

        XCOMBINE = CDate((CStr(XMANTH & "/" & XDEYT & "/" & XYEAR)))

        'For count = xFROM To xTO
    Do While ZFROMDATE <> YCOMBINE
     
        If XCOMBINE = YCOMBINE Then
            HOLIDAY = HOLIDAY + 1
        End If
        
        YCOMBINE = YCOMBINE - 1
        'CYDATE = CYDATE - 1
    Loop
    'Next count
    RSHOLIDAY.MoveNext
    Loop
    
    SEARCHHOLIDAY = HOLIDAY
    Set RSHOLIDAY = Nothing
End Function

Function GetBasicPay(XXX As String) As Double

    
    Dim RSTMP                                                         As ADODB.Recordset
    
    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT BASICSALARY, EMPSTATUS,ALLOWANCE FROM HRMS_EMPINFO WHERE EMPNO = " & N2Str2Null(XXX) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!EMPSTATUS) = "M" Then
            'GetBasicPay = N2Str2Zero(RSTMP!BASICSALARY)
            GetBasicPay = ((N2Str2Zero(RSTMP!BASICSALARY) / 2) + N2Str2Zero(RSTMP!ALLOWANCE) / 2)
        ElseIf Null2String(RSTMP!EMPSTATUS) = "D" Then
            GetBasicPay = (N2Str2Zero(RSTMP!BASICSALARY) * DAYS_OF_WORK) / 12
        End If
    End If
    Set RSTMP = Nothing
End Function

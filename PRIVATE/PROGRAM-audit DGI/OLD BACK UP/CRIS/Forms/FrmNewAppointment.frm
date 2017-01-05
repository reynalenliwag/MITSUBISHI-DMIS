VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCSMSNewAppointment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Entry"
   ClientHeight    =   8505
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   10440
   ForeColor       =   &H00D8E9EC&
   Icon            =   "FrmNewAppointment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   10440
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picReason 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   2250
      ScaleHeight     =   6165
      ScaleWidth      =   8025
      TabIndex        =   66
      Top             =   930
      Visible         =   0   'False
      Width           =   8025
      Begin VB.CommandButton cmdOther 
         Caption         =   "Add &Other Jobs"
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
         Left            =   1590
         TabIndex        =   12
         Top             =   5640
         Width           =   1425
      End
      Begin VB.TextBox txtRecomendation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   98
         Top             =   4140
         Width           =   7755
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Height          =   585
         Left            =   150
         TabIndex        =   88
         Top             =   4980
         Width           =   7755
         Begin VB.TextBox txtRecorded 
            BackColor       =   &H00D8E9EC&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1560
            TabIndex        =   89
            Top             =   150
            Width           =   1305
         End
         Begin MSComCtl2.DTPicker dtPromised 
            Height          =   345
            Left            =   5220
            TabIndex        =   10
            Top             =   150
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   609
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
            CustomFormat    =   "MM/dd/yyyy hh:mm:ss tt"
            Format          =   52232195
            CurrentDate     =   38936
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Promised Date/Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3330
            TabIndex        =   91
            Top             =   210
            Width           =   1755
         End
         Begin VB.Label Label17 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Recorded"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   210
            TabIndex        =   90
            Top             =   210
            Width           =   1605
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add &Canned Labor"
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
         Left            =   4320
         TabIndex        =   14
         Top             =   5640
         Width           =   1515
      End
      Begin VB.CommandButton cmdPMS 
         Caption         =   "Add &PMS Jobs"
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
         Left            =   3030
         TabIndex        =   13
         Top             =   5640
         Width           =   1275
      End
      Begin MSComctlLib.ListView lblJob4Service 
         Height          =   1515
         Left            =   120
         TabIndex        =   70
         Top             =   450
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   2672
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmNewAppointment.frx":08CA
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Job Type"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "  Jobs Description"
            Object.Width           =   8467
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Flat Rate"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Std.Time"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Discount"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Charged To"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Note"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Job"
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
         Left            =   5850
         TabIndex        =   15
         Top             =   5640
         Width           =   1245
      End
      Begin VB.CommandButton cmdAddJobs 
         Caption         =   "Add &General Job"
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
         Left            =   150
         TabIndex        =   11
         Top             =   5640
         Width           =   1425
      End
      Begin MSComctlLib.ListView lstPMSDet 
         Height          =   1305
         Left            =   120
         TabIndex        =   73
         Top             =   2400
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   2302
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmNewAppointment.frx":0A2C
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Job Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   11289
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Model"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Jobs for Service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   5
         Left            =   150
         TabIndex        =   67
         Top             =   60
         Width           =   2835
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Total STD Time :"
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
         Left            =   5760
         TabIndex        =   102
         Top             =   2100
         Width           =   1305
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Flat Rate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3330
         TabIndex        =   101
         Top             =   2100
         Width           =   1215
      End
      Begin VB.Label labNotes 
         BackStyle       =   0  'Transparent
         Caption         =   "&Diagnosis entry/Recommendation for future servicing :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   210
         TabIndex        =   100
         Top             =   3810
         Width           =   5985
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&PMS/Canned Job Details :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   8
         Left            =   150
         TabIndex        =   74
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label lbltlFaltRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4590
         TabIndex        =   69
         Top             =   2070
         Width           =   1095
      End
      Begin VB.Label lblStdHrs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7020
         TabIndex        =   68
         Top             =   2040
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdCancel 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9450
      MouseIcon       =   "FrmNewAppointment.frx":0B8E
      MousePointer    =   99  'Custom
      Picture         =   "FrmNewAppointment.frx":0CE0
      Style           =   1  'Graphical
      TabIndex        =   131
      ToolTipText     =   "Cancel Transaction"
      Top             =   7560
      Width           =   795
   End
   Begin VB.TextBox txtAppointmentEdit 
      Height          =   285
      Left            =   2190
      TabIndex        =   130
      Text            =   "txtAppointmentEdit"
      Top             =   7950
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtEstimateEdit 
      Height          =   285
      Left            =   2190
      TabIndex        =   129
      Text            =   "txtEstimateEdit"
      Top             =   7620
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtTranNo 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2250
      Width           =   1755
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   1380
      Top             =   6960
   End
   Begin VB.CommandButton cmdNext 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8730
      Picture         =   "FrmNewAppointment.frx":101E
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Next"
      Top             =   7560
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   2430
      ScaleHeight     =   705
      ScaleWidth      =   7725
      TabIndex        =   122
      Top             =   150
      Width           =   7725
      Begin VB.TextBox txtCustName 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   1710
         TabIndex        =   124
         Top             =   90
         Width           =   6015
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   123
         Top             =   90
         Width           =   825
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   30
         TabIndex        =   125
         Top             =   180
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdBack 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7950
      Picture         =   "FrmNewAppointment.frx":136C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Previous "
      Top             =   7560
      Width           =   795
   End
   Begin VB.PictureBox picEstimate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C8CBFD&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   2250
      ScaleHeight     =   6165
      ScaleWidth      =   8025
      TabIndex        =   103
      Top             =   930
      Visible         =   0   'False
      Width           =   8025
      Begin VB.CommandButton Command2 
         Caption         =   "&Add Parts && Accessories"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   5610
         Width           =   2475
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C8CBFD&
         Caption         =   "Estimated Cost "
         Height          =   2145
         Left            =   210
         TabIndex        =   105
         Top             =   3240
         Width           =   7665
         Begin VB.TextBox txtRateAces 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Enabled         =   0   'False
            Height          =   330
            Left            =   5130
            TabIndex        =   37
            Top             =   1290
            Width           =   675
         End
         Begin VB.TextBox txtRateparts 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Enabled         =   0   'False
            Height          =   330
            Left            =   5130
            TabIndex        =   36
            Top             =   900
            Width           =   675
         End
         Begin VB.TextBox txtRateLabor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Enabled         =   0   'False
            Height          =   330
            Left            =   5130
            TabIndex        =   35
            Top             =   510
            Width           =   675
         End
         Begin VB.TextBox txtEstLabor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
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
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   510
            Width           =   975
         End
         Begin VB.TextBox txtEstParts 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
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
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   900
            Width           =   975
         End
         Begin VB.TextBox txtEstAces 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
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
            Left            =   1230
            TabIndex        =   21
            Top             =   1290
            Width           =   975
         End
         Begin VB.TextBox txtTotalAmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
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
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtCompLabor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   2250
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   510
            Width           =   975
         End
         Begin VB.TextBox txtCompPart 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   2250
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   900
            Width           =   975
         End
         Begin VB.TextBox txtCompAces 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   2250
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   1290
            Width           =   975
         End
         Begin VB.TextBox txtCompTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   2250
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtSalesLabor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   510
            Width           =   795
         End
         Begin VB.TextBox txtSalesParts 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   900
            Width           =   795
         End
         Begin VB.TextBox txtSalesAces 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   1290
            Width           =   795
         End
         Begin VB.TextBox txtSalesTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   1680
            Width           =   825
         End
         Begin VB.TextBox txtWarLabor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   4110
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   510
            Width           =   975
         End
         Begin VB.TextBox txtWarParts 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   4110
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   900
            Width           =   975
         End
         Begin VB.TextBox txtWarLaborAces 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   4110
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1290
            Width           =   975
         End
         Begin VB.TextBox txtWarLaborTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   4110
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtDiscLabor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Enabled         =   0   'False
            Height          =   330
            Left            =   5850
            TabIndex        =   38
            Top             =   510
            Width           =   855
         End
         Begin VB.TextBox txtDiscParts 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Enabled         =   0   'False
            Height          =   330
            Left            =   5850
            TabIndex        =   39
            Top             =   900
            Width           =   855
         End
         Begin VB.TextBox txtDiscAces 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Enabled         =   0   'False
            Height          =   330
            Left            =   5850
            TabIndex        =   40
            Top             =   1290
            Width           =   855
         End
         Begin VB.TextBox txtDiscTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Enabled         =   0   'False
            Height          =   330
            Left            =   5850
            TabIndex        =   41
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtVatLabor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   6750
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   510
            Width           =   825
         End
         Begin VB.TextBox txtVatParts 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   6750
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   900
            Width           =   825
         End
         Begin VB.TextBox txtVatAces 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   6750
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   1290
            Width           =   825
         End
         Begin VB.TextBox txtVatTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C8CBFD&
            Height          =   330
            Left            =   6750
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   1680
            Width           =   825
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Disc.Rate"
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
            Left            =   5130
            TabIndex        =   127
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Labor"
            Height          =   285
            Left            =   690
            TabIndex        =   115
            Top             =   570
            Width           =   555
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Parts"
            Height          =   285
            Left            =   720
            TabIndex        =   114
            Top             =   960
            Width           =   555
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Accessories"
            Height          =   285
            Left            =   210
            TabIndex        =   113
            Top             =   1350
            Width           =   1095
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
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
            Left            =   600
            TabIndex        =   112
            Top             =   1740
            Width           =   615
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
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
            Left            =   1500
            TabIndex        =   111
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Company"
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
            Left            =   2400
            TabIndex        =   110
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales"
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
            Left            =   3570
            TabIndex        =   109
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Warranty"
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
            Left            =   4200
            TabIndex        =   108
            Top             =   270
            Width           =   945
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Discount"
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
            Left            =   6000
            TabIndex        =   107
            Top             =   270
            Width           =   825
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "VAT"
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
            Left            =   6990
            TabIndex        =   106
            Top             =   270
            Width           =   435
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2505
         Left            =   180
         TabIndex        =   116
         Top             =   630
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   4419
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
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
         MouseIcon       =   "FrmNewAppointment.frx":16BA
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Parts No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Parts Description"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "SRP"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Parts and Accessories Estimate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   210
         TabIndex        =   104
         Top             =   150
         Width           =   3615
      End
   End
   Begin VB.PictureBox picAppointment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   2250
      ScaleHeight     =   6165
      ScaleWidth      =   8025
      TabIndex        =   52
      Top             =   930
      Visible         =   0   'False
      Width           =   8025
      Begin VB.TextBox txtApointmentNo 
         BackColor       =   &H00D8E9EC&
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
         Height          =   375
         Left            =   6330
         MaxLength       =   10
         TabIndex        =   96
         Text            =   "Text1"
         Top             =   840
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Vehicle information"
         Height          =   1575
         Left            =   180
         TabIndex        =   76
         Top             =   2490
         Width           =   7755
         Begin VB.TextBox txtKm_rdg 
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
            Left            =   6810
            MaxLength       =   9
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1020
            Width           =   795
         End
         Begin VB.TextBox txtVIN 
            BackColor       =   &H00D8E9EC&
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
            Left            =   3660
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   85
            Text            =   "Text1"
            Top             =   1020
            Width           =   1845
         End
         Begin VB.TextBox txtYear 
            BackColor       =   &H00D8E9EC&
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
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   83
            Text            =   "Text1"
            Top             =   450
            Width           =   795
         End
         Begin VB.TextBox txtPlate_No 
            BackColor       =   &H00D8E9EC&
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
            Left            =   930
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   79
            Text            =   "Text1"
            Top             =   1020
            Width           =   1785
         End
         Begin VB.TextBox txtMake 
            BackColor       =   &H00D8E9EC&
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
            Left            =   2190
            Locked          =   -1  'True
            TabIndex        =   78
            Text            =   "Text1"
            Top             =   450
            Width           =   2055
         End
         Begin VB.TextBox cboModel 
            BackColor       =   &H00D8E9EC&
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
            Left            =   4980
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   77
            Text            =   "Text1"
            Top             =   450
            Width           =   2625
         End
         Begin VB.Label Label14 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "KM Reading"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5700
            TabIndex        =   87
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label27 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "VIN No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2940
            TabIndex        =   86
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label Label26 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   240
            TabIndex        =   84
            Top             =   510
            Width           =   525
         End
         Begin VB.Label Label10 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Plate No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   90
            TabIndex        =   82
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label Label25 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Make"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1680
            TabIndex        =   81
            Top             =   510
            Width           =   585
         End
         Begin VB.Label Label24 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4410
            TabIndex        =   80
            Top             =   510
            Width           =   585
         End
      End
      Begin VB.TextBox txtParticipat 
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
         Height          =   495
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   5400
         Width           =   3855
      End
      Begin VB.ComboBox cboRecd_by 
         Appearance      =   0  'Flat
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
         Left            =   1650
         Sorted          =   -1  'True
         TabIndex        =   7
         Text            =   "cboRecd_by"
         Top             =   4920
         Width           =   3855
      End
      Begin VB.TextBox txtRep_Or 
         BackColor       =   &H00D8E9EC&
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
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox txtNiym 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   1440
         Width           =   5595
      End
      Begin VB.TextBox txtEstimateno 
         BackColor       =   &H00D8E9EC&
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
         Height          =   375
         Left            =   3750
         MaxLength       =   10
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtAcct_No 
         BackColor       =   &H00D8E9EC&
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
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   1440
         Width           =   1065
      End
      Begin VB.TextBox txtSektion 
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
         Left            =   6840
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   4890
         Width           =   735
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00D8E9EC&
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
         Height          =   405
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   1890
         Width           =   6705
      End
      Begin MSComCtl2.DTPicker txtDte_recd 
         Height          =   375
         Left            =   1650
         TabIndex        =   6
         Top             =   4380
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   661
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
         CustomFormat    =   "MMM. dd, yyyy"
         Format          =   52232195
         CurrentDate     =   38936
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   5190
         TabIndex        =   97
         Top             =   930
         Width           =   1485
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Recorded"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   75
         Top             =   4440
         Width           =   1605
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance             Participation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         TabIndex        =   65
         Top             =   5460
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   270
         TabIndex        =   64
         Top             =   1500
         Width           =   1635
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RO Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   63
         Top             =   930
         Width           =   1035
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   420
         TabIndex        =   62
         Top             =   1950
         Width           =   1035
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Estimate No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2580
         TabIndex        =   61
         Top             =   900
         Width           =   1485
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Section No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5730
         TabIndex        =   60
         Top             =   4950
         Width           =   1305
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Advisor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   330
         TabIndex        =   59
         Top             =   4980
         Width           =   1515
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Repair Order Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   4
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   4155
      End
   End
   Begin VB.PictureBox picVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   2250
      ScaleHeight     =   6165
      ScaleWidth      =   8025
      TabIndex        =   50
      Top             =   930
      Visible         =   0   'False
      Width           =   8025
      Begin VB.TextBox txtVehName 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   630
         Width           =   6555
      End
      Begin VB.CommandButton cmdAddVeh 
         Caption         =   "&Add/Edit/Delete Vehicle"
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
         Left            =   210
         TabIndex        =   4
         Top             =   5670
         Width           =   2355
      End
      Begin MSComctlLib.ListView lstVehicle 
         Height          =   4485
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7911
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
         MouseIcon       =   "FrmNewAppointment.frx":181C
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Model"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Plate No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Serial No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Engine"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Prod'n. No."
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   570
         TabIndex        =   71
         Top             =   630
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Select Vehicle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   51
         Top             =   90
         Width           =   1605
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D8E9EC&
      Height          =   6165
      Left            =   2250
      TabIndex        =   117
      Top             =   930
      Width           =   8025
      Begin VB.OptionButton optEndUser 
         BackColor       =   &H00D8E9EC&
         Caption         =   "End User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   134
         Top             =   660
         Width           =   1305
      End
      Begin VB.CommandButton cmdAddeditCustomer 
         Caption         =   "Add/Edit/Delete Customer"
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
         Left            =   210
         TabIndex        =   133
         ToolTipText     =   "Add/Edit/Delete Customer"
         Top             =   5670
         Width           =   2475
      End
      Begin VB.OptionButton optFullName 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Full Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4890
         TabIndex        =   120
         Top             =   660
         Width           =   1305
      End
      Begin VB.OptionButton optLN 
         BackColor       =   &H00D8E9EC&
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
         Height          =   285
         Left            =   2070
         TabIndex        =   119
         Top             =   660
         Width           =   1305
      End
      Begin VB.OptionButton optFN 
         BackColor       =   &H00D8E9EC&
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
         Height          =   285
         Left            =   3480
         TabIndex        =   118
         Top             =   660
         Width           =   1305
      End
      Begin VB.TextBox textSearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2040
         TabIndex        =   1
         Top             =   180
         Width           =   5835
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   4575
         Left            =   180
         TabIndex        =   2
         Top             =   1020
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   8070
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
         MouseIcon       =   "FrmNewAppointment.frx":197E
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Last Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "First Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Address"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Province"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Phone No."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "CusName"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - SEARCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   132
         Top             =   690
         Width           =   1110
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Select Customer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   121
         Top             =   210
         Width           =   1725
      End
   End
   Begin VB.Label labEdit 
      Height          =   405
      Left            =   390
      TabIndex        =   128
      Top             =   -390
      Width           =   1215
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   765
      Left            =   2220
      TabIndex        =   126
      Top             =   120
      Width           =   8055
      BackColor       =   14215660
      Size            =   "14208;1349"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label labTranType 
      BackStyle       =   0  'Transparent
      Caption         =   "Repair Order No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   99
      Top             =   1920
      Width           =   1635
   End
   Begin VB.Label labType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Repair Order"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   94
      Top             =   7590
      Width           =   1755
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Appointment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   93
      Top             =   4710
      Width           =   900
   End
   Begin VB.Shape ShpAppointment 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   405
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   4620
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   92
      Top             =   5730
      Width           =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   300
      X2              =   10470
      Y1              =   7470
      Y2              =   7470
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   210
      Picture         =   "FrmNewAppointment.frx":1AE0
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1725
   End
   Begin VB.Shape shpCustomer 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   405
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   3090
      Width           =   1875
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jobs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   49
      Top             =   5220
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Repair Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   48
      Top             =   4200
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   47
      Top             =   3690
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   210
      TabIndex        =   46
      Top             =   3180
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   2100
      X2              =   2100
      Y1              =   120
      Y2              =   8310
   End
   Begin VB.Label labType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Repair Order"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   270
      TabIndex        =   95
      Top             =   7590
      Width           =   1755
   End
   Begin VB.Shape shpVehicle 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   405
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Shape shpRO 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   405
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   4110
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Shape shpJobs 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   405
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   5130
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Shape ShpEstimate 
      BackColor       =   &H00C8CBFD&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   405
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Visible         =   0   'False
      Width           =   1875
   End
End
Attribute VB_Name = "frmCSMSNewAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAddRepor                          As ADODB.Recordset
Dim rsAddRepor2                         As ADODB.Recordset
Dim rsFind                              As ADODB.Recordset
Dim ctl                                 As Control
Dim bevvy                               As Long
Dim tlHrs                               As Double
Dim tlFR                                As Double
Dim xPartsAmt                           As Double
Dim xAcesAmt                            As Double
Dim xTransType                          As String
Dim xApptNo                             As String
Dim xESTIMATENO                         As String

Dim JobTotal                            As Double
Dim JobComTotal                         As Double
Dim JobSalesTotal                       As Double
Dim JobWarTotal                         As Double
Dim JobDiscTotal                        As Double
Dim JobVatTotal                         As Double

Dim PartsTotal                          As Double
Dim PartsComTotal                       As Double
Dim PartsSalesTotal                     As Double
Dim PartsWarTotal                       As Double
Dim PartsDiscTotal                      As Double
Dim PartsVatTotal                       As Double

Dim MatTotal                            As Double
Dim MatComTotal                         As Double
Dim MatSalesTotal                       As Double
Dim MatWarTotal                         As Double
Dim MatDiscTotal                        As Double
Dim MatVatTotal                         As Double
Dim COMTotal                            As Double
Dim SALESTotal                          As Double
Dim WARTotal                            As Double
Dim VATTotal                            As Double
Dim ROTotal                             As Double

Dim EndUserCode                         As String

Sub EditTransaction()

    On Error GoTo Errorcode

    optFullName.Value = True
    Set rsFind = New ADODB.Recordset
    If labType(0).Caption = "Repair Order" Then
        labTranType.Caption = "Repair Order  No."
        xTransType = "R"
        Set rsFind = gconDMIS.Execute("select * from CSMS_vw_REPAIRORDER where RO_NO = '" & txtTranNo & "'")
    ElseIf labType(0).Caption = "Estimate" Then
        labTranType.Caption = "Estimate No."
        xTransType = "E"
        Set rsFind = gconDMIS.Execute("select * from CSMS_vw_REPAIRORDER where ESTIMATENO = '" & txtTranNo & "'")
    ElseIf labType(0).Caption = "Appointment" Then
        labTranType.Caption = "Appointment No."
        xTransType = "A"
        Set rsFind = gconDMIS.Execute("select * from CSMS_vw_REPAIRORDER where ApptNo = '" & txtTranNo & "'")
    End If
    If Not rsFind.EOF And Not rsFind.BOF Then
        If labType(0).Caption = "Repair Order" Then
            txtRep_Or = txtTranNo
        ElseIf labType(0).Caption = "Estimate" Then
            txtEstimateno = txtTranNo
        ElseIf labType(0).Caption = "Appointment" Then
            txtApointmentNo = txtTranNo
        End If
        txtEstimateEdit = Null2String(rsFind![EstimateNo])
        txtAppointmentEdit = Null2String(rsFind![ApptNo])
        textSearch = Null2String(rsFind![Customer])
        txtCustName = Null2String(UCase(rsFind![Customer]))
        txtNiym = Null2String(rsFind![Customer])
        txtID = Null2String(rsFind![ACCT_NO])
        txtAddress = Null2String(rsFind![CustomerAdd])
        txtAcct_No = Null2String(rsFind![ACCT_NO])
        txtDte_recd.Value = Null2String(rsFind![AppointmentDate])
        dtPromised.Value = Null2String(rsFind![promisedate])

        Dim rsVehicleKo                 As ADODB.Recordset
        Set rsVehicleKo = New ADODB.Recordset
        Set rsVehicleKo = gconDMIS.Execute("select * from CSMS_Cusveh where Cuscde = '" & txtID & "' and plate_no = '" & rsFind![Plate_no] & "'")
        If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
            txtPlate_No = Null2String(rsVehicleKo![Plate_no])
            cboModel = Null2String(rsVehicleKo![Model])
            txtMake = Null2String(rsVehicleKo![Make])
            txtyear = Null2String(rsVehicleKo![Yer])
            txtVIN = Null2String(rsVehicleKo![Vin])
            txtVehName = Trim(cboModel) & "   " & txtPlate_No
        End If

        Set rsVehicleKo = New ADODB.Recordset
        Set rsVehicleKo = gconDMIS.Execute("select model,Plate_no,serial,engine,prodno from CSMS_CusVeh where Cuscde = '" & txtID & "'")
        If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
            Listview_Loadval Me.lstVehicle.ListItems, rsVehicleKo
            lstCustomer.Refresh
        End If
        'JOBS
        Set rsFind = New ADODB.Recordset
        If labType(0).Caption = "Repair Order" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_vw_EditRO where rep_OR = '" & txtTranNo & "' order by jobtype,detdsc asc")
        ElseIf labType(0).Caption = "Estimate" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_vw_EditEstimate where ESTIMATENO = '" & txtTranNo & "' order by jobtype,detdsc asc")
        ElseIf labType(0).Caption = "Appointment" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_vw_EditAppt where ApptNo = '" & txtTranNo & "' order by jobtype,detdsc asc")
        End If
        If Not rsFind.EOF And Not rsFind.BOF Then
            txtKm_rdg.Text = Null2String(rsFind![KM_RDG])
            cboRecd_by.Text = SetSAname(Null2String(rsFind![RECD_BY]))
            txtSektion.Text = Null2String(rsFind![sektion])
            txtRecorded = Null2String(rsFind![dte_recd])
            txtDte_recd.Value = Null2String(rsFind![dte_recd])
            txtParticipat.Text = Null2String(rsFind![participat])
            txtRecomendation = Null2String(rsFind![NOTE])
            Do Until rsFind.EOF
                With lblJob4Service
                    .Sorted = False
                    .ListItems.Add , , Null2String(rsFind![DETCDE])
                    .ListItems(.ListItems.Count).ListSubItems.Add 1, , Null2String(rsFind![JOBTYPE])
                    .ListItems(.ListItems.Count).ListSubItems.Add 2, , Null2String(rsFind![Detdsc])
                    .ListItems(.ListItems.Count).ListSubItems.Add 3, , NumericVal(rsFind![FlatRate])
                    .ListItems(.ListItems.Count).ListSubItems.Add 4, , NumericVal(rsFind![DET_HRS])
                    .ListItems(.ListItems.Count).ListSubItems.Add 5, , NumericVal(rsFind![discrate])
                    .ListItems(.ListItems.Count).ListSubItems.Add 6, , Null2String(rsFind![wCode])
                    .ListItems(.ListItems.Count).ListSubItems.Add 7, , Null2String(rsFind![Detail])
                End With
                rsFind.MoveNext
            Loop
        End If
        'PMS DETAILS
        Set rsFind = New ADODB.Recordset
        If labType(0).Caption = "Repair Order" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_PMS_Job_Det where rep_OR = '" & txtTranNo & "' order by pms_model,detcde asc")
        ElseIf labType(0).Caption = "Estimate" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_PMS_Job_Det where ESTIMATENO = '" & txtTranNo & "' order by pms_model,detcde asc")
        ElseIf labType(0).Caption = "Appointment" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_PMS_Job_Det where ApptNo = '" & txtTranNo & "' order by pms_model,detcde asc")
        End If
        If Not rsFind.EOF And Not rsFind.BOF Then
            Do Until rsFind.EOF
                With lstPMSDet
                    .Sorted = False
                    .ListItems.Add , , Null2String(rsFind![DETCDE])
                    .ListItems(.ListItems.Count).ListSubItems.Add 1, , Null2String(rsFind![JOBTYPE])
                    .ListItems(.ListItems.Count).ListSubItems.Add 2, , Null2String(rsFind![Detdsc])
                    .ListItems(.ListItems.Count).ListSubItems.Add 3, , Null2String(rsFind![PMS_Model])
                End With
                rsFind.MoveNext
            Loop
        End If
        'ESTIMATE
        Set rsFind = New ADODB.Recordset
        Set rsFind = gconDMIS.Execute("select * from CSMS_vw_EstimateDetails where ESTIMATENO = '" & txtTranNo & "'")
        If Not rsFind.EOF And Not rsFind.BOF Then
            Do Until rsFind.EOF
                With ListView1
                    .Sorted = False
                    .ListItems.Add , , Null2String(rsFind![Type])
                    .ListItems(.ListItems.Count).ListSubItems.Add 1, , Null2String(rsFind![partno])
                    .ListItems(.ListItems.Count).ListSubItems.Add 2, , Null2String(rsFind![PartDesc])
                    .ListItems(.ListItems.Count).ListSubItems.Add 3, , NumericVal(rsFind![QTY])
                    .ListItems(.ListItems.Count).ListSubItems.Add 4, , NumericVal(rsFind![SRP])
                End With
                rsFind.MoveNext
            Loop
        End If
    End If

    Exit Sub

Errorcode:

    ShowVBError

    Exit Sub

End Sub

Private Sub cmdOther_Click()
    frmMain.MousePointer = 11

    With frmCSMSOtherJobs
        .txtCustomer = txtNiym
        .txtActNo = txtAcct_No
        .txtROno = txtRep_Or
        .txtAppt = "NewAppt"
        .txtCheckMe = ""
        .txtCheckMe = "ro"
        .txtVehicle = cboModel.Text
    End With
    frmCSMSOtherJobs.Show 1

    frmMain.MousePointer = 0
End Sub

Private Sub Command2_Click()
    frmMain.MousePointer = 11

    '    With frmCSMSAddEstimate
    '        .lblFrom.Caption = "Estimate"
    '        .txtEstNo.Text = txtEstimateno.Text
    '        .txtCustomer.Text = txtCustName.Text
    '        .txtVehicle.Text = Trim(txtyear.Text) & " " & Trim(txtMake.Text) & " " & Trim(cboModel.Text)
    '        .txtPlateno.Text = txtPlate_No.Text
    '    End With
    '    frmCSMSAddEstimate.Show 1

    frmMain.MousePointer = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If shpCustomer.Visible = True Then
        If KeyCode = vbKeyF3 Then textSearch.SetFocus
    End If
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        If MsgBox("DELETE! this parts...  " & ListView1.SelectedItem.SubItems(2) & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
            Exit Sub
        End If
        Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
        Call ComputeMe
    End If
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Stop
    If optEndUser.Value = True Then
        txtCustName = UCase(lstCustomer.SelectedItem.SubItems(5))
        txtNiym = lstCustomer.SelectedItem.SubItems(5)
        txtID = lstCustomer.SelectedItem.SubItems(6)
        txtAddress = lstCustomer.SelectedItem.SubItems(2)
        txtAcct_No = lstCustomer.SelectedItem.SubItems(6)
        GetVehicleforEndUser lstCustomer.SelectedItem.SubItems(2)
    Else
        txtCustName = UCase(lstCustomer.SelectedItem.SubItems(5))
        txtNiym = lstCustomer.SelectedItem.SubItems(5)
        txtID = lstCustomer.SelectedItem.SubItems(6)
        txtAddress = lstCustomer.SelectedItem.SubItems(2)
        txtAcct_No = lstCustomer.SelectedItem.SubItems(6)
        Call GetVehicleforCustomer
    End If
End Sub
 
Private Sub cmdAddCust_Click()
    frmAllCustomer.Show 1
End Sub
Private Sub cmdAddeditCustomer_Click()
    frmAllCustomer.Show 1
End Sub

Private Sub cmdAddJobs_Click()
    frmMain.MousePointer = 11

    With frmCSMSReqJobs
        .txtCustomer.Text = txtNiym.Text
        .txtActNo.Text = txtAcct_No.Enabled
        .txtROno.Text = txtRep_Or.Text
        .txtAppt.Text = "NewAppt"
        .txtCheckMe.Text = ""
        .txtCheckMe.Text = "ro"
    End With
    frmCSMSReqJobs.Show 1

    frmMain.MousePointer = 0
End Sub

Private Sub cmdAddVeh_Click()
    If optEndUser.Value = True Then
        'With frmCSMSAddVehicleEndUser
        '    .CustomerCode = txtPlate_No.Text
        '    .labCustCode.Caption = txtAcct_No.Text
        '    .labCustomer.Caption = txtCustName
        'End With
        'frmCSMSAddVehicleEndUser.Show 1
        'Call GetVehicleforCustomer
    Else
        With frmCSMSAddVehicle
            .CustomerCode = txtID
            .labCustCode.Caption = txtAcct_No.Text
            .labCustomer = txtCustName
        End With
        frmCSMSAddVehicle.Show 1
        Call GetVehicleforCustomer
    End If
End Sub

Private Sub cmdBack_Click()
    If shpVehicle.Visible = True Then
        shpVehicle.Visible = False
        shpCustomer.Visible = True
        If lstCustomer.Enabled = True And lstCustomer.ListItems.Count > 0 Then
            lstCustomer.SetFocus
        End If
        cmdAddeditCustomer.Visible = True
        picVehicle.Visible = False
    ElseIf shpRO.Visible = True Then
        picVehicle.Visible = True
        picAppointment.Visible = False
        shpRO.Visible = False
        shpVehicle.Visible = True
        If lstVehicle.ListItems.Count > 0 And lstVehicle.Enabled = True Then
            lstVehicle.SetFocus
        End If
    ElseIf ShpAppointment.Visible = True Then                '
        ShpAppointment.Visible = False
        picVehicle.Visible = True
        picAppointment.Visible = False
        shpRO.Visible = False
        shpVehicle.Visible = True
        If lstVehicle.Enabled = True And lstVehicle.ListItems.Count > 0 Then
            lstVehicle.SetFocus
        End If
    ElseIf shpJobs.Visible = True Then
        If labType(0).Caption = "Appointment" Then
            ShpAppointment.Visible = True
            picReason.Visible = False
            shpJobs.Visible = False
            picAppointment.Visible = True
            cmdNext.Caption = "&Next  >>"
            Label6(4).Caption = "Appointment Information"
        Else
            picReason.Visible = False
            shpJobs.Visible = False
            picAppointment.Visible = True
            shpRO.Visible = True
            cmdNext.Caption = "&Next  >>"
            Label6(4).Caption = "Repair Order Information"
        End If
    ElseIf ShpEstimate.Visible = True Then
        shpJobs.Visible = True
        ShpEstimate.Visible = False
        picEstimate.Visible = False
        cmdNext.Caption = "&Next  >>"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    If MsgBox("DELETE! this job...  " & lblJob4Service.SelectedItem.SubItems(1) & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
        Exit Sub
    End If


    frmMain.MousePointer = 11
    Call RemovePMSJobDetails(lblJob4Service.SelectedItem.Text)


    Me.lblJob4Service.ListItems.Remove Me.lblJob4Service.SelectedItem.Index


    Call ComputeResultOfRatenTimeWhenJobDelete

    frmMain.MousePointer = 0

End Sub

Sub RemovePMSJobDetails(code As String)
    Dim x                               As Integer

    For x = lstPMSDet.ListItems.Count To 1 Step -1
        If code = lstPMSDet.ListItems(x).SubItems(3) Then
            lstPMSDet.ListItems.Remove (x)
        End If
    Next
End Sub

Function ComputeResultOfRatenTimeWhenJobDelete()
    Dim x                               As Integer
    Dim FR                              As Double
    Dim ST                              As Double

    For x = 1 To lblJob4Service.ListItems.Count
        FR = FR + lblJob4Service.ListItems(x).SubItems(3)
        ST = ST + lblJob4Service.ListItems(x).SubItems(4)
    Next

    lbltlFaltRate.Caption = FR
    lblStdHrs.Caption = ST
End Function

Private Sub cmdEdit_Click()
    With frmCSMSJobSelected
        For Each ctl In .ControlS
            If TypeOf ctl Is TextBox Then
                ctl.Text = ""
            End If
        Next ctl
        .cboJobChargeTo.Clear
        .cboJobChargeTo.AddItem "W"
        .cboJobChargeTo.AddItem "S"
        .cboJobChargeTo.AddItem "C"
        Dim rsUpload                    As ADODB.Recordset
        Set rsUpload = New ADODB.Recordset
        Set rsUpload = gconDMIS.Execute("Select * from CSMS_Ro_Det where REP_OR = '" & txtRep_Or & "' and detcde = '" & lblJob4Service.SelectedItem & "'")
        If Not rsUpload.EOF And Not rsUpload.BOF Then
            .txtCustomer = txtNiym
            .txtROno = txtRep_Or
            .txtJobCat = GetJobCat(rsUpload![DETCDE])
            .txtJobDesc = Null2String(rsUpload![Detdsc])
            .txtjCode = Null2String(rsUpload![DETCDE])
            .txtFlatrate = NumericVal(rsUpload![DetPrc])
            .txtstdrate = NumericVal(rsUpload![DET_HRS])
            .txtnote = Null2String(rsUpload![Detail])
            .cboJobChargeTo = Null2String(rsUpload![wCode])
            .txtJobDiscount = Null2String(rsUpload![discrate])
            .txtSaveorEdit = "Edit"
            If IsBodyOrSublet(Trim(rsUpload![DETCDE])) = True Then
                .txtDetCost.Visible = True
                .labDetCost.Visible = True
            Else
                .txtDetCost.Visible = False
                .labDetCost.Visible = False
            End If
        End If
    End With
    frmCSMSJobSelected.Show 1
End Sub

Function IsBodyOrSublet(XXX As String) As Boolean
    Dim rsJOBS                          As ADODB.Recordset
    Set rsJOBS = New ADODB.Recordset
    Set rsJOBS = gconDMIS.Execute("Select * from CSMS_Jobs Where JCode = '" & XXX & "'")
    If Not rsJOBS.EOF And Not rsJOBS.BOF Then
        If Trim(Null2String(rsJOBS!MAIN_CAT)) = "60" Or Trim(Null2String(rsJOBS!MAIN_CAT)) = "99" Or Left(Trim(Null2String(rsJOBS!JCode)), 2) = "SR" Then
            IsBodyOrSublet = True
        Else
            IsBodyOrSublet = False
        End If
    End If
End Function

Function GetJobCat(XXX As Variant)
    Dim rsGetJC                         As ADODB.Recordset
    Set rsGetJC = New ADODB.Recordset
    Set rsGetJC = gconDMIS.Execute("Select JobCategory from CSMS_vw_Jobs where [jcode] = '" & XXX & "'")
    If Not rsGetJC.EOF And Not rsGetJC.BOF Then
        GetJobCat = Null2String(rsGetJC!CSMS_JobCategory)
    End If
    Set rsGetJC = Nothing
End Function

Private Sub cmdNext_Click()
    Dim Flag_mode                       As Boolean
    '    Stop
    Flag_mode = False


    If shpCustomer.Visible = True Then
        shpCustomer.Visible = False
        shpVehicle.Visible = True
        picVehicle.Visible = True
        cmdAddeditCustomer.Visible = False
        If lstVehicle.ListItems.Count > 0 And lstVehicle.Enabled = True Then
            lstVehicle.SetFocus
        End If
    ElseIf shpVehicle.Visible = True Then
        If txtVehName = "" And xTransType <> "A" Then
            MsgBox "Please select vehicle..."
            Exit Sub
        End If

        If xTransType = "E" Then
            'do nothing
        Else

            Dim rsVehicle               As ADODB.Recordset
            Set rsVehicle = New ADODB.Recordset

            Set rsVehicle = gconDMIS.Execute("SELECT Status,ro_no from CSMS_RepairOrder where TransType = '" & xTransType & "' AND PLATE_No = '" & Trim(txtPlate_No.Text) & "' AND UPPER(STATUS) <> 'RELEASED'")

            If IsNull(rsVehicle!RO_NO) Then
                Flag_mode = True
            End If


            If xTransType = "R" Then
                If Flag_mode = False Then
                    If Not rsVehicle.BOF And Not rsVehicle.EOF Then
                        If MsgBox("Repair Order is already open. Continue anyway?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            frmPlateNumberVerification.Show 1
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            End If

        End If


        If frmPlateNumberVerification.lblFlag.Caption = "Cancel" Then Exit Sub
        If labType(0).Caption = "Appointment" Then
            shpVehicle.Visible = False
            picVehicle.Visible = False
            picAppointment.Visible = True
            ShpAppointment.Visible = True
            Label6(4).Caption = "Appointment Information"
        Else
            shpVehicle.Visible = False
            picVehicle.Visible = False
            picAppointment.Visible = True
            shpRO.Visible = True
            Label6(4).Caption = "Repair Order Information"
        End If


        On Error Resume Next
        txtKm_rdg.SetFocus


    ElseIf ShpAppointment.Visible = True Then

        If cboRecd_by = "" Then
            MsgBox "Please select Service Adviser assigned...", vbInformation, "Select"
            Exit Sub
        End If



        ShpAppointment.Visible = False
        picAppointment.Visible = False
        shpRO.Visible = False
        picReason.Visible = True
        shpJobs.Visible = True
        cmdNext.Caption = "Finish"
    ElseIf shpRO.Visible = True Then                         '

        If cboRecd_by = "" Then
            MsgBox "Please select Service Adviser assigned...", vbInformation, "Select"
            cboRecd_by.SetFocus
            Exit Sub
        End If


        If cboRecd_by.ListIndex = -1 Then
            MsgBox "Please select Service Adviser From The List Provided...", vbInformation, "Select"
            cboRecd_by.SetFocus
            Exit Sub
        End If


        picAppointment.Visible = False
        shpRO.Visible = False
        picReason.Visible = True
        shpJobs.Visible = True
        If labType(0).Caption <> "Estimate" Then
            cmdNext.Caption = "Finish"
        End If
    ElseIf shpJobs.Visible = True Then
        If labType(0).Caption = "Estimate" Then
            shpJobs.Visible = False
            ShpEstimate.Visible = True
            picEstimate.Visible = True
            picEstimate.ZOrder 0
            cmdNext.Caption = "Finish"
        Else
            Call SaveAllInfo
        End If
    ElseIf ShpEstimate.Visible = True Then
        Call SaveAllInfo
    End If
End Sub

Sub SaveAllInfo()
    If cboModel = "" And xTransType <> "A" Then
        MsgBox "Please select vehicle..."
        Exit Sub
    End If
    If txtNiym = "" Then
        MsgBox "Please select Customer..."
        Exit Sub
    End If
    If txtRep_Or = "" And txtEstimateno = "" And txtApointmentNo = "" Then
        MsgBox "Please Check RO No./Estimate No. or Appointment No..."
        Exit Sub
    End If
    If cboRecd_by.Text = "" Then
        MsgBox "Please Select Sales Advisor.."
        Exit Sub
    End If
    If labType(0).Caption <> "Appointment" Then

    End If
    If MsgBox("SAVE! All...  " & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
        Exit Sub
    End If
    If labEdit.Caption = "Edit" Then
        If labType(0).Caption = "Repair Order" Then          'R/O
            gconDMIS.Execute "delete from CSMS_Repor where REP_OR = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_Ro_Det where REP_OR = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_RepairOrder where RO_No = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_PMS_Job_Det where REP_OR = '" & txtTranNo & "'"
        ElseIf labType(0).Caption = "Estimate" Then          'ESTIMATE
            gconDMIS.Execute "delete from CSMS_Repor where ESTIMATENO = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_Ro_Det where ESTIMATENO = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_RepairOrder where ESTIMATENO = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_PMS_Job_Det where ESTIMATENO = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_EstHD where ESTIMATENO = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_EstDETAILS where ESTIMATENO = '" & txtTranNo & "'"

        ElseIf labType(0).Caption = "Appointment" Then       'APPOINTMENT
            gconDMIS.Execute "delete from CSMS_Repor where ApptNo = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_Ro_Det where ApptNo = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_RepairOrder where ApptNo = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_PMS_Job_Det where ApptNo = '" & txtTranNo & "'"
        End If
    End If

    Call SaveRepor
    If xTransType <> "A" Then Call TrigerTheRefresh          'BTT - 05212007
End Sub

Sub SaveEstimate()

    On Error GoTo Errorcode

    Screen.MousePointer = 11
    Dim TOTJOBAMT
    Dim xESTIMATENO, xACCT_NO, XTYPE, xPARTNO, xPARTDESC As String
    Dim xQTY, xSRP                      As Double
    Dim x                               As Long
    xESTIMATENO = N2Str2Null(txtEstimateno)
    xACCT_NO = N2Str2Null(txtAcct_No)

    If ListView1.ListItems.Count() <= 0 Then Exit Sub
    For x = 1 To ListView1.ListItems.Count
        XTYPE = N2Str2Null(ListView1.ListItems(x))
        xPARTNO = N2Str2Null(ListView1.ListItems(x).SubItems(1))
        xPARTDESC = N2Str2Null(ListView1.ListItems(x).SubItems(2))
        xQTY = NumericVal(ListView1.ListItems(x).SubItems(3))
        xSRP = NumericVal(ListView1.ListItems(x).SubItems(4))
        gconDMIS.Execute "insert into CSMS_EstDETAILS " & _
                         "(TRANSTYPE,LIVIL,LINE_NO,DETCDE,DETDSC,DETVOL,DETPRC,DETAMT,EstimateNo,ACCT_NO,Type,PARTNO,PARTDESC,QTY,SRP)" & _
                       " values ('E','2','" & Format(x, "00") & "'," & xPARTNO & "," & xPARTDESC & "," & xQTY & "," & xSRP & "," & (xQTY * xSRP) & "," & xESTIMATENO & "," & xACCT_NO & "," & XTYPE & "," & xPARTNO & "," & xPARTDESC & "," & xQTY & "," & xSRP & ")"
        gconDMIS.Execute "insert into CSMS_EstI_DET " & _
                         "(TRANSTYPE,LIVIL,LINE_NO,DETCDE,DETDSC,DETVOL,DETPRC,DETAMT,DET_AMT,EstimateNo)" & _
                       " values ('E','2','" & Format(x, "00") & "'," & xPARTNO & "," & xPARTDESC & "," & xQTY & "," & xSRP & "," & (xQTY * xSRP) & "," & (xQTY * xSRP) & "," & xESTIMATENO & ")"
    Next x

    Dim xNIYM, xPLATE_NO, xModel, xLabor_Cash, xParts_Cash As String
    Dim xAcesories_Cash, xLabor_Comp, xParts_Comp, xAcesories_Comp, xLabor_Sales, xParts_Sales As Double
    Dim xAcesories_Sales, xLabor_Waranty, xParts_Waranty, xAcesories_Waranty, xLabor_DiscRate, xParts_DiscRate As Double
    Dim xAcesories_DiscRate, xLabor_Disc, xParts_Disc, xAcesories_Disc, xLabor_Vat, xParts_Vat, xAcesories_Vat As Double

    xESTIMATENO = N2Str2Null(txtEstimateno)
    xACCT_NO = N2Str2Null(txtAcct_No)
    xNIYM = N2Str2Null(txtCustName)
    xPLATE_NO = N2Str2Null(txtPlate_No)
    xModel = N2Str2Null(cboModel)
    xLabor_Cash = NumericVal(txtEstLabor)
    xParts_Cash = NumericVal(txtEstParts)
    xAcesories_Cash = NumericVal(txtEstAces)
    xLabor_Comp = NumericVal(txtCompLabor)
    xParts_Comp = NumericVal(txtCompPart)
    xAcesories_Comp = NumericVal(txtCompAces)
    xLabor_Sales = NumericVal(txtSalesLabor)
    xParts_Sales = NumericVal(txtSalesParts)
    xAcesories_Sales = NumericVal(txtSalesAces)
    xLabor_Waranty = NumericVal(txtWarLabor)
    xParts_Waranty = NumericVal(txtWarParts)
    xAcesories_Waranty = NumericVal(txtWarLaborAces)
    xLabor_DiscRate = NumericVal(txtRateLabor)
    xParts_DiscRate = NumericVal(txtRateparts)
    xAcesories_DiscRate = NumericVal(txtRateAces)
    xLabor_Disc = NumericVal(txtDiscLabor)
    xParts_Disc = NumericVal(txtDiscParts)
    xAcesories_Disc = NumericVal(txtDiscAces)
    xLabor_Vat = NumericVal(txtVatLabor)
    xParts_Vat = NumericVal(txtVatParts)
    xAcesories_Vat = NumericVal(txtVatAces)
    gconDMIS.Execute "insert into CSMS_EstHD " & _
                     "(EstimateNo, ACCT_NO, NIYM, PLATE_NO, MODEL, Labor_Cash, Parts_Cash, Acesories_Cash, Labor_Comp, Parts_Comp, Acesories_Comp, Labor_Sales, Parts_Sales, Acesories_Sales, Labor_Waranty, Parts_Waranty, Acesories_Waranty, Labor_DiscRate, Parts_DiscRate, Acesories_DiscRate, Labor_Disc, Parts_Disc, Acesories_Disc, Labor_Vat, Parts_Vat, Acesories_Vat)" & _
                   " values (" & xESTIMATENO & "," & xACCT_NO & "," & xNIYM & "," & xPLATE_NO & "," & xModel & "," & xLabor_Cash & "," & xParts_Cash & "," & xAcesories_Cash & "," & xLabor_Comp & "," & xParts_Comp & "," & xAcesories_Comp & "," & xLabor_Sales & "," & xParts_Sales & "," & xAcesories_Sales & "," & xLabor_Waranty & "," & xParts_Waranty & "," & xAcesories_Waranty & "," & xLabor_DiscRate & "," & xParts_DiscRate & "," & xAcesories_DiscRate & "," & xLabor_Disc & "," & xParts_Disc & "," & xAcesories_Disc & "," & xLabor_Vat & "," & xParts_Vat & "," & xAcesories_Vat & ")"

    Dim VTXTrep_or, VTXTestimateno, VTXTROType As String
    Dim VTXTSvc_No, VTXTAcct_No, VTXTNiym As String
    Dim VTXTPlate_No, VcboModel, VTXTMake As String
    Dim VTXTTerm, VTXTSektion, VTXTKm_rdg As String
    Dim VTXTDte_recd, VTXTCertific8, VTXTDte_comp As String
    Dim VTXTDte_Rel, VtxtAddress        As String
    Dim VtxtVIN, VLastUpdateTime        As String
    Dim Vusercode, VLastUpdate          As String
    Dim VTXTParticipat, VcboRecd_by     As String
    Dim XNOTE                           As String

    VTXTestimateno = xESTIMATENO
    Dim rsRO_DET                        As ADODB.Recordset
    TOTJOBAMT = 0: TOTJOBDISC = 0: TOTJOBDISCVAL = 0: TOTJOBTAX = 0
    JobComTotal = 0: JobSalesTotal = 0: JobWarTotal = 0
    Set rsRO_DET = New ADODB.Recordset

    Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where EstimateNo = " & VTXTestimateno & " and livil = '1' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                JobComTotal = JobComTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then JobSalesTotal = JobSalesTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then JobWarTotal = JobWarTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            Else
                TOTJOBAMT = TOTJOBAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                TOTJOBDISC = TOTJOBDISC + N2Str2Zero(rsRO_DET!discount_2)
                TOTJOBDISCVAL = TOTJOBDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTJOBTAX = TOTJOBTAX + N2Str2Zero(rsRO_DET!taxval)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTJOBAMT = Round(TOTJOBAMT, 2): TOTJOBDISC = Round(TOTJOBDISC, 2): TOTJOBDISCVAL = Round(TOTJOBDISCVAL, 2): TOTJOBTAX = Round(TOTJOBTAX, 2)


    TOTPARTSAMT = 0: TOTPARTSDISC = 0: TOTPARTSDISCVAL = 0: TOTPARTSTAX = 0
    PartsComTotal = 0: PartsSalesTotal = 0: PartsWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where EstimateNo = " & VTXTestimateno & " and livil = '2' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        rsRO_DET.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                PartsComTotal = PartsComTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then PartsSalesTotal = PartsSalesTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then PartsWarTotal = PartsWarTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            Else
                TOTPARTSAMT = TOTPARTSAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                TOTPARTSDISC = TOTPARTSDISC + N2Str2Zero(rsRO_DET!discount_2)
                TOTPARTSDISCVAL = TOTPARTSDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTPARTSTAX = TOTPARTSTAX + N2Str2Zero(rsRO_DET!taxval)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTPARTSAMT = Round(TOTPARTSAMT, 2): TOTPARTSDISC = Round(TOTPARTSDISC, 2): TOTPARTSDISCVAL = Round(TOTPARTSDISCVAL, 2): TOTPARTSTAX = Round(TOTPARTSTAX, 2)

    TOTMATAMT = 0: TOTMATDISC = 0: TOTMATDISCVAL = 0: TOTMATTAX = 0
    MatComTotal = 0: MatSalesTotal = 0: MatWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where EstimateNo = " & VTXTestimateno & " and livil = '3' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                MatComTotal = MatComTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then MatSalesTotal = MatSalesTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then MatWarTotal = MatWarTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            Else
                TOTMATAMT = TOTMATAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                TOTMATDISC = TOTMATDISC + N2Str2Zero(rsRO_DET!discount_2)
                TOTMATDISCVAL = TOTMATDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTMATTAX = TOTMATTAX + N2Str2Zero(rsRO_DET!taxval)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTMATAMT = Round(TOTMATAMT, 2): TOTMATDISC = Round(TOTMATDISC, 2): TOTMATDISCVAL = Round(TOTMATDISCVAL, 2): TOTMATTAX = Round(TOTMATTAX, 2)

    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
    Dim FilterType                      As String
    Set rsRO_DET = New ADODB.Recordset
    If labType(0).Caption = "Repair Order" Then
        FilterType = "rep_or = " & VTXTrep_or
    ElseIf labType(0).Caption = "Estimate" Then
        FilterType = "EstimateNo = " & VTXTestimateno
    ElseIf labType(0).Caption = "Appointment" Then
        FilterType = "APPTNO = " & xApptNo
    Else
    End If
    gconDMIS.Execute "update CSMS_RepOr set" & _
                   " labor = " & Round(TOTJOBAMT - TOTJOBTAX, 2) & "," & _
                   " l_amtvalue = " & Round(TOTJOBAMT, 2) & "," & _
                   " l_disc = " & Round(TOTJOBDISCVAL, 2) & "," & _
                   " l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
                   " l_taxval = " & Round(TOTJOBTAX, 2) & "," & _
                   " l_discount = " & Round(TOTJOBDISC, 2) & "," & _
                   " amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC, 2) & "," & _
                   " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX, 2) & "," & _
                   " wl_amt = " & 0 & "," & _
                   " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC, 2) & _
                   " where " & FilterType

    Screen.MousePointer = 0


    Exit Sub

Errorcode:

    ShowVBError

    Exit Sub
End Sub

Sub SaveROMonitoring()
    Dim xModel, xAppointmentDate, xRO_No, xACCT_NO, xPLATE_NO, xjCode, xDescreption, xRecommendation, xStatus, xWriter, xPromiseDate, xPromiseTime As String
    Dim Vusercode, VLastUpdate, VLastUpdateTime As String
    Dim xHours                          As Double

    On Error GoTo Errorcode

    xAppointmentDate = N2Str2Null(Format(txtDte_recd, "MM/dd/yyyy"))
    xRO_No = N2Str2Null(txtRep_Or)
    xACCT_NO = N2Str2Null(txtAcct_No)
    xPLATE_NO = N2Str2Null(txtPlate_No)
    xModel = N2Str2Null(cboModel)
    xRecommendation = "''"
    xHours = NumericVal(lblStdHrs)
    xStatus = "'Park'"
    xWriter = N2Str2Null(cboRecd_by)
    xRecommendation = N2Str2Null(txtRecomendation)
    xPromiseDate = N2Str2Null(dtPromised)
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    xApptNo = N2Str2Null(txtApointmentNo)
    xESTIMATENO = N2Str2Null(txtEstimateno)
    If txtEstimateEdit <> "" Then
        xESTIMATENO = N2Str2Null(txtEstimateEdit)
    End If
    If txtAppointmentEdit <> "" Then
        xApptNo = N2Str2Null(txtAppointmentEdit)
    End If

    gconDMIS.Execute "insert into CSMS_RepairOrder " & _
                     "(ESTIMATENO,ApptNo,TransType,model,AppointmentDate, RO_No,ACCT_NO,PLATE_NO,Recommendation,Hours,Status,Writer,PromiseDate,USERCDE,SAVEDATE,savetime)" & _
                   " values (" & xESTIMATENO & "," & xApptNo & ",'" & xTransType & "'," & xModel & "," & xAppointmentDate & ", " & xRO_No & ", " & xACCT_NO & ", " & xPLATE_NO & ", " & xRecommendation & ", " & xHours & ", " & xStatus & ", " & xWriter & ", " & xPromiseDate & "," & Vusercode & "," & VLastUpdate & "," & VLastUpdateTime & ")"


    Exit Sub

Errorcode:

    ShowVBError
    Exit Sub

End Sub

Sub SaveRepor()
    On Error GoTo Errorcode
    If txtNiym.Text = "" Then
        MsgSpeechBox "Customer must have a name"
        On Error Resume Next
        txtNiym.SetFocus
        Exit Sub
    End If
    If cboRecd_by.Text = "" Then
        MsgSpeechBox "Service Advisor must not be Empty!"
        On Error Resume Next
        cboRecd_by.SetFocus
        Exit Sub
    Else
        Dim rsEmpNo                     As ADODB.Recordset
        Set rsEmpNo = New ADODB.Recordset
        Set rsEmpNo = gconDMIS.Execute("select naym from CSMS_vw_EmpNo where naym = '" & cboRecd_by.Text & "'")
        If rsEmpNo.EOF And rsEmpNo.BOF Then
            MsgSpeechBox "Invalid Service Advisor"
            On Error Resume Next

            Exit Sub
        End If
        Set rsEmpNo = Nothing
    End If

    Dim rsDupRepor                      As ADODB.Recordset
    Set rsDupRepor = New ADODB.Recordset
    Set rsDupRepor = gconDMIS.Execute("select rep_or from CSMS_RepOr where rep_or = " & N2Str2Null(txtTranNo.Text))
    If Not rsDupRepor.EOF And Not rsDupRepor.BOF Then
        MsgSpeechBox "Repair Order Number Already Exist!"
        On Error Resume Next
        txtRep_Or.SetFocus
        Exit Sub
    End If
    Set rsDupRepor = Nothing

    Dim VTXTrep_or, VTXTestimateno, VTXTROType As String
    Dim VTXTSvc_No, VTXTAcct_No, VTXTNiym As String
    Dim VTXTPlate_No, VcboModel, VTXTMake As String
    Dim VTXTTerm, VTXTSektion, VTXTKm_rdg As String
    Dim VTXTDte_recd, VTXTCertific8, VTXTDte_comp As String
    Dim VTXTDte_Rel, VtxtAddress        As String
    Dim VtxtVIN, VLastUpdateTime        As String
    Dim Vusercode, VLastUpdate          As String
    Dim VTXTParticipat, VcboRecd_by     As String
    Dim XNOTE                           As String

    VTXTrep_or = N2Str2Null(txtRep_Or.Text)
    VTXTestimateno = N2Str2Null(txtEstimateno.Text)

    VTXTROType = "''"

    VTXTSvc_No = "''"
    VTXTAcct_No = N2Str2Null(txtAcct_No.Text)
    VTXTNiym = N2Str2Null(txtNiym.Text)
    Dim kAdd                            As Integer
    For kAdd = 1 To Len(txtAddress.Text)
        If Mid(txtAddress.Text, kAdd, 1) = "-" And Mid(txtAddress.Text, kAdd + 1, 1) = "-" And Mid(txtAddress.Text, kAdd + 1, 1) = "-" Then Exit For
        VtxtAddress = VtxtAddress & Mid(txtAddress.Text, kAdd, 1)
    Next
    VtxtAddress = N2Str2Null(VtxtAddress)
    VTXTPlate_No = N2Str2Null(txtPlate_No.Text)
    VcboModel = N2Str2Null(cboModel.Text)
    VTXTMake = N2Str2Null(txtMake.Text)

    VTXTTerm = "''"
    VTXTSektion = N2Str2Null(txtSektion.Text)
    VTXTKm_rdg = N2Str2Null(txtKm_rdg.Text)
    VTXTDte_recd = N2Date2Null(txtDte_recd)

    VTXTCertific8 = "''"

    VTXTDte_comp = "''"

    VTXTDte_Rel = "''"
    VtxtVIN = N2Str2Null(txtVIN.Text)
    VTXTParticipat = N2Str2Null(txtParticipat.Text)
    VcboRecd_by = N2Str2Null(SetCodeSA(cboRecd_by.Text))
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    xApptNo = N2Str2Null(txtApointmentNo)
    xESTIMATENO = N2Str2Null(txtEstimateno)
    XNOTE = N2Str2Null(txtRecomendation)

    If txtEstimateEdit <> "" Then
        xESTIMATENO = N2Str2Null(txtEstimateEdit)
    End If
    If txtAppointmentEdit <> "" Then
        xApptNo = N2Str2Null(txtAppointmentEdit)
    End If
    gconDMIS.Execute "insert into CSMS_RepOr " & _
                     "(ESTIMATENO,ApptNo,TransType,[note],rep_or,rotype,svc_no,acct_no,niym,plate_no,model,term,sektion,Recd_by,km_rdg,dte_recd,certific8,VIN,participat,status,USERCDE,SAVEDATE,SAVETIME)" & _
                   " values (" & xESTIMATENO & "," & xApptNo & ",'" & xTransType & "'," & XNOTE & "," & VTXTrep_or & ", " & VTXTROType & ", " & VTXTSvc_No & _
                     ", " & VTXTAcct_No & ", " & VTXTNiym & ", " & VTXTPlate_No & ", " & VcboModel & ", " & VTXTTerm & ", " & VTXTSektion & _
                     ", " & VcboRecd_by & ", " & VTXTKm_rdg & ", " & VTXTDte_recd & ", " & VTXTCertific8 & _
                     ", " & VtxtVIN & ", " & VTXTParticipat & ",'N'," & Vusercode & "," & VLastUpdate & "," & VLastUpdateTime & ")"

    Call SaveROMonitoring
    Call SaveJobs
    Call SavePMSJObs

    If labType(0).Caption = "Repair Order" Then

    ElseIf labType(0).Caption = "Estimate" Then
        Call SaveEstimate
    ElseIf labType(0).Caption = "Appointment" Then
        Call UpdateAppointmentSkid
    End If

    Dim rsRO_DET                        As ADODB.Recordset
    TOTJOBAMT = 0: TOTJOBDISC = 0: TOTJOBDISCVAL = 0: TOTJOBTAX = 0
    JobComTotal = 0: JobSalesTotal = 0: JobWarTotal = 0
    Set rsRO_DET = New ADODB.Recordset

    If labType(0).Caption = "Repair Order" Then
        Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where rep_or = " & VTXTrep_or & " and livil = '1' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Estimate" Then
        Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where EstimateNo = " & VTXTestimateno & " and livil = '1' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Appointment" Then
        Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where APPTNO = " & xApptNo & " and livil = '1' order by LINE_NO asc")
    Else
    End If
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                JobComTotal = JobComTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then JobSalesTotal = JobSalesTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then JobWarTotal = JobWarTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            Else
                TOTJOBAMT = TOTJOBAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                TOTJOBDISC = TOTJOBDISC + N2Str2Zero(rsRO_DET!discount_2)
                TOTJOBDISCVAL = TOTJOBDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTJOBTAX = TOTJOBTAX + N2Str2Zero(rsRO_DET!taxval)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTJOBAMT = Round(TOTJOBAMT, 2): TOTJOBDISC = Round(TOTJOBDISC, 2): TOTJOBDISCVAL = Round(TOTJOBDISCVAL, 2): TOTJOBTAX = Round(TOTJOBTAX, 2)


    TOTPARTSAMT = 0: TOTPARTSDISC = 0: TOTPARTSDISCVAL = 0: TOTPARTSTAX = 0
    PartsComTotal = 0: PartsSalesTotal = 0: PartsWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    If labType(0).Caption = "Repair Order" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = " & VTXTrep_or & " and livil = '2' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Estimate" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where EstimateNo = " & VTXTestimateno & " and livil = '2' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Appointment" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where APPTNO = " & xApptNo & " and livil = '2' order by LINE_NO asc")
    Else
    End If
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        rsRO_DET.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                PartsComTotal = PartsComTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then PartsSalesTotal = PartsSalesTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then PartsWarTotal = PartsWarTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            Else
                TOTPARTSAMT = TOTPARTSAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                TOTPARTSDISC = TOTPARTSDISC + N2Str2Zero(rsRO_DET!discount_2)
                TOTPARTSDISCVAL = TOTPARTSDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTPARTSTAX = TOTPARTSTAX + N2Str2Zero(rsRO_DET!taxval)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTPARTSAMT = Round(TOTPARTSAMT, 2): TOTPARTSDISC = Round(TOTPARTSDISC, 2): TOTPARTSDISCVAL = Round(TOTPARTSDISCVAL, 2): TOTPARTSTAX = Round(TOTPARTSTAX, 2)

    TOTMATAMT = 0: TOTMATDISC = 0: TOTMATDISCVAL = 0: TOTMATTAX = 0
    MatComTotal = 0: MatSalesTotal = 0: MatWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    If labType(0).Caption = "Repair Order" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = " & VTXTrep_or & " and livil = '3' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Estimate" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where EstimateNo = " & VTXTestimateno & " and livil = '3' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Appointment" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where APPTNO = " & xApptNo & " and livil = '3' order by LINE_NO asc")
    Else
    End If
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                MatComTotal = MatComTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then MatSalesTotal = MatSalesTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then MatWarTotal = MatWarTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            Else
                TOTMATAMT = TOTMATAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                TOTMATDISC = TOTMATDISC + N2Str2Zero(rsRO_DET!discount_2)
                TOTMATDISCVAL = TOTMATDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTMATTAX = TOTMATTAX + N2Str2Zero(rsRO_DET!taxval)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTMATAMT = Round(TOTMATAMT, 2): TOTMATDISC = Round(TOTMATDISC, 2): TOTMATDISCVAL = Round(TOTMATDISCVAL, 2): TOTMATTAX = Round(TOTMATTAX, 2)

    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT

    Dim FilterType                      As String
    Set rsRO_DET = New ADODB.Recordset

    If labType(0).Caption = "Repair Order" Then
        FilterType = "rep_or = " & VTXTrep_or
    ElseIf labType(0).Caption = "Estimate" Then
        FilterType = "EstimateNo = " & VTXTestimateno
    ElseIf labType(0).Caption = "Appointment" Then
        FilterType = "APPTNO = " & xApptNo
    Else
    End If

    gconDMIS.Execute "update CSMS_RepOr set" & _
                   " labor = " & Round(TOTJOBAMT - TOTJOBTAX, 2) & "," & _
                   " l_amtvalue = " & Round(TOTJOBAMT, 2) & "," & _
                   " l_disc = " & Round(TOTJOBDISCVAL, 2) & "," & _
                   " l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
                   " l_taxval = " & Round(TOTJOBTAX, 2) & "," & _
                   " l_discount = " & Round(TOTJOBDISC, 2) & "," & _
                   " amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC, 2) & "," & _
                   " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX, 2) & "," & _
                   " wl_amt = " & 0 & "," & _
                   " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC, 2) & _
                   " where " & FilterType
    Unload Me

    'If labType(0).Caption = "Estimate" Then frmCSMSEstimateEntry.Show
    Exit Sub
Errorcode:
    Screen.MousePointer = 0: ShowVBError
    Exit Sub
End Sub

Sub UpdateAppointmentSkid()
    Dim xTranDate, xApptTime, xCuscde, xCUSNAM, xPLATE_NO, xModel, xMake, XNOTE As String
    Dim xKM_RDG                         As Double
    xApptNo = N2Str2Null(txtTranNo)
    xCuscde = N2Str2Null(txtAcct_No)
    xCUSNAM = N2Str2Null(txtNiym)
    xPLATE_NO = N2Str2Null(txtPlate_No)
    xModel = N2Str2Null(cboModel)
    xMake = N2Str2Null(txtMake)
    xKM_RDG = NumericVal(txtKm_rdg)
    XNOTE = N2Str2Null(txtRecomendation)
    gconDMIS.Execute "update CSMS_Appointment set" & _
                   " CUSCDE = " & xCuscde & "," & _
                   " CUSNAM = " & xCUSNAM & "," & _
                   " PLATE_NO = " & xPLATE_NO & "," & _
                   " model = " & xModel & "," & _
                   " Make = " & xMake & "," & _
                   " KM_RDG = " & xKM_RDG & "," & _
                   " Note = " & XNOTE & _
                   " where ApptNo = " & xApptNo & ""
                   
                LogAudit "A", "CUSTOMER APPOINTMENT", txtNiym & " " & cboModel
   
End Sub
Function SetCodeSA(nam As String) As String
    Dim rsEmpNo                         As ADODB.Recordset
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("Select code,naym,empno from CSMS_vw_EmpNo where naym = '" & nam & "'")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetCodeSA = Null2String(rsEmpNo!code)
    Set rsEmpNo = Nothing
End Function

Function SetSAname(nam As String)
    Dim rsEmpNo                         As ADODB.Recordset
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("Select code,naym,empno from CSMS_vw_EmpNo where empno = '" & nam & "'")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetSAname = Null2String(rsEmpNo!naym)
    Set rsEmpNo = Nothing
End Function

Sub SaveJobs()

    Dim JOBREP_OR, JOBLEVEL, JOBLINE_NO, JOBDETCDE, VLastUpdateTime As String
    Dim JOBDETDSC, JOBDETUNT, VLastUpdate, Vusercode As String
    Dim JOBDETVOL, JOBDETPRC, JOBDETAMT As Double
    Dim JOBCODE, JOBWCODE               As String
    Dim JOBTAXRATE, JOBDISCRATE         As Double
    Dim JOBTAXVAL, JOBDISVAL            As Double
    Dim JOBPOCODE, JOBRep_Or2, JOBDETAIL As String
    Dim JOBDET_AMT, JOBDIS_VAL, JOBDISCOUNT_2, xFLATRATE As Double
    Dim JOBREMARKS                      As String
    Dim JOBTECHNICIAN                   As String
    Dim JOBDET_HRS                      As String
    Dim xJobType                        As String
    Dim x                               As Long
    JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
    JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0

    xApptNo = N2Str2Null(txtTranNo)
    JOBLINE_NO = "0"
    gconDMIS.Execute "delete from CSMS_RO_Det where ApptNo = " & xApptNo & ""
    For x = 1 To lblJob4Service.ListItems.Count
        JOBREP_OR = N2Str2Null(txtRep_Or)
        JOBLEVEL = "'1'"
        JOBLINE_NO = Format(Val(JOBLINE_NO) + 1, "00")
        JOBDETCDE = N2Str2Null(lblJob4Service.ListItems(x))
        xJobType = N2Str2Null(lblJob4Service.ListItems(x).SubItems(1))
        JOBDETDSC = N2Str2Null(Mid(lblJob4Service.ListItems(x).SubItems(2), 1, 500))
        xFLATRATE = NumericVal(lblJob4Service.ListItems(x).SubItems(3))
        JOBDET_HRS = NumericVal(lblJob4Service.ListItems(x).SubItems(4))
        JOBDISCRATE = NumericVal(lblJob4Service.ListItems(x).SubItems(5)) / 100
        JOBWCODE = N2Str2Null(lblJob4Service.ListItems(x).SubItems(6))
        JOBDETUNT = "NULL"
        JOBDETVOL = NumericVal(0)
        JOBDETPRC = NumericVal(xFLATRATE) * JOBDET_HRS
        JOBCODE = "NULL"
        JOBTAXRATE = (VAT_RATE / 100)
        JOBDETAMT = JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE)
        JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
        If Left(lblJob4Service.ListItems(x).SubItems(1), 6) = "Starex" Then
            JOBPOCODE = "'PM'"
        Else
            JOBPOCODE = "NULL"
        End If
        JOBRep_Or2 = "NULL"
        JOBDETAIL = N2Str2Null(CheckChar(lblJob4Service.ListItems(x).SubItems(7)))
        JOBDET_AMT = JOBDETPRC
        JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
        JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
        JOBREMARKS = N2Str2Null(CheckChar(txtRecomendation.Text))
        JOBTECHNICIAN = "NULL"
        JOBTAXVAL = Round(((JOBDETAMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
        Vusercode = "" & N2Str2Null(LOGCODE) & ""
        VLastUpdate = "'" & LOGDATE & "'"
        VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
        xApptNo = N2Str2Null(txtApointmentNo)
        xESTIMATENO = N2Str2Null(txtEstimateno)

        If txtEstimateEdit <> "" Then
            xESTIMATENO = N2Str2Null(txtEstimateEdit)
        End If
        If txtAppointmentEdit <> "" Then
            xApptNo = N2Str2Null(txtAppointmentEdit)
        End If

        gconDMIS.Execute "insert into CSMS_RO_Det " & _
                         "(ESTIMATENO,JobType,TransType,ApptNo,FLATRATE,rep_or,livil,LINE_NO,detcde,detdsc,technician,det_hrs,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,USERCDE,SAVEDATE,SAVETIME)" & _
                       " values (" & xESTIMATENO & "," & xJobType & ",'" & xTransType & "'," & xApptNo & "," & xFLATRATE & "," & JOBREP_OR & ", " & JOBLEVEL & ", " & JOBLINE_NO & "," & _
                       " " & JOBDETCDE & "," & JOBDETDSC & "," & JOBTECHNICIAN & "," & JOBDET_HRS & "," & _
                       " " & JOBDETUNT & ", " & JOBDETVOL & "," & _
                       " " & JOBDETPRC & ", " & JOBDETAMT & ", " & JOBCODE & _
                         ", " & JOBWCODE & ", " & (JOBTAXRATE * 100) & ", " & (JOBDISCRATE * 100) & _
                         ", " & JOBTAXVAL & ", " & JOBDISVAL & ", " & JOBPOCODE & _
                         ", " & JOBRep_Or2 & ", " & JOBDETAIL & ", " & JOBDET_AMT & _
                         ", " & JOBDIS_VAL & ", " & JOBDISCOUNT_2 & _
                         ", " & Vusercode & _
                         ", " & VLastUpdate & _
                         ", " & VLastUpdateTime & ")"
    Next x
End Sub

Private Sub cmdPMS_Click()
    frmMain.MousePointer = 11

    frmCSMSPMS.txtCheck.Text = "AddJobs"
    frmCSMSPMS.Show 1

    frmMain.MousePointer = 0
End Sub

Private Sub Command1_Click()
    frmMain.MousePointer = 11

    frmCSMSGetCannedLabor.txtCheckMe = "ro"
    frmCSMSGetCannedLabor.Show 1

    frmMain.MousePointer = 0
End Sub

Sub ComputeMe()
    tlHrs = 0: tlFR = 0
    For bevvy = 1 To Me.lblJob4Service.ListItems.Count
        tlHrs = tlHrs + NumericVal(lblJob4Service.ListItems(bevvy).SubItems(4))
        tlFR = tlFR + NumericVal(lblJob4Service.ListItems(bevvy).SubItems(3))
    Next bevvy
    lblStdHrs.Caption = tlHrs
    lbltlFaltRate.Caption = tlFR
    txtEstLabor = tlFR
    xPartsAmt = 0: xAcesAmt = 0
    For bevvy = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(bevvy).Text = "P" Then
            xPartsAmt = xPartsAmt + (NumericVal(ListView1.ListItems(bevvy).SubItems(3)) * NumericVal(ListView1.ListItems(bevvy).SubItems(4)))
        Else
            xAcesAmt = xAcesAmt + (NumericVal(ListView1.ListItems(bevvy).SubItems(3)) * NumericVal(ListView1.ListItems(bevvy).SubItems(4)))
        End If
    Next bevvy
    txtEstParts.Text = xPartsAmt
    txtEstAces.Text = xAcesAmt
    txtTotalAmt.Text = Val(txtEstLabor) + Val(txtEstParts) + Val(txtEstAces)
End Sub

Sub ViewJobs()
    Dim rsUpload                        As ADODB.Recordset
    lblJob4Service.Sorted = False: lblJob4Service.ListItems.Clear
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,det_hrs from CSMS_Ro_Det where REP_OR = '" & txtRep_Or & "' Order by det_hrs  desc")    '[LINE_NO]
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lblJob4Service.ListItems, rsUpload
    End If

    tlHrs = 0: tlFR = 0
    For bevvy = 1 To Me.lblJob4Service.ListItems.Count
        tlHrs = tlHrs + NumericVal(lblJob4Service.ListItems(bevvy).SubItems(3))
        tlFR = tlFR + NumericVal(lblJob4Service.ListItems(bevvy).SubItems(2))
    Next bevvy
    lblStdHrs.Caption = tlHrs
    lbltlFaltRate.Caption = tlFR
End Sub

Private Sub Form_Activate()
    Call ComputeMe
End Sub

Private Sub Form_Load()
    picVehicle.Top = 990
    picVehicle.Left = 2310
    picAppointment.Top = 990
    picAppointment.Left = 2310
    picReason.Top = 990
    picReason.Left = 2310
    picEstimate.Top = 990
    picEstimate.Left = 2310
    optLN.Value = True

    dtPromised.Value = DateValue(Now) & " " & TimeValue(Now)
    txtDte_recd.Value = Format(Now(), "MM/dd/yyyy")
    txtRecorded.Text = Format(Now(), "MM/dd/yyyy")
    Dim ctl                             As Control
    With frmCSMSNewAppointment
        For Each ctl In .ControlS
            If TypeOf ctl Is TextBox Then
                ctl.Text = ""

            End If
        Next ctl
    End With
    lblJob4Service.Sorted = False: lblJob4Service.ListItems.Clear

    Dim rsEmpNo                         As ADODB.Recordset
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("select naym from CSMS_vw_EmpNo")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then
        rsEmpNo.MoveFirst
        cboRecd_by.Clear

        Do While Not rsEmpNo.EOF
            cboRecd_by.AddItem Null2String(rsEmpNo!naym)
            rsEmpNo.MoveNext
        Loop
    End If
    txtEstimateEdit.Text = "": txtAppointmentEdit.Text = ""
    Call FillGrid
    SendKeys "{end}"


End Sub

Sub DiplayRecentSalesAdvisor()
    Dim rsTmp                           As ADODB.Recordset

    'Set rsTmp = gconDMIS.Execute("Select Writer From CSMS_RepairOrder Where ApptNO = " & _
     '                             Mid(frmCSMSEdit.lstEdit.SelectedItem, 2, Len(frmCSMSEdit.lstEdit.SelectedItem) - 2) & "")
    'If Not (rsTmp.BOF And rsTmp.EOF) Then
    '    cboRecd_by.Text = rsTmp!writer
    'End If
End Sub
Sub GetDefaultTransactionType()
    If labType(0).Caption = "Repair Order" Then
        xTransType = "R"
        Set rsAddRepor = New ADODB.Recordset
        rsAddRepor.Open "select id,rep_or from CSMS_RepOr where TransType='R' order by rep_or desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsAddRepor.EOF And Not rsAddRepor.BOF Then
            rsAddRepor.MoveFirst
            txtRep_Or.Text = Format(NumericVal(Mid$(rsAddRepor!REP_OR, 3, 8)) + 1, "R-00000000")
        Else
            txtRep_Or.Text = "R-00000001"
        End If
        labNotes.Caption = "Diagnosis entry/Recommendation for future servicing :"
        labTranType.Caption = "Repair Order  No."
        txtTranNo.Text = txtRep_Or
        txtTranNo.Locked = False
    ElseIf labType(0).Caption = "Estimate" Then              '
        xTransType = "E"
        Set rsAddRepor = New ADODB.Recordset

        rsAddRepor.Open "select id,ESTIMATENO from CSMS_EstHD order by ESTIMATENO desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsAddRepor.EOF And Not rsAddRepor.BOF Then
            rsAddRepor.MoveFirst
            txtEstimateno.Text = Format(NumericVal(Mid$(rsAddRepor!EstimateNo, 3, 8)) + 1, "E-00000000")
        Else
            txtEstimateno.Text = "E-00000001"
        End If
        labNotes.Caption = "ESTIMATE NOTE :"
        labTranType.Caption = "Estimate  No."
        txtTranNo.Text = txtEstimateno
        txtTranNo.Locked = False
    ElseIf labType(0).Caption = "Appointment" Then
        labNotes.Caption = "APPOINTMENT NOTE :"
        xTransType = "A"
        labTranType.Caption = "Appointment  No."
        txtApointmentNo.Text = txtTranNo
        txtTranNo.Locked = True
    End If
End Sub

Private Sub lblJob4Service_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        If MsgBox("DELETE! this job...  " & lblJob4Service.SelectedItem.SubItems(2) & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
            Exit Sub
        End If
        Me.lblJob4Service.ListItems.Remove Me.lblJob4Service.SelectedItem.Index
        ComputeMe
    End If
End Sub

Private Sub lstCustomer_DblClick()
    If Not lstCustomer.ListItems.Count = 0 Then
        cmdNext.Value = True

    End If
End Sub

Sub GetVehicleforCustomer()
    txtPlate_No = "": cboModel = "": txtMake = "": txtyear = "": txtVIN = "": txtVehName = ""
    'On Error Resume Next
    Dim rsVehicle                       As ADODB.Recordset

    lstVehicle.Sorted = False: lstVehicle.ListItems.Clear

    Set rsVehicle = New ADODB.Recordset
    Set rsVehicle = gconDMIS.Execute("select model,Plate_no,serial,engine,prodno from CSMS_CusVeh where (Cuscde = '" & txtID & "')")
    If Not (rsVehicle.EOF And rsVehicle.BOF) Then
        Listview_Loadval Me.lstVehicle.ListItems, rsVehicle
        lstCustomer.Refresh

        Dim rsVehicleKo                 As ADODB.Recordset
        Set rsVehicleKo = New ADODB.Recordset
        Set rsVehicleKo = gconDMIS.Execute("select * from CSMS_Cusveh where (Cuscde = '" & txtID & "') and plate_no = '" & lstVehicle.SelectedItem.SubItems(1) & "'")
        If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
            txtPlate_No = Null2String(rsVehicleKo![Plate_no])
            cboModel = Null2String(rsVehicleKo![Model])
            txtMake = Null2String(rsVehicleKo![Make])
            txtyear = Null2String(rsVehicleKo![Yer])
            txtVIN = Null2String(rsVehicleKo![Vin])
            txtVehName = Trim(cboModel) & "   " & txtPlate_No
        End If

    End If
    '    Stop
End Sub

Sub GetVehicleforEndUser(XXX As String)
    txtPlate_No = "": cboModel = "": txtMake = "": txtyear = "": txtVIN = "": txtVehName = ""
    On Error Resume Next
    Dim rsVehicle                       As ADODB.Recordset

    lstVehicle.Sorted = False: lstVehicle.ListItems.Clear

    Set rsVehicle = New ADODB.Recordset
    Set rsVehicle = gconDMIS.Execute("select model,Plate_no,serial,engine,prodno from CSMS_CusVeh where PLATE_NO = '" & XXX & "'")
    If Not (rsVehicle.EOF And rsVehicle.BOF) Then
        Listview_Loadval Me.lstVehicle.ListItems, rsVehicle
        lstCustomer.Refresh

        Dim rsVehicleKo                 As ADODB.Recordset
        Set rsVehicleKo = New ADODB.Recordset
        Set rsVehicleKo = gconDMIS.Execute("select * from CSMS_Cusveh where plate_no = '" & XXX & "'")
        If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
            txtPlate_No = Null2String(rsVehicleKo![Plate_no])
            cboModel = Null2String(rsVehicleKo![Model])
            txtMake = Null2String(rsVehicleKo![Make])
            txtyear = Null2String(rsVehicleKo![Yer])
            txtVIN = Null2String(rsVehicleKo![Vin])
            txtVehName = Trim(cboModel) & "   " & txtPlate_No
        End If

    End If
End Sub
Private Sub lstCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lstCustomer_DblClick
End Sub

Private Sub lstPMSDet_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        If MsgBox("DELETE! this job...  " & lstPMSDet.SelectedItem.SubItems(1) & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
            Exit Sub
        End If
        Me.lstPMSDet.ListItems.Remove Me.lstPMSDet.SelectedItem.Index
    End If
End Sub

Private Sub lstVehicle_DblClick()
    cmdNext.Value = True
End Sub

Private Sub lstVehicle_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim rsVehicleKo                     As ADODB.Recordset
    Set rsVehicleKo = New ADODB.Recordset
    'Set rsVehicleKo = gconDMIS.Execute("select * from CSMS_Cusveh where (Cuscde = '" & txtID & "' OR ENDUSER = '" & txtID & "') and plate_no = '" & lstVehicle.SelectedItem.SubItems(1) & "'")
    Set rsVehicleKo = gconDMIS.Execute("select * from CSMS_Cusveh where (Cuscde = '" & txtID & "') and plate_no = '" & lstVehicle.SelectedItem.SubItems(1) & "'")
    If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
        txtPlate_No = Null2String(rsVehicleKo![Plate_no])
        cboModel = Null2String(rsVehicleKo![Model])
        txtMake = Null2String(rsVehicleKo![Make])
        txtyear = Null2String(rsVehicleKo![Yer])
        txtVIN = Null2String(rsVehicleKo![Vin])
        txtVehName = Trim(cboModel) & "   " & txtPlate_No
    End If
End Sub



Private Sub optEndUser_Click()
    textSearch_Change
End Sub

Private Sub optFN_Click()
    textSearch_Change
End Sub

Private Sub optFullName_Click()
    textSearch_Change
End Sub

Private Sub optLN_Click()
    textSearch_Change
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Sub FillGrid()
    Dim rsCustomer                      As ADODB.Recordset
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("select lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer order by lastname asc")
    If Not (rsCustomer.EOF And rsCustomer.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer
        lstCustomer.Refresh
    End If
    If optEndUser.Value = True Then
        lstCustomer.ColumnHeaders(1).Text = "EndUser Name"
        lstCustomer.ColumnHeaders(2).Text = "Account Name"
        lstCustomer.ColumnHeaders(3).Text = "Plate No"
        lstCustomer.ColumnHeaders(4).Text = "Model"
        lstCustomer.ColumnHeaders(5).Text = "Description"
    Else
        lstCustomer.ColumnHeaders(1).Text = "Last Name"
        lstCustomer.ColumnHeaders(2).Text = "First Name"
        lstCustomer.ColumnHeaders(3).Text = "Address"
        lstCustomer.ColumnHeaders(4).Text = "Province"
        lstCustomer.ColumnHeaders(5).Text = "Phone No."
    End If

End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsCustomer                      As ADODB.Recordset
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomer = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    If optLN.Value = True Then
        Set rsCustomer = gconDMIS.Execute("select lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer where lastname like'" & XXX & "%' order by lastname asc")
    ElseIf optFN.Value = True Then
        Set rsCustomer = gconDMIS.Execute("select lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer where firstname like'" & XXX & "%' order by firstname asc")
    ElseIf optFullName.Value = True Then
        Set rsCustomer = gconDMIS.Execute("select lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer where AcctName like'" & XXX & "%' order by AcctName asc")
    ElseIf optEndUser.Value = True Then
        Set rsCustomer = gconDMIS.Execute("select CSMS_CUSVEH.ENDUSER,CSMS_CUSVEH.CUSCDE,CSMS_CusVeh.Plate_No,CSMS_CusVeh.Model,CSMS_CusVeh.Description,CSMS_CusVeh.CusCde,All_Customer.CustomerAdd from ALL_Customer inner join CSMS_CusVeh on ALL_Customer.CusCde = CSMS_CusVeh.EndUser where (ALL_Customer.LastName like '" & XXX & "%' OR ALL_Customer.FirstName like '" & XXX & "%') order by ALL_Customer.AcctName asc")
        lstCustomer.ColumnHeaders(1).Text = "EndUser Name"
        lstCustomer.ColumnHeaders(2).Text = "Account Name"
        lstCustomer.ColumnHeaders(3).Text = "Plate No"
        lstCustomer.ColumnHeaders(4).Text = "Model"
        lstCustomer.ColumnHeaders(5).Text = "Description"
    End If
    If Not (rsCustomer.EOF And rsCustomer.BOF) Then
        If optEndUser.Value = True Then
            rsCustomer.MoveFirst
            Do While Not rsCustomer.EOF
                With lstCustomer
                    .Sorted = False
                    .ListItems.Add , , SetEndUserName(Null2String(rsCustomer![ENDUSER]))
                    .ListItems(.ListItems.Count).ListSubItems.Add 1, , SetEndACCTName(Null2String(rsCustomer![CUSCDE]))
                    .ListItems(.ListItems.Count).ListSubItems.Add 2, , Null2String(rsCustomer![Plate_no])
                    .ListItems(.ListItems.Count).ListSubItems.Add 3, , Null2String(rsCustomer![Model])
                    .ListItems(.ListItems.Count).ListSubItems.Add 4, , Null2String(rsCustomer![Description])
                    .ListItems(.ListItems.Count).ListSubItems.Add 5, , SetEndACCTName(Null2String(rsCustomer![CUSCDE]))
                    .ListItems(.ListItems.Count).ListSubItems.Add 6, , Null2String(rsCustomer![CUSCDE])
                End With
                rsCustomer.MoveNext
            Loop
        Else
            Listview_Loadval Me.lstCustomer.ListItems, rsCustomer
            lstCustomer.Refresh
        End If
    End If
End Sub

Function SetEndUserName(XXX As String) As String
    Dim rsEndUser                       As ADODB.Recordset
    Set rsEndUser = New ADODB.Recordset
    Set rsEndUser = gconDMIS.Execute("Select * from ALL_CUSTOMER WHERE CUSCDE = '" & XXX & "'")
    If Not rsEndUser.EOF And Not rsEndUser.BOF Then
        SetEndUserName = Null2String(rsEndUser!lastname) & ", " & Null2String(rsEndUser!Firstname)
    End If
    Set rsEndUser = Nothing
End Function

Function SetEndACCTName(XXX As String) As String
    Dim rsCustomer                      As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_CUSTOMER WHERE CUSCDE = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetEndACCTName = Null2String(rsCustomer!lastname) & ", " & Null2String(rsCustomer!Firstname)
    End If
    Set rsCustomer = Nothing
End Function

Sub SavePMSJObs()
    On Error Resume Next
    Dim x                               As Long
    Dim JOBREP_OR, JOBLEVEL, JOBLINE_NO, JOBDETCDE, VLastUpdateTime As String
    Dim JOBDETDSC, JOBDETUNT, VLastUpdate, Vusercode As String
    Dim JOBDETVOL, JOBDETPRC, JOBDETAMT As Double
    Dim JOBCODE, JOBWCODE               As String
    Dim JOBTAXRATE, JOBDISCRATE         As Double
    Dim JOBTAXVAL, JOBDISVAL            As Double
    Dim JOBPOCODE, JOBRep_Or2, JOBDETAIL As String
    Dim JOBDET_AMT, JOBDIS_VAL, JOBDISCOUNT_2, xFLATRATE As Double
    Dim JOBREMARKS                      As String
    Dim JOBTECHNICIAN                   As String
    Dim JOBDET_HRS                      As String
    Dim xJobType, xPMD_Model            As String
    JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
    JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0
    xApptNo = "NULL"
    JOBLINE_NO = "0"
    For x = 1 To lstPMSDet.ListItems.Count
        JOBREP_OR = N2Str2Null(txtRep_Or)
        JOBLEVEL = "'1'"
        JOBLINE_NO = Format(Val(JOBLINE_NO) + 1, "00")
        JOBDETCDE = N2Str2Null(lstPMSDet.ListItems(x))
        xJobType = N2Str2Null(lstPMSDet.ListItems(x).SubItems(1))
        JOBDETDSC = N2Str2Null(Mid(lstPMSDet.ListItems(x).SubItems(2), 1, 500))
        JOBDETUNT = "NULL"
        JOBDETVOL = NumericVal(0)
        JOBDET_HRS = NumericVal(lbltlFaltRate)
        xFLATRATE = NumericVal(lblStdHrs)
        JOBDETPRC = NumericVal(xFLATRATE) * JOBDET_HRS
        JOBCODE = "NULL"
        JOBWCODE = "NULL"
        JOBTAXRATE = (VAT_RATE / 100)
        JOBDISCRATE = NumericVal(0)
        JOBDETAMT = Round(JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE), 2)
        JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
        JOBPOCODE = "NULL"
        JOBRep_Or2 = "NULL"
        JOBDETAIL = "NULL"
        JOBDET_AMT = JOBDETPRC
        JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
        JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
        JOBREMARKS = "NULL"
        JOBTECHNICIAN = "NULL"
        JOBTAXVAL = Round(((JOBDETAMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
        Vusercode = "" & N2Str2Null(LOGCODE) & ""
        VLastUpdate = "'" & LOGDATE & "'"
        VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
        xApptNo = N2Str2Null(txtApointmentNo)
        xESTIMATENO = N2Str2Null(txtEstimateno)
        xPMD_Model = N2Str2Null(N2Str2Null(lstPMSDet.ListItems(x).SubItems(3)))

        If txtEstimateEdit <> "" Then
            xESTIMATENO = N2Str2Null(txtEstimateEdit)
        End If
        If txtAppointmentEdit <> "" Then
            xApptNo = N2Str2Null(txtAppointmentEdit)
        End If

        gconDMIS.Execute "insert into CSMS_PMS_Job_Det " & _
                         "(PMS_Model,ESTIMATENO,ApptNo,JobType,TransType,rep_or,LINE_NO,detcde,detdsc)" & _
                       " values (" & xPMD_Model & "," & xESTIMATENO & "," & xApptNo & "," & xJobType & ",'" & xTransType & "'," & JOBREP_OR & ", " & JOBLINE_NO & "," & JOBDETCDE & "," & JOBDETDSC & ")"
    Next x

End Sub

Private Sub textSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstCustomer.SetFocus
End Sub

Private Sub Timer1_Timer()
    If labType(0).Visible = True Then labType(0).Visible = False Else labType(0).Visible = True
End Sub
Private Sub txtDte_recd_Click()
    txtRecorded = txtDte_recd
End Sub
Private Sub txtRateLabor_Change()
    txtDiscLabor = NumericVal(txtEstLabor) * (NumericVal(txtRateLabor) / 100)
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub
Private Sub txtRateparts_Change()
    txtDiscParts = NumericVal(txtEstParts) * (NumericVal(txtRateparts) / 100)
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub
Private Sub txtRateAces_Change()
    txtDiscAces = NumericVal(txtEstAces) * (NumericVal(txtRateAces) / 100)
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub
Private Sub txtDiscAces_Change()
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub
Private Sub txtDiscLabor_Change()
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub
Private Sub txtDiscParts_Change()
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub

Sub TrigerTheRefresh()
    'frmCSMSServiceCounter.cmdRefresh.Value = True
End Sub

Private Sub txtTranNo_Change()
    txtRep_Or.Text = txtTranNo
End Sub

Private Sub txtTranNo_LostFocus()
    If xTransType <> "A" Then
        If Left(txtTranNo.Text, 2) = xTransType & "-" Then
            txtTranNo.Text = xTransType & "-" & Format(NumericVal(Right(txtTranNo.Text, Len(txtTranNo.Text) - 2)), "00000000")
        Else
            txtTranNo.Text = xTransType & "-" & Format(NumericVal(Right(txtTranNo.Text, Len(txtTranNo.Text))), "00000000")
        End If
    Else
    End If
    txtRep_Or.Text = txtTranNo
End Sub

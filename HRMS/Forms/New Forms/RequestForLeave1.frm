VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMS_RequestForLeave 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REQUEST FOR LEAVE & OVERTIME"
   ClientHeight    =   5790
   ClientLeft      =   1110
   ClientTop       =   2625
   ClientWidth     =   11220
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
   Icon            =   "RequestForLeave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   11220
   Begin Crystal.CrystalReport rpt 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5580
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   42
      Top             =   4830
      Width           =   5580
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
         Left            =   4860
         MouseIcon       =   "RequestForLeave.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Exit Window"
         Top             =   30
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
         Left            =   4170
         MouseIcon       =   "RequestForLeave.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Print this Record"
         Top             =   30
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
         Left            =   3480
         MouseIcon       =   "RequestForLeave.frx":0EFA
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave.frx":104C
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
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
         Left            =   2790
         MouseIcon       =   "RequestForLeave.frx":1377
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave.frx":14C9
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
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
         Left            =   2100
         MouseIcon       =   "RequestForLeave.frx":1825
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave.frx":1977
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Add Record"
         Top             =   30
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
         Left            =   1410
         MouseIcon       =   "RequestForLeave.frx":1C8A
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Find a Record"
         Top             =   30
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
         Left            =   720
         MouseIcon       =   "RequestForLeave.frx":20D6
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave.frx":2228
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Move to Next Record"
         Top             =   30
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
         MouseIcon       =   "RequestForLeave.frx":2580
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9720
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   39
      Top             =   4830
      Width           =   1440
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
         Left            =   720
         MouseIcon       =   "RequestForLeave.frx":2A31
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave.frx":2B83
         Style           =   1  'Graphical
         TabIndex        =   40
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
         Left            =   30
         MouseIcon       =   "RequestForLeave.frx":2EC1
         MousePointer    =   99  'Custom
         Picture         =   "RequestForLeave.frx":3013
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   11115
      TabIndex        =   0
      Top             =   30
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
            Height          =   345
            Left            =   150
            TabIndex        =   5
            Text            =   "Combo2"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txtEmployeeName 
            Height          =   375
            Left            =   2040
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   420
            Width           =   6405
         End
         Begin MSComCtl2.DTPicker dtFiling 
            Height          =   375
            Left            =   8490
            TabIndex        =   7
            Top             =   420
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39513
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Employee Name"
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
            Width           =   1320
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Employee Name"
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
            Width           =   1320
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
            Left            =   8520
            TabIndex        =   4
            Top             =   210
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
         TabIndex        =   36
         Top             =   930
         Width           =   5025
         Begin RichTextLib.RichTextBox txtReason 
            Height          =   3555
            Left            =   60
            TabIndex        =   37
            Top             =   180
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   6271
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"RequestForLeave.frx":3363
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
         TabIndex        =   23
         Top             =   930
         Width           =   3045
         Begin VB.ComboBox cboCutOff 
            Height          =   345
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   58
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
            TabIndex        =   34
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
               TabIndex        =   35
               Top             =   0
               Width           =   2895
            End
         End
         Begin MSComCtl2.DTPicker dtFromDate 
            Height          =   345
            Left            =   120
            TabIndex        =   26
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
            TabIndex        =   27
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
            TabIndex        =   33
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
            TabIndex        =   30
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
            TabIndex        =   31
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
            TabIndex        =   59
            Top             =   180
            Width           =   585
         End
         Begin VB.Label LABOTID 
            AutoSize        =   -1  'True
            Caption         =   "Label3"
            Height          =   225
            Left            =   1620
            TabIndex        =   57
            Top             =   2280
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label LABLEAVEID 
            AutoSize        =   -1  'True
            Caption         =   "Label6"
            Height          =   225
            Left            =   2340
            TabIndex        =   56
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
            TabIndex        =   24
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
            TabIndex        =   25
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
            TabIndex        =   32
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
            TabIndex        =   28
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
            TabIndex        =   29
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
            TabIndex        =   15
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
               TabIndex        =   22
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
               TabIndex        =   17
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
               TabIndex        =   19
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
               TabIndex        =   18
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
               TabIndex        =   20
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
               TabIndex        =   21
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
               TabIndex        =   16
               Top             =   240
               Width           =   2625
            End
         End
         Begin VB.OptionButton optLeave 
            Caption         =   "&Leave"
            Height          =   225
            Left            =   240
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
            TabIndex        =   13
            Top             =   630
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
               TabIndex        =   14
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
            TabIndex        =   11
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
               TabIndex        =   12
               Top             =   240
               Width           =   2595
            End
         End
      End
      Begin VB.Label LABID 
         Caption         =   "0"
         Height          =   345
         Left            =   3810
         TabIndex        =   38
         Top             =   4950
         Width           =   1125
      End
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5745
      Left            =   0
      ScaleHeight     =   5715
      ScaleWidth      =   11175
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   11205
      Begin MSComctlLib.ListView ListView1 
         Height          =   5115
         Left            =   60
         TabIndex        =   55
         Top             =   510
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
         Left            =   10800
         TabIndex        =   54
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox TXTSEARCH 
         Height          =   405
         Left            =   1710
         TabIndex        =   52
         Top             =   60
         Width           =   2715
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee  Name"
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
         Height          =   225
         Left            =   120
         TabIndex        =   53
         Top             =   150
         Width           =   1410
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   495
         Left            =   30
         TabIndex        =   60
         Top             =   0
         Width           =   11115
         _Version        =   655364
         _ExtentX        =   19606
         _ExtentY        =   873
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
End
Attribute VB_Name = "frmHRMS_RequestForLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRequest                                                         As ADODB.Recordset
Dim ADDOREDIT                                                         As String

Function GenerateCode() As String
    Dim rsID                                                          As ADODB.Recordset
    Set rsID = gconDMIS.Execute("SELECT MAX(ISNULL(REQNO,0)) AS HRMS_REQUESTLEAVE_OT FROM HRMS_REQUESTLEAVE_OT ")
    If rsID.FIELDS(0).Value = 0 Then
        GenerateCode = Format(1, "000000")
    Else
        GenerateCode = Format(Val(N2Str2Zero(rsID(0))) + 1, "000000")
    End If
    Set rsID = Nothing
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
        If UCase(C.List(I)) = UCase(Trim(C.Text)) Then
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
    Dim rsTMP As New ADODB.Recordset
    Dim ITEM As ListItem
    
    If XXX = "" Then
        'Listview_Loadval ListView1.ListItems, gconDMIS.Execute("SELECT R.DTE_FILING, R.REQNO, R.EMPNO, E.Lastname + ', ' +  E.Firstname, R.REQTYPE, R.REQDESC, R.STATUS,R.DTE_ACTION, R.ID, R.DTE_FROM  FROM HRMS_EMPINFO E INNER JOIN HRMS_REQUESTLEAVE_OT  R ON E.EMPNO = R.EMPNO ORDER BY R.DTE_FILING ")
        Set rsTMP = gconDMIS.Execute("SELECT R.DTE_FILING as a1, R.REQNO as a2, R.EMPNO as a3, E.Lastname + ', ' +  E.Firstname as fname, R.REQTYPE as a4, R.REQDESC as a5, R.STATUS as a6, R.DTE_ACTION as a7, R.ID as a8, R.DTE_FROM as a9, r.dte_to as a10  FROM HRMS_EMPINFO E INNER JOIN HRMS_REQUESTLEAVE_OT  R ON E.EMPNO = R.EMPNO ORDER BY R.DTE_FILING ")
    Else
        'Listview_Loadval ListView1.ListItems, gconDMIS.Execute("SELECT R.DTE_FILING, R.REQNO, R.EMPNO, E.Lastname + ', ' +  E.Firstname, R.REQTYPE, R.REQDESC, R.STATUS,R.DTE_ACTION, R.ID, R.DTE_FROM  FROM HRMS_EMPINFO E INNER JOIN HRMS_REQUESTLEAVE_OT  R ON E.EMPNO = R.EMPNO WHERE E.LASTNAME + ' ' +  E.FIRSTNAME LIKE'" & Repleys(XXX) & "%' ORDER BY DTE_FILING")
        Set rsTMP = gconDMIS.Execute("SELECT R.DTE_FILING as a1, R.REQNO as a2, R.EMPNO as a3, E.Lastname + ', ' +  E.Firstname as fname, R.REQTYPE as a4, R.REQDESC as a5, R.STATUS as a6, R.DTE_ACTION as a7, R.ID as a8, R.DTE_FROM as a9, r.dte_to as a10  FROM HRMS_EMPINFO E INNER JOIN HRMS_REQUESTLEAVE_OT  R ON E.EMPNO = R.EMPNO WHERE E.LASTNAME + ' ' +  E.FIRSTNAME LIKE'" & Repleys(XXX) & "%' ORDER BY DTE_FILING")
    End If
    ListView1.ListItems.Clear
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        Do While Not rsTMP.EOF
            Set ITEM = ListView1.ListItems.Add(, , Null2String(rsTMP!A1))
            ITEM.SubItems(1) = Null2String(rsTMP!a2)
            ITEM.SubItems(2) = Null2String(rsTMP!a3)
            ITEM.SubItems(3) = Null2String(rsTMP!fname)
            ITEM.SubItems(4) = Null2String(rsTMP!a4)
            ITEM.SubItems(5) = Null2String(rsTMP!a5)
            ITEM.SubItems(6) = Null2String(rsTMP!a6)
            ITEM.SubItems(7) = Null2String(rsTMP!a7)
            ITEM.SubItems(8) = Null2String(rsTMP!a8)
            ITEM.SubItems(9) = Null2String(rsTMP!a9) & "-" & Null2String(rsTMP!a10)
        
            rsTMP.MoveNext
        Loop
    End If
    Set rsTMP = Nothing
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
    Combo_Loadval cboEmployeeNumber, gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A'")
    Dim rsot                                                          As ADODB.Recordset
    Set rsot = gconDMIS.Execute("SELECT LTRIM(RTRIM(PAY_DESC)) AS OTCODE, ID FROM HRMS_OTCODES ORDER BY PAY_DESC ASC")
    While Not rsot.EOF
        cboOvertime.AddItem (Null2String(rsot!OTCODE))
        cboOvertime.ItemData(cboOvertime.NewIndex) = rsot!ID
        rsot.MoveNext
    Wend
    Set rsot = gconDMIS.Execute("SELECT LEAVE_DESC, ID FROM HRMS_LEAVEMASTER ORDER BY 1 ASC")
    While Not rsot.EOF
        cboLeaveType.AddItem (Null2String(rsot!LEAVE_desc))
        cboLeaveType.ItemData(cboLeaveType.NewIndex) = rsot!ID
        rsot.MoveNext
    Wend
    Set rsot = Nothing

End Sub

Sub rsrefresh()
    Set rsRequest = New ADODB.Recordset
    rsRequest.Open "SELECT * FROM HRMS_REQUESTLEAVE_OT ORDER BY REQNO DESC", gconDMIS, adOpenKeyset
End Sub

Sub StoreMemVars()
    If Not rsRequest.EOF Or Not rsRequest.BOF Then
        LABID = rsRequest!ID
        If NumericVal(rsRequest!CUT_OFF) = 1 Then
            cboCutOff.ListIndex = 0
        Else
            cboCutOff.ListIndex = 1
        End If
        cboEmployeeNumber = Null2String(rsRequest!EMPNO)
        labREQNO = rsRequest!REQNO
        If Null2String(rsRequest!REQTYPE) = "O" Then
            optOverTime.Value = True
            cboOvertime = Null2String(rsRequest!REQDESC)
        Else
            optLeave.Value = True
            cboLeaveType.Text = Null2String(rsRequest!REQDESC)
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
        labApprovedBy = Null2String(rsRequest!APPROVEDBY)
        labDateApproved = Null2String(rsRequest!DTE_ACTION)

        If Null2String(rsRequest!STATUS) = "A" Then
            labStatus = "APPROVED"
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdPrint.Enabled = False
        ElseIf Null2String(rsRequest!STATUS) = "D" Then
            labStatus = "DISAPPROVED"
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdPrint.Enabled = False
        Else
            labStatus = "NOT PROCESSED"
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            cmdPrint.Enabled = True
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub cboEmployeeNumber_Change()
    GetEmployeeDetails
End Sub

Private Sub cboEmployeeNumber_Click()
    cboEmployeeNumber_Change
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
    cboEmployeeNumber.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    picMain.Enabled = True
    
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

Function CheckStatusOfRequest() As Boolean
    Dim rsTMP                                   As New ADODB.Recordset
    Set rsTMP = gconDMIS.Execute("SELECT STATUS FROM HRMS_REQUESTLEAVE_OT WHERE ID = " & LABID & "")
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        If Null2String(rsTMP.FIELDS(0)) = "" Then
            CheckStatusOfRequest = True
        Else
            CheckStatusOfRequest = False
        End If
    End If
    Set rsTMP = Nothing
End Function

Private Sub cmdDelete_Click()
    If CheckStatusOfRequest = False Then
        MessagePop InfoStop, "Access Denied", "Request Leave/ OT cannot be deleted once its been process already"
        Exit Sub
    End If
    
    If ShowConfirmDelete = True Then
        gconDMIS.Execute ("DELETE FROM HRMS_REQUESTLEAVE_OT WHERE ID = " & LABID)
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

Private Sub cmdPrevious_Click()
    rsRequest.MovePrevious
    If rsRequest.BOF Then
        rsRequest.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

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
                        " WHERE ID=" & LABID)

    End If

    rsrefresh
    rsRequest.Find ("REQNO='" & labREQNO & "'")
    cmdCancel.Value = True
End Sub



Private Sub Command1_Click()
    picSearch.ZOrder 1
    picSearch.Visible = False
End Sub

Private Sub dtReportingDate_LostFocus()
    If dtReportingDate < dtToDate Then
        MsgBox "Invalid Reporting Date", vbInformation
    End If
End Sub

Private Sub dtToDate_LostFocus()
    If dtFromDate > dtToDate Then
        MsgBox "Invalid Date ", vbInformation
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
    INITCBO
    optLeave.Value = True
    picAdds.Visible = True
    picSaves.Visible = False
    picMain.Enabled = False
    rsrefresh
    StoreMemVars
    ResizeGrid
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
    End If
End Sub

Private Sub txtsearch_Change()
    FillSearchGrid TXTSEARCH
End Sub


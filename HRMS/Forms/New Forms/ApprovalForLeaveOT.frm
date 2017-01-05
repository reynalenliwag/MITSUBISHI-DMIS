VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMS_ApprovalForLeave 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "APPROVAL FOR LEAVE & OVERTIME"
   ClientHeight    =   5715
   ClientLeft      =   1110
   ClientTop       =   2625
   ClientWidth     =   11130
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
   Icon            =   "ApprovalForLeaveOT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   11130
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   4140
      ScaleHeight     =   4485
      ScaleWidth      =   4245
      TabIndex        =   37
      Top             =   570
      Visible         =   0   'False
      Width           =   4275
      Begin VB.PictureBox picML 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   30
         ScaleHeight     =   735
         ScaleWidth      =   4185
         TabIndex        =   71
         Top             =   1290
         Visible         =   0   'False
         Width           =   4185
         Begin VB.ComboBox cboML 
            Height          =   345
            ItemData        =   "ApprovalForLeaveOT.frx":058A
            Left            =   30
            List            =   "ApprovalForLeaveOT.frx":0594
            Style           =   2  'Dropdown List
            TabIndex        =   72
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
            TabIndex        =   73
            Top             =   30
            Width           =   1770
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   2730
         ScaleHeight     =   885
         ScaleWidth      =   1440
         TabIndex        =   45
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
            MouseIcon       =   "ApprovalForLeaveOT.frx":05AA
            MousePointer    =   99  'Custom
            Picture         =   "ApprovalForLeaveOT.frx":06FC
            Style           =   1  'Graphical
            TabIndex        =   46
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
            MouseIcon       =   "ApprovalForLeaveOT.frx":0A3A
            MousePointer    =   99  'Custom
            Picture         =   "ApprovalForLeaveOT.frx":0B8C
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Save this Record"
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.ComboBox cboApprovedBy 
         Height          =   345
         Left            =   60
         TabIndex        =   42
         Text            =   "Combo1"
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
         Height          =   1335
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Text            =   "ApprovalForLeaveOT.frx":0EDC
         Top             =   2220
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtApproved 
         Height          =   345
         Left            =   1470
         TabIndex        =   40
         Top             =   390
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   609
         _Version        =   393216
         Format          =   51511297
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
         TabIndex        =   43
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
         TabIndex        =   39
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
         TabIndex        =   41
         Top             =   660
         Width           =   1170
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   0
         TabIndex        =   38
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
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   16777215
      End
   End
   Begin Crystal.CrystalReport rpt 
      Left            =   1260
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4050
      ScaleHeight     =   855
      ScaleWidth      =   7590
      TabIndex        =   50
      Top             =   4860
      Width           =   7590
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
         Left            =   6300
         MouseIcon       =   "ApprovalForLeaveOT.frx":0EE2
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":1034
         Style           =   1  'Graphical
         TabIndex        =   60
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
         Left            =   5610
         MouseIcon       =   "ApprovalForLeaveOT.frx":139A
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":14EC
         Style           =   1  'Graphical
         TabIndex        =   59
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
         Left            =   4920
         MouseIcon       =   "ApprovalForLeaveOT.frx":1852
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":19A4
         Style           =   1  'Graphical
         TabIndex        =   58
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
         Left            =   4230
         MouseIcon       =   "ApprovalForLeaveOT.frx":1CCF
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":1E21
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
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
         Left            =   3540
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "ApprovalForLeaveOT.frx":217D
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":22CF
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Unpost this Transaction"
         Top             =   30
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
         Left            =   2850
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "ApprovalForLeaveOT.frx":2614
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":2766
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Unpost this Transaction"
         Top             =   30
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
         Left            =   2160
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "ApprovalForLeaveOT.frx":2AAB
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":2BFD
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Post this Transaction"
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
         Left            =   1470
         MouseIcon       =   "ApprovalForLeaveOT.frx":2F22
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":3074
         Style           =   1  'Graphical
         TabIndex        =   53
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
         Left            =   780
         MouseIcon       =   "ApprovalForLeaveOT.frx":336E
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":34C0
         Style           =   1  'Graphical
         TabIndex        =   52
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
         Left            =   90
         MouseIcon       =   "ApprovalForLeaveOT.frx":3818
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":396A
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   4785
      Left            =   0
      ScaleHeight     =   4785
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
         Left            =   30
         TabIndex        =   1
         Top             =   -30
         Width           =   11055
         Begin VB.ComboBox cboEmployeeNumber 
            Height          =   345
            Left            =   180
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
            Format          =   51511297
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
            Caption         =   "Employee No"
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
            Width           =   1065
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
         TabIndex        =   34
         Top             =   930
         Width           =   5025
         Begin RichTextLib.RichTextBox txtReason 
            Height          =   3525
            Left            =   30
            TabIndex        =   35
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   6218
            _Version        =   393217
            TextRTF         =   $"ApprovalForLeaveOT.frx":3CC9
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
         Height          =   3825
         Left            =   2940
         TabIndex        =   23
         Top             =   930
         Width           =   3045
         Begin VB.ComboBox cboCutOff 
            Height          =   345
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   510
            Width           =   2805
         End
         Begin MSComCtl2.DTPicker dtFromDate 
            Height          =   345
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   609
            _Version        =   393216
            Format          =   51511297
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker dtToDate 
            Height          =   345
            Left            =   1560
            TabIndex        =   27
            Top             =   1080
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            Format          =   51511297
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker dtReportingDate 
            Height          =   345
            Left            =   120
            TabIndex        =   33
            Top             =   2280
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            Format          =   51511297
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker dtFromTime 
            Height          =   345
            Left            =   120
            TabIndex        =   30
            Top             =   1680
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            Format          =   51511298
            CurrentDate     =   39513
         End
         Begin MSComCtl2.DTPicker dtToTime 
            Height          =   345
            Left            =   1560
            TabIndex        =   31
            Top             =   1680
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            Format          =   51511298
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
            TabIndex        =   70
            Top             =   270
            Width           =   585
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
            Top             =   840
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
            Top             =   840
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
            Top             =   2040
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
            Top             =   1470
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
            Top             =   1470
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
         Height          =   3825
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
            TabIndex        =   11
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
            TabIndex        =   13
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
               Left            =   120
               TabIndex        =   14
               Text            =   "cboLeaveType"
               Top             =   240
               Width           =   2595
            End
         End
      End
      Begin VB.Label LABID 
         Caption         =   "0"
         Height          =   345
         Left            =   450
         TabIndex        =   36
         Top             =   4800
         Width           =   1125
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9630
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   61
      Top             =   4860
      Width           =   1440
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
         Left            =   720
         MouseIcon       =   "ApprovalForLeaveOT.frx":3D4A
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":3E9C
         Style           =   1  'Graphical
         TabIndex        =   62
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
         MouseIcon       =   "ApprovalForLeaveOT.frx":41DA
         MousePointer    =   99  'Custom
         Picture         =   "ApprovalForLeaveOT.frx":432C
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5745
      Left            =   0
      ScaleHeight     =   5715
      ScaleWidth      =   11115
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   11145
      Begin VB.TextBox TXTSEARCH 
         Height          =   405
         Left            =   1590
         TabIndex        =   67
         Top             =   60
         Width           =   2415
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
         Left            =   10560
         TabIndex        =   66
         Top             =   60
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5175
         Left            =   30
         TabIndex        =   65
         Top             =   510
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9128
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
         TabIndex        =   68
         Top             =   120
         Width           =   1410
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   495
         Left            =   0
         TabIndex        =   74
         Top             =   0
         Width           =   11085
         _Version        =   655364
         _ExtentX        =   19553
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
   Begin VB.Label LABOTID 
      Caption         =   "Label3"
      Height          =   345
      Left            =   180
      TabIndex        =   48
      Top             =   5160
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label LABLEAVEID 
      Caption         =   "Label6"
      Height          =   435
      Left            =   1260
      TabIndex        =   49
      Top             =   5130
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "frmHRMS_ApprovalForLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRequest                                                         As ADODB.Recordset
Dim ADDOREDIT                                                         As String
Dim STATUSX                                                           As String

Function GenerateCode() As String
    Dim rsID                                                          As ADODB.Recordset
    Set rsID = gconDMIS.Execute("Select MAX( ISNULL (REQNO,0) ) as HRMS_RequestLeave_OT from HRMS_RequestLeave_OT ")
    If rsID.FIELDS(0).Value = 0 Then
        GenerateCode = Format(1, "000000")
    Else
        GenerateCode = Format(val(N2Str2Zero(rsID(0))) + 1, "000000")
    End If
    Set rsID = Nothing

End Function

Function GetLeaveCode(XXX As String)
    Dim rsLeaveLook                                                   As ADODB.Recordset
    Set rsLeaveLook = gconDMIS.Execute("select LEAVE_CODE FROM HRMS_leavemaster WHERE ID = " & XXX)
    If Not (rsLeaveLook.BOF Or rsLeaveLook.EOF) Then
        GetLeaveCode = LTrim(RTrim(Null2String(rsLeaveLook!LEAVE_CODE)))
    End If
    Set rsLeaveLook = Nothing
End Function

Function GetLeaveDescription(XXX As String)
    Dim rsLeaveLook                                                   As ADODB.Recordset
    Set rsLeaveLook = gconDMIS.Execute("select * FROM HRMS_leavemaster WHERE LEAVE_CODE='" & Repleys(XXX) & "'")
    If Not (rsLeaveLook.BOF Or rsLeaveLook.EOF) Then
        GetLeaveDescription = LTrim(RTrim(Null2String(rsLeaveLook!LEAVE_desc)))
    End If
    Set rsLeaveLook = Nothing
End Function

Function GetOverTimeCode(XXX As String)
    Dim rsOTLook                                                      As ADODB.Recordset
    Set rsOTLook = gconDMIS.Execute("select PAY_CODE FROM HRMS_OTCodes WHERE ID=" & XXX)
    If Not (rsOTLook.BOF Or rsOTLook.EOF) Then
        GetOverTimeCode = LTrim(RTrim(Null2String(rsOTLook!PAY_CODE)))
    End If
    Set rsOTLook = Nothing
End Function

Function GetOverTimeDescription(XXX As String)
    Dim rsOTLook                                                      As ADODB.Recordset
    Set rsOTLook = gconDMIS.Execute("select * FROM HRMS_OTCodes WHERE PAY_CODE='" & Repleys(XXX) & "'")
    If Not (rsOTLook.BOF Or rsOTLook.EOF) Then
        GetOverTimeDescription = LTrim(RTrim(Null2String(rsOTLook!PAY_DESC)))
    End If
    Set rsOTLook = Nothing
End Function

Function SelectCombo(C As ComboBox) As Integer
    If C.ListCount = 0 Then: SelectCombo = -1: C.Text = "": Exit Function
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

Function OTRATE(XHours)
    'TO CHECK ON IT


End Function

Function GetBasicPay(XXX)

End Function

Sub GetEmployeeDetails()
    txtEmployeeName = ""
    Dim rsEmployee                                                    As ADODB.Recordset
    Set rsEmployee = gconDMIS.Execute("select * FROM HRMS_EMPINFO WHERE EMPNO='" & Repleys(cboEmployeeNumber) & "'")
    If Not (rsEmployee.BOF Or rsEmployee.EOF) Then
        txtEmployeeName = Null2String(rsEmployee!lastname) & " ," & Null2String(rsEmployee!FIRSTNAME)

    End If
    Set rsEmployee = Nothing
End Sub

Sub INITCBO()
    cboCUTOFF.Clear
    cboCUTOFF.AddItem "1st Cut-Off"
    cboCUTOFF.AddItem "2nd Cut-Off"

    Combo_Loadval cboEmployeeNumber, gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A'")
    Dim rsot                                                          As ADODB.Recordset
    Set rsot = gconDMIS.Execute("Select ltrim(rtrim(pay_desc)) as OTCODE ,ID from hrms_otcodes order by pay_desc asc")
    While Not rsot.EOF
        cboOvertime.AddItem (Null2String(rsot!OTCODE))
        cboOvertime.ItemData(cboOvertime.NewIndex) = rsot!ID
        rsot.MoveNext
    Wend
    Set rsot = gconDMIS.Execute("SELECT LEAVE_desc, ID FROM HRMS_LEAVEMASTER   ORDER BY 1 ASC")
    While Not rsot.EOF
        cboLeaveType.AddItem (Null2String(rsot!LEAVE_desc))
        cboLeaveType.ItemData(cboLeaveType.NewIndex) = rsot!ID
        rsot.MoveNext
    Wend
    Combo_Loadval cboApprovedBy, gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME  FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A'")

    Set rsot = Nothing
End Sub

Sub rsrefresh()
    Set rsRequest = New ADODB.Recordset
    rsRequest.Open "SELECT * FROM HRMS_RequestLeave_OT ORDER BY REQNO DESC", gconDMIS, adOpenKeyset
End Sub

Sub StoreMemVars()
    If Not rsRequest.EOF Or Not rsRequest.BOF Then
        labId = rsRequest!ID
        cboEmployeeNumber = Null2String(rsRequest!EMPNO)
        labREQNO = rsRequest!REQNO
        
        If NumericVal(rsRequest!CUT_OFF) = 1 Then
            cboCUTOFF.ListIndex = 0
        Else
            cboCUTOFF.ListIndex = 1
        End If
        
        If Null2String(rsRequest!REQTYPE) = "O" Then
            optOverTime.Value = True: cboOvertime = Null2String(rsRequest!REQDESC)
        Else
            optLeave.Value = True: cboLeaveType.Text = Null2String(rsRequest!REQDESC)
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
        txtNotes = Null2String(rsRequest!NOTES)
        labApprovedBy = Null2String(rsRequest!APPROVEDBY)
        labDateApproved = Null2String(rsRequest!DTE_ACTION)
        cmdPrint.Enabled = True
        If Null2String(rsRequest!STATUS) = "A" Then
            labStatus = "APPROVED"
            cmdEdit.Enabled = False
            cmdUnPost.Enabled = True
            cmdPost.Enabled = False
            cmdCancelCO.Enabled = True
            cmdDelete.Enabled = False
        ElseIf Null2String(rsRequest!STATUS) = "D" Then
            labStatus = "DISAPPROVED"
            cmdEdit.Enabled = False
            cmdUnPost.Enabled = False
            cmdPost.Enabled = True
            cmdCancelCO.Enabled = True
            cmdDelete.Enabled = False
        ElseIf Null2String(rsRequest!STATUS) = "C" Then
            labStatus = "CANCELLED"
            cmdEdit.Enabled = False
            cmdUnPost.Enabled = False
            cmdPost.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPrint.Enabled = False
            cmdDelete.Enabled = False
        Else
            labStatus = "NOT PROCESSED"
            cmdEdit.Enabled = True
            cmdUnPost.Enabled = True
            cmdPost.Enabled = True
            cmdCancelCO.Enabled = True
            cmdDelete.Enabled = True
        End If

        If rsRequest!CUT_OFF = 1 Then cboCUTOFF.Text = "1st Cut-Off"
        If rsRequest!CUT_OFF = 2 Then cboCUTOFF.Text = "2nd Cut-Off"
    Else
        ShowNoRecord
        Unload Me
    End If
End Sub

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
    'If XXX = "" Then
    '    Listview_Loadval ListView1.ListItems, gconDMIS.Execute("SELECT     R.DTE_FILING, R.REQNO, R.EMPNO, E.Lastname + ', ' +  E.Firstname,  R.REQTYPE, R.REQDESC,R.STATUS,R.DTE_ACTION,R.ID  FROM HRMS_EmpInfo E INNER JOIN HRMS_RequestLeave_OT  R ON E.EmpNo = R.EMPNO order by R.DTE_FILING ")
    'Else
    '    Listview_Loadval ListView1.ListItems, gconDMIS.Execute("SELECT     R.DTE_FILING, R.REQNO, R.EMPNO,E.Lastname + ', ' +  E.Firstname,  R.REQTYPE, R.REQDESC,R.STATUS,R.DTE_ACTION,R.ID  FROM HRMS_EmpInfo E INNER JOIN HRMS_RequestLeave_OT  R ON E.EmpNo = R.EMPNO where E.Lastname + ' ' +  E.Firstname LIKE'" & Repleys(XXX) & "%' order by DTE_FILING")
    'End If
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
                         "," & N2Str2Null(txtNotes) & _
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
                       " Justification = " & N2Str2Null(txtNotes) & "," & _
                       " CUT_OFF = " & vCUTOFF & "," & _
                       " PAY_MONTH = " & MM & "," & _
                       " PAY_YEAR = " & YY & _
                       " where REQ_CODE = '" & labREQNO & "'"

    End If
End Sub

Sub DISAPPROVEOT()
    gconDMIS.Execute ("DELETE FROM HRMS_OVERTIME WHERE REQ_CODE='" & labREQNO & "'")
End Sub

Private Sub cboApprovedBy_LostFocus()
    cboApprovedBy.ListIndex = SelectCombo(cboApprovedBy)
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

Private Sub cmdCancel_Click()
    ADDOREDIT = ""
    picAdds.Visible = True: picSaves.Visible = False: picMain.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdCancelCO_Click()
    STATUSX = "C": picStatus.Visible = True: picStatus.ZOrder 0
End Sub

Private Sub cmdDelete_Click()
    If ShowConfirmDelete = True Then
        gconDMIS.Execute ("DELETE FROM HRMS_RequestLeave_OT where id=" & labId)
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
    
    Call SearchREQCODE

End Sub
Private Sub SearchREQCODE()

On Error GoTo Errorcode

    Dim rsREQCODE As New ADODB.Recordset
    Dim buffer As String
     
    buffer = cboLeaveType.Text
    
    Set rsREQCODE = New ADODB.Recordset
    rsREQCODE.Open ("Select ID from  hrms_leavemaster where leave_desc = '" & buffer & "'"), gconDMIS
    If Not (rsREQCODE.EOF And rsREQCODE.BOF) Then
        LABLEAVEID = Trim(rsREQCODE!ID)
    End If

    Set rsREQCODE = Nothing
    
Errorcode:
    Exit Sub
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    picSearch.ZOrder 0: picSearch.Visible = True
    FillSearchGrid ""
End Sub

Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    On Error Resume Next
    rsRequest.MoveFirst
    rsRequest.Find ("id=" & ListView1.SelectedItem.ListSubItems(8).Text)
    StoreMemVars
    picSearch.ZOrder 1: picSearch.Visible = False
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ListView1_DblClick
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
    STATUSX = "A"
    picStatus.Visible = True
    picStatus.ZOrder 0
    cboApprovedBy.SetFocus
    dtApproved.Value = Date
    
    If cboLeaveType.Text = UCase("Maternity Leave") Then
        cboML.ListIndex = 0
        picML.Visible = True
    Else
        picML.Visible = False
    End If
    

End Sub

Sub DeductToLeave(XSTAT As String, CONT As String)
    Dim DAY_CNT As Integer
    Dim RSTMP As New ADODB.Recordset
    Dim XUSED As Integer
    Dim XAVAIL As Integer
    DAY_CNT = DateDiff("d", dtFromDate, dtToDate) + 1
    
    If XSTAT = "A" Then
        Set RSTMP = gconDMIS.Execute("SELECT (AVAILABLE - USED) AS REMAIN, USED, AVAILABLE FROM HRMS_LEAVE " & _
            " WHERE EMPLNO = " & cboEmployeeNumber & _
            " AND TYPE = " & N2Str2Null(GetLEaveType(cboLeaveType)) & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            XUSED = NumericVal(RSTMP!used)
            If RSTMP!REMAIN < DAY_CNT Then
                MsgBox "Employee dont have sufficient leave available", vbInformation, "HRMS"
                Exit Sub
            End If
        End If
        Set RSTMP = Nothing
        
        CONT = "PROCEED"
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

Function GetLEaveType(XDESC As String) As String
    Dim RSTMP                                               As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT LEAVE_CODE FROM HRMS_LEAVEMASTER WHERE LEAVE_DESC = " & N2Str2Null(XDESC) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetLEaveType = Null2String(RSTMP!LEAVE_CODE)
    End If
    Set RSTMP = Nothing
End Function

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
    Dim SQL                                                           As String
    Dim VREQNO, VREQTYPE, VREQCODE, VEMPNO, VDTE_FROM, VDTE_TO, VOT_FROM, VOT_TO, VDTE_REPORTING, VDTE_FILING, VREQ_BY, VREASON_REQ, VUSERCODE, VLASTUPDATED, VREQDESC
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
        gconDMIS.Execute ("INSERT INTO HRMS_RequestLeave_OT(REQNO, REQTYPE, REQDESC, REQCODE, EMPNO, DTE_FROM, DTE_TO, OT_FROM, OT_TO, DTE_REPORTING, DTE_FILING, REQ_BY, REASON_REQ, USERCODE, LASTUPDATED)VALUES (" & _
            VREQNO & _
            "," & VREQTYPE & _
            "," & VREQDESC & _
            "," & VREQCODE & _
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
            "," & VLASTUPDATED & ")")
    Else
        gconDMIS.Execute ("UPDATE HRMS_RequestLeave_OT  SET " & _
            " REQNO= " & VREQNO & "," & _
            " REQTYPE= " & VREQTYPE & "," & _
            " REQCODE = " & VREQCODE & "," & _
            " REQDESC = " & VREQDESC & "," & _
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
            " LASTUPDATED= " & VLASTUPDATED & " WHERE ID=" & labId)
    End If
    rsrefresh
    rsRequest.Find ("REQNO='" & labREQNO & "'")
    cmdCancel.Value = True

End Sub

Private Sub cmdStatusCancel_Click()
    picStatus.Visible = False: picStatus.ZOrder 1: STATUSX = ""
End Sub

Private Sub cmdStatusOK_Click()
    Dim CONT                                    As String
    
    If cboApprovedBy.ListIndex = -1 Then
        MsgBox "Please Select Proper Approver From The List!", vbInformation
        Exit Sub
    End If


    If STATUSX = "A" Then
        If MsgBox("Are You Sure You Want to Approve This Application", vbInformation + vbYesNo) = vbNo Then Exit Sub
        
        'UPDATE BY   : MJP 04142009 0506PM-------------------------------------------------
        'DESCRIPTION : LEAVE MAINTENANCE MODULE
            Call DeductToLeave(STATUSX, CONT)
        'DESCRIPTION : LEAVE MAINTENANCE MODULE
        'UPDATE BY   : MJP 04142009 0506PM-------------------------------------------------
        If CONT = "PROCEED" Then
            vDTE_APPROVED = N2Str2Null(dtApproved)
            vAPPROVEDBY = N2Str2Null(cboApprovedBy)
            vNotes = N2Str2Null(txtNotes)
            gconDMIS.Execute ("UPDATE HRMS_REQUESTLEAVE_OT SET DTE_ACTION=" & vDTE_APPROVED & " ,STATUS='A' , APPROVEDBY=" & vAPPROVEDBY & " ,NOTES=" & vNotes & ", ML_TYPE = " & N2Str2Null(cboML) & " where id=" & labId)
            rsrefresh
            rsRequest.Find ("REQNO='" & labREQNO & "'")
            StoreMemVars
            
            cmdStatusCancel_Click
            If optOverTime.Value = True Then: APPROVEOT
        End If
    ElseIf STATUSX = "D" Then
        If MsgBox("Are You Sure You Want to Dis-Approve This Application", vbInformation + vbYesNo) = vbNo Then Exit Sub
        vDTE_APPROVED = N2Str2Null(dtApproved)
        vAPPROVEDBY = N2Str2Null(cboApprovedBy)
        vNotes = N2Str2Null(txtNotes)
        gconDMIS.Execute ("UPDATE HRMS_REQUESTLEAVE_OT SET DTE_ACTION=" & vDTE_APPROVED & " ,STATUS='D' , APPROVEDBY=" & vAPPROVEDBY & " ,NOTES=" & vNotes & " where id=" & labId)
        rsrefresh
        rsRequest.Find ("REQNO='" & labREQNO & "'")
        StoreMemVars
        
        'UPDATE BY   : MJP 04142009 0506PM-------------------------------------------------
        'DESCRIPTION : LEAVE MAINTENANCE MODULE
            Call DeductToLeave(STATUSX, CONT)
        'DESCRIPTION : LEAVE MAINTENANCE MODULE
        'UPDATE BY   : MJP 04142009 0506PM-------------------------------------------------
        
        cmdStatusCancel_Click
        If optOverTime.Value = True Then: DISAPPROVEOT
    Else
        If MsgBox("Are You Sure You Want to Cancel This Application", vbInformation + vbYesNo) = vbNo Then Exit Sub
        vDTE_APPROVED = N2Str2Null(dtApproved)
        vAPPROVEDBY = N2Str2Null(cboApprovedBy)
        vNotes = N2Str2Null(txtNotes)
        gconDMIS.Execute ("UPDATE HRMS_REQUESTLEAVE_OT SET DTE_ACTION=" & vDTE_APPROVED & " ,STATUS='C' , APPROVEDBY=" & vAPPROVEDBY & " ,NOTES=" & vNotes & " where id=" & labId)
        rsrefresh
        rsRequest.Find ("REQNO='" & labREQNO & "'")
        StoreMemVars
        cmdStatusCancel_Click
        If optOverTime.Value = True Then: DISAPPROVEOT
    End If
End Sub

Private Sub cmdUnPost_Click()
    STATUSX = "D": picStatus.Visible = True: picStatus.ZOrder 0
End Sub

Private Sub Command1_Click()
    picSearch.ZOrder 1: picSearch.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        If Null2String(rsRequest!STATUS) <> "C" And Null2String(rsRequest!STATUS) <> "" Then
            If MsgBox("Are You Sure, You Want To Revert This Application!", vbYesNo + vbInformation) = vbYes Then
                vDTE_APPROVED = N2Str2Null("")
                vAPPROVEDBY = N2Str2Null("")
                vNotes = N2Str2Null("")
                gconDMIS.Execute ("UPDATE HRMS_REQUESTLEAVE_OT SET DTE_ACTION=" & vDTE_APPROVED & " ,STATUS=null , APPROVEDBY=" & vAPPROVEDBY & " ,NOTES=" & vNotes & " where id=" & labId)
                rsrefresh
                rsRequest.Find ("REQNO='" & labREQNO & "'")
                StoreMemVars
            End If



        End If
    Else
        MoveKeyPress KeyCode
    End If

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitMemvars

    INITCBO
    picAdds.Visible = True: picSaves.Visible = False: picMain.Enabled = False
    rsrefresh
    ResizeGrid
    StoreMemVars

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

Private Sub InitMemvars()
    cboOvertime.ListIndex = -1
    cboLeaveType.ListIndex = -1
    cboEmployeeNumber = ""
    txtDepartment = ""
    txtEmployeeName = ""
    txtReason = ""
    dtFromDate = LOGDATE
    dtFromTime = LOGTIME
    dtToDate = LOGDATE
    dtToTime = LOGTIME
    dtFiling = LOGDATE
    dtReportingDate = LOGDATE
    'dtReportingDate.MinDate = LOGDATE
    optLeave_Click
    LABOTID = 0
    LABLEAVEID = 0
    labApprovedBy = ""
    labDateApproved = ""
    labStatus = ""
    labREQNO = ""
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



Private Sub txtNotes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtSEARCH_Change()
    FillSearchGrid txtSEARCH
End Sub


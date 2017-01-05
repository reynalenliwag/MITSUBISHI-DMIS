VERSION 5.00
Begin VB.Form frmAMISAccountingPeriod 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounting Period"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9555
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   9555
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   8700
      MouseIcon       =   "AccountingPeriod.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "AccountingPeriod.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Close Window"
      Top             =   3030
      Width           =   705
   End
   Begin VB.OptionButton optFiscal 
      BackColor       =   &H00808080&
      Caption         =   "Fiscal Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   405
      Left            =   3390
      TabIndex        =   62
      Top             =   3450
      Width           =   2565
   End
   Begin VB.OptionButton optCalendar 
      BackColor       =   &H00808080&
      Caption         =   "Calendar Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   405
      Left            =   3390
      TabIndex        =   61
      Top             =   3120
      Width           =   2565
   End
   Begin VB.CommandButton cmdClosePeriod 
      Caption         =   "&Close Accounting Period"
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
      Height          =   765
      Left            =   120
      MouseIcon       =   "AccountingPeriod.frx":0490
      MousePointer    =   99  'Custom
      Picture         =   "AccountingPeriod.frx":05E2
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Process Importing of Cash Receipts "
      Top             =   3060
      Width           =   2970
   End
   Begin VB.TextBox txtPeriodYear 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   990
      TabIndex        =   0
      Top             =   90
      Width           =   1245
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00808080&
      Height          =   2355
      Left            =   3300
      TabIndex        =   83
      Top             =   540
      Width           =   6105
      Begin VB.CheckBox chkGJMonth12 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   5760
         TabIndex        =   60
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkGJMonth11 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   5250
         TabIndex        =   59
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkGJMonth10 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   4710
         TabIndex        =   58
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkGJMonth9 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   4200
         TabIndex        =   57
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkGJMonth8 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   3690
         TabIndex        =   56
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkGJMonth7 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   3180
         TabIndex        =   55
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkGJMonth6 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   2670
         TabIndex        =   54
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkGJMonth5 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   2160
         TabIndex        =   53
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkGJMonth4 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   1650
         TabIndex        =   52
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkGJMonth3 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   1140
         TabIndex        =   66
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkGJMonth2 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   630
         TabIndex        =   50
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkGJMonth1 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   150
         TabIndex        =   49
         Top             =   1890
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth12 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   5760
         TabIndex        =   48
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth11 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   5250
         TabIndex        =   47
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth10 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   4710
         TabIndex        =   46
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth9 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   4200
         TabIndex        =   45
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth8 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   3690
         TabIndex        =   44
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth7 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   3180
         TabIndex        =   43
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth6 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   2670
         TabIndex        =   42
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth5 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   2160
         TabIndex        =   41
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth4 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   1650
         TabIndex        =   40
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth3 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   1140
         TabIndex        =   39
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth2 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   630
         TabIndex        =   38
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkCRJMonth1 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   150
         TabIndex        =   37
         Top             =   1470
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth12 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   5760
         TabIndex        =   36
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth11 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   5250
         TabIndex        =   35
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth10 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   4710
         TabIndex        =   34
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth9 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   4200
         TabIndex        =   33
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth8 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   3690
         TabIndex        =   32
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth7 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   3180
         TabIndex        =   31
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth6 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   2670
         TabIndex        =   30
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth5 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   2160
         TabIndex        =   29
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth4 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   1650
         TabIndex        =   28
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth3 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   1140
         TabIndex        =   27
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth2 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   630
         TabIndex        =   26
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkSJMonth1 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   150
         TabIndex        =   25
         Top             =   1050
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth12 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   5760
         TabIndex        =   24
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth11 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   5250
         TabIndex        =   23
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth10 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   4710
         TabIndex        =   22
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth9 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   4200
         TabIndex        =   21
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth8 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   3690
         TabIndex        =   20
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth7 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   3180
         TabIndex        =   19
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth6 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   2670
         TabIndex        =   18
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth5 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   2160
         TabIndex        =   17
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth4 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   1650
         TabIndex        =   16
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth3 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   1140
         TabIndex        =   15
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth2 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   630
         TabIndex        =   14
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkCDJMonth1 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth12 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   5760
         TabIndex        =   12
         Top             =   210
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth11 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   5250
         TabIndex        =   11
         Top             =   210
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth10 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   4710
         TabIndex        =   10
         Top             =   210
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth9 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   4200
         TabIndex        =   9
         Top             =   210
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth8 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   3690
         TabIndex        =   8
         Top             =   210
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth7 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   3180
         TabIndex        =   7
         Top             =   210
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth6 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   2670
         TabIndex        =   6
         Top             =   210
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth5 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   210
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth4 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Top             =   210
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth3 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   210
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth2 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   630
         TabIndex        =   2
         Top             =   210
         Width           =   285
      End
      Begin VB.CheckBox chkAPJMonth1 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   210
         Width           =   285
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Update"
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
      Left            =   8010
      MouseIcon       =   "AccountingPeriod.frx":087D
      MousePointer    =   99  'Custom
      Picture         =   "AccountingPeriod.frx":09CF
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Save Changes"
      Top             =   3030
      Width           =   705
   End
   Begin VB.Label Label19 
      BackColor       =   &H00808080&
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   150
      TabIndex        =   84
      Top             =   120
      Width           =   885
   End
   Begin VB.Label labMonth12 
      BackColor       =   &H00808080&
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   8970
      TabIndex        =   82
      Top             =   150
      Width           =   465
   End
   Begin VB.Label labMonth11 
      BackColor       =   &H00808080&
      Caption         =   "Nov"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   8460
      TabIndex        =   81
      Top             =   150
      Width           =   465
   End
   Begin VB.Label labMonth10 
      BackColor       =   &H00808080&
      Caption         =   "Oct"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   7950
      TabIndex        =   80
      Top             =   150
      Width           =   465
   End
   Begin VB.Label labMonth9 
      BackColor       =   &H00808080&
      Caption         =   "Sep"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   7440
      TabIndex        =   79
      Top             =   150
      Width           =   465
   End
   Begin VB.Label labMonth8 
      BackColor       =   &H00808080&
      Caption         =   "Aug"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   6900
      TabIndex        =   78
      Top             =   150
      Width           =   465
   End
   Begin VB.Label labMonth7 
      BackColor       =   &H00808080&
      Caption         =   "Jul"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   6450
      TabIndex        =   77
      Top             =   150
      Width           =   465
   End
   Begin VB.Label labMonth6 
      BackColor       =   &H00808080&
      Caption         =   "Jun"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   5940
      TabIndex        =   76
      Top             =   150
      Width           =   465
   End
   Begin VB.Label labMonth5 
      BackColor       =   &H00808080&
      Caption         =   "May"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   5430
      TabIndex        =   75
      Top             =   150
      Width           =   465
   End
   Begin VB.Label labMonth4 
      BackColor       =   &H00808080&
      Caption         =   "Apr"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   4920
      TabIndex        =   74
      Top             =   150
      Width           =   465
   End
   Begin VB.Label labMonth3 
      BackColor       =   &H00808080&
      Caption         =   "Mar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   4410
      TabIndex        =   73
      Top             =   150
      Width           =   465
   End
   Begin VB.Label labMonth2 
      BackColor       =   &H00808080&
      Caption         =   "Feb"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   3900
      TabIndex        =   72
      Top             =   150
      Width           =   465
   End
   Begin VB.Label labMonth1 
      BackColor       =   &H00808080&
      Caption         =   "Jan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   3420
      TabIndex        =   71
      Top             =   150
      Width           =   465
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "General Journal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   70
      Top             =   2460
      Width           =   3165
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Cash Receipts Journal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   69
      Top             =   2010
      Width           =   3165
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Sales Journal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   68
      Top             =   1590
      Width           =   3165
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Check Disbursement Journal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   67
      Top             =   1170
      Width           =   3165
   End
   Begin VB.Label Label 
      BackColor       =   &H00808080&
      Caption         =   "Accounts Payable Journal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   51
      Top             =   750
      Width           =   3165
   End
End
Attribute VB_Name = "frmAMISAccountingPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAcctngPeriod                                                    As ADODB.Recordset

Sub StoreMemvars()
    If Not rsAcctngPeriod.EOF And Not rsAcctngPeriod.BOF Then
        'txtPeriodYear.Text = N2Str2IntZero(rsAcctngPeriod![Yeer])
        chkAPJMonth1 = Null2Bit(rsAcctngPeriod![APJMonth1])
        chkAPJMonth2 = Null2Bit(rsAcctngPeriod![APJMonth2])
        chkAPJMonth3 = Null2Bit(rsAcctngPeriod![APJMonth3])
        chkAPJMonth4 = Null2Bit(rsAcctngPeriod![APJMonth4])
        chkAPJMonth5 = Null2Bit(rsAcctngPeriod![APJMonth5])
        chkAPJMonth6 = Null2Bit(rsAcctngPeriod![APJMonth6])
        chkAPJMonth7 = Null2Bit(rsAcctngPeriod![APJMonth7])
        chkAPJMonth8 = Null2Bit(rsAcctngPeriod![APJMonth8])
        chkAPJMonth9 = Null2Bit(rsAcctngPeriod![APJMonth9])
        chkAPJMonth10 = Null2Bit(rsAcctngPeriod![APJMonth10])
        chkAPJMonth11 = Null2Bit(rsAcctngPeriod![APJMonth11])
        chkAPJMonth12 = Null2Bit(rsAcctngPeriod![APJMonth12])
        chkCDJMonth1 = Null2Bit(rsAcctngPeriod![CDJMonth1])
        chkCDJMonth2 = Null2Bit(rsAcctngPeriod![CDJMonth2])
        chkCDJMonth3 = Null2Bit(rsAcctngPeriod![CDJMonth3])
        chkCDJMonth4 = Null2Bit(rsAcctngPeriod![CDJMonth4])
        chkCDJMonth5 = Null2Bit(rsAcctngPeriod![CDJMonth5])
        chkCDJMonth6 = Null2Bit(rsAcctngPeriod![CDJMonth6])
        chkCDJMonth7 = Null2Bit(rsAcctngPeriod![CDJMonth7])
        chkCDJMonth8 = Null2Bit(rsAcctngPeriod![CDJMonth8])
        chkCDJMonth9 = Null2Bit(rsAcctngPeriod![CDJMonth9])
        chkCDJMonth10 = Null2Bit(rsAcctngPeriod![CDJMonth10])
        chkCDJMonth11 = Null2Bit(rsAcctngPeriod![CDJMonth11])
        chkCDJMonth12 = Null2Bit(rsAcctngPeriod![CDJMonth12])
        chkSJMonth1 = Null2Bit(rsAcctngPeriod![SJMonth1])
        chkSJMonth2 = Null2Bit(rsAcctngPeriod![SJMonth2])
        chkSJMonth3 = Null2Bit(rsAcctngPeriod![SJMonth3])
        chkSJMonth4 = Null2Bit(rsAcctngPeriod![SJMonth4])
        chkSJMonth5 = Null2Bit(rsAcctngPeriod![SJMonth5])
        chkSJMonth6 = Null2Bit(rsAcctngPeriod![SJMonth6])
        chkSJMonth7 = Null2Bit(rsAcctngPeriod![SJMonth7])
        chkSJMonth8 = Null2Bit(rsAcctngPeriod![SJMonth8])
        chkSJMonth9 = Null2Bit(rsAcctngPeriod![SJMonth9])
        chkSJMonth10 = Null2Bit(rsAcctngPeriod![SJMonth10])
        chkSJMonth11 = Null2Bit(rsAcctngPeriod![SJMonth11])
        chkSJMonth12 = Null2Bit(rsAcctngPeriod![SJMonth12])
        chkCRJMonth1 = Null2Bit(rsAcctngPeriod![CRJMonth1])
        chkCRJMonth2 = Null2Bit(rsAcctngPeriod![CRJMonth2])
        chkCRJMonth3 = Null2Bit(rsAcctngPeriod![CRJMonth3])
        chkCRJMonth4 = Null2Bit(rsAcctngPeriod![CRJMonth4])
        chkCRJMonth5 = Null2Bit(rsAcctngPeriod![CRJMonth5])
        chkCRJMonth6 = Null2Bit(rsAcctngPeriod![CRJMonth6])
        chkCRJMonth7 = Null2Bit(rsAcctngPeriod![CRJMonth7])
        chkCRJMonth8 = Null2Bit(rsAcctngPeriod![CRJMonth8])
        chkCRJMonth9 = Null2Bit(rsAcctngPeriod![CRJMonth9])
        chkCRJMonth10 = Null2Bit(rsAcctngPeriod![CRJMonth10])
        chkCRJMonth11 = Null2Bit(rsAcctngPeriod![CRJMonth11])
        chkCRJMonth12 = Null2Bit(rsAcctngPeriod![CRJMonth12])
        chkGJMonth1 = Null2Bit(rsAcctngPeriod![GJMonth1])
        chkGJMonth2 = Null2Bit(rsAcctngPeriod![GJMonth2])
        chkGJMonth3 = Null2Bit(rsAcctngPeriod![GJMonth3])
        chkGJMonth4 = Null2Bit(rsAcctngPeriod![GJMonth4])
        chkGJMonth5 = Null2Bit(rsAcctngPeriod![GJMonth5])
        chkGJMonth6 = Null2Bit(rsAcctngPeriod![GJMonth6])
        chkGJMonth7 = Null2Bit(rsAcctngPeriod![GJMonth7])
        chkGJMonth8 = Null2Bit(rsAcctngPeriod![GJMonth8])
        chkGJMonth9 = Null2Bit(rsAcctngPeriod![GJMonth9])
        chkGJMonth10 = Null2Bit(rsAcctngPeriod![GJMonth10])
        chkGJMonth11 = Null2Bit(rsAcctngPeriod![GJMonth11])
        chkGJMonth12 = Null2Bit(rsAcctngPeriod![GJMonth12])
        If CheckIfAllBooksISClosed(txtPeriodYear.Text) = True Then
            cmdClosePeriod.Enabled = True
        Else
            cmdClosePeriod.Enabled = False
        End If
    Else
        'Insert new Accounting Year
        InsertNewAcctngYear
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Function GetVoucherNo() As String
    Dim rsJournal_HD                                   As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'CLO' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function Setacctname(VVV As String) As String
    Dim rsChartAccount2                                As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    If Left(VVV, 1) = "'" Then
       rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & VVV, gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
       rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!Description))
    Else
        Setacctname = ""
    End If
End Function

Private Sub cmdClosePeriod_Click()
    Dim Cnt As Integer
    Dim J_ACCT_CODE, J_ACCT_NAME                       As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET       As Double
    Dim J_STATUS, J_JITEMNO                            As String
    
    Dim TOTAL_DEBIT, TOTAL_CREDIT                      As Double
    Dim J_JDATE As String, J_VOUCHERNO As String, J_JTYPE As String
    Dim J_JNO As String, J_REMARKS As String
    Dim rsJournal_HD As ADODB.Recordset
    Dim TOTAL_DEBIT_BALANCE, TOTAL_CREDIT_BALANCE As Double
    Dim DEBIT_BALANCE, CREDIT_BALANCE As Double
    
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select SUM(DEBIT) AS DEBIT_TOTAL, SUM(CREDIT) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where LEFT(ACCT_CODE,1) > 3 AND jtype <> 'CLO' and Status = 'P' AND YEAR(jdate) <= " & txtPeriodYear.Text & " group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst
        TOTAL_DEBIT_BALANCE = 0: TOTAL_CREDIT_BALANCE = 0
        Screen.MousePointer = 11
        gconDMIS.Execute ("Delete from AMIS_Journal_HD Where Jtype = 'CLO' and YEAR(JDate) = " & NumericVal(txtPeriodYear) + 1)
        gconDMIS.Execute ("Delete from AMIS_Journal_Det Where Jtype = 'CLO' and YEAR(JDate) = " & NumericVal(txtPeriodYear) + 1)
        
        Dim rsJournal_HDDup As ADODB.Recordset
        Set rsJournal_HDDup = New ADODB.Recordset
        Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
        If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
            J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
        Else
            J_JNO = "'000001'"
        End If
        Set rsJournal_HDDup = Nothing
        J_JDATE = N2Str2Null("1/1/" & NumericVal(txtPeriodYear) + 1)
        J_VOUCHERNO = N2Str2Null(GetVoucherNo())
        Cnt = 0
        gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                       " (Jno,jdate,voucherno,jtype,remarks)" & _
                       " values (" & J_JNO & "," & J_JDATE & ", " & J_VOUCHERNO & ", 'CLO', 'CLOSING ENTRIES (AUTOMATED BY SYSTEM)')"
                               
        Do While Not rsJournal_HD.EOF
            If NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) > NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) Then
                DEBIT_BALANCE = NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) - NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL))
                CREDIT_BALANCE = 0
            Else
                If NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) > NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) Then
                    CREDIT_BALANCE = NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) - NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL))
                    DEBIT_BALANCE = 0
                Else
                    CREDIT_BALANCE = 0: DEBIT_BALANCE = 0
                End If
            End If
            TOTAL_DEBIT_BALANCE = TOTAL_DEBIT_BALANCE + DEBIT_BALANCE
            TOTAL_CREDIT_BALANCE = TOTAL_CREDIT_BALANCE + CREDIT_BALANCE
            Cnt = Cnt + 1
            
            'gconDMIS.Execute "update AMIS_ChartAccount Set" & _
                           " Debit_Total = " & Round(NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)), 2) & "," & _
                           " Credit_Total = " & Round(NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)), 2) & "," & _
                           " DebitBalance = " & DEBIT_BALANCE & "," & _
                           " CreditBalance = " & CREDIT_BALANCE & _
                           " Where AcctCode = '" & Null2String(rsJournal_HD!Acct_Code) & "'"
                           
            J_JITEMNO = "'" & Format(Cnt, "0000") & "'"
            J_ACCT_CODE = N2Str2Null(rsJournal_HD!Acct_Code)
            J_ACCT_NAME = N2Str2Null(Setacctname(Null2String(rsJournal_HD!Acct_Code)))
            J_DEBIT = Round(CREDIT_BALANCE, 2)
            J_CREDIT = Round(DEBIT_BALANCE, 2)
            J_TAX = 0
            J_GROSS = 0
            J_NET = 0
            J_STATUS = "'N'"
            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
            gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                             "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                           " values (" & J_JDATE & ", " & J_VOUCHERNO & ", 'CLO', " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                             ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
    
            rsJournal_HD.MoveNext
         Loop
         If TOTAL_DEBIT_BALANCE - TOTAL_CREDIT_BALANCE > 0 Then
            Cnt = Cnt + 1
            J_JITEMNO = "'" & Format(Cnt, "0000") & "'"
            J_ACCT_CODE = N2Str2Null("31-00002-00")
            J_ACCT_NAME = N2Str2Null(Setacctname("31-00002-00"))
            J_DEBIT = Round(TOTAL_DEBIT_BALANCE - TOTAL_CREDIT_BALANCE, 2)
            J_CREDIT = 0
            J_TAX = 0
            J_GROSS = 0
            J_NET = 0
            J_STATUS = "'N'"
            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
            
            gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                             "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                           " values (" & J_JDATE & ", " & J_VOUCHERNO & ", 'CLO', " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                             ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
         Else
            Cnt = Cnt + 1
            J_JITEMNO = "'" & Format(Cnt, "0000") & "'"
            J_ACCT_CODE = N2Str2Null("31-00002-00")
            J_ACCT_NAME = N2Str2Null(Setacctname("31-00002-00"))
            J_DEBIT = 0
            J_CREDIT = Round(TOTAL_CREDIT_BALANCE - TOTAL_DEBIT_BALANCE, 2)
            J_TAX = 0
            J_GROSS = 0
            J_NET = 0
            J_STATUS = "'N'"
            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
            
            gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                             "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                           " values (" & J_JDATE & ", " & J_VOUCHERNO & ", 'CLO', " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                             ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
         End If
         MsgBox "Closing Entries for Accounting Year = " & txtPeriodYear.Text & " Successfully Created!", vbInformation, "Done"
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdSave_Click()
    UpdateAcctngYear
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Icon = frmMain.Icon
    txtPeriodYear.Text = Year(LOGDATE)
    rsRefresh
    StoreMemvars
End Sub

Sub rsRefresh()
    Set rsAcctngPeriod = New ADODB.Recordset
    Set rsAcctngPeriod = gconDMIS.Execute("Select * from AMIS_AcctngPeriod Where Yeer = " & txtPeriodYear.Text)
End Sub

Function CheckIfAllBooksISClosed(XYeer As Integer) As Boolean
    CheckIfAllBooksISClosed = True
    If chkAPJMonth1.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkAPJMonth2.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkAPJMonth3.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkAPJMonth4.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkAPJMonth5.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkAPJMonth6.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkAPJMonth7.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkAPJMonth8.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkAPJMonth9.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkAPJMonth10.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkAPJMonth11.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkAPJMonth12.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth1.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth2.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth3.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth4.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth5.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth6.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth7.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth8.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth9.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth10.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth11.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCDJMonth12.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth1.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth2.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth3.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth4.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth5.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth6.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth7.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth8.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth9.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth10.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth11.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkSJMonth12.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth1.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth2.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth3.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth4.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth5.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth6.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth7.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth8.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth9.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth10.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth11.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkCRJMonth12.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth1.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth2.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth3.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth4.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth5.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth6.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth7.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth8.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth9.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth10.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth11.Value = 0 Then
        CheckIfAllBooksISClosed = False
    ElseIf chkGJMonth12.Value = 0 Then
        CheckIfAllBooksISClosed = False
    Else
        CheckIfAllBooksISClosed = True
    End If
End Function

Sub InsertNewAcctngYear()
    Screen.MousePointer = 11

    chkAPJMonth1 = 0: chkAPJMonth2 = 0: chkAPJMonth3 = 0: chkAPJMonth4 = 0: chkAPJMonth5 = 0: chkAPJMonth6 = 0: chkAPJMonth7 = 0: chkAPJMonth8 = 0: chkAPJMonth9 = 0: chkAPJMonth10 = 0: chkAPJMonth11 = 0: chkAPJMonth12 = 0
    chkCDJMonth1 = 0: chkCDJMonth2 = 0: chkCDJMonth3 = 0: chkCDJMonth4 = 0: chkCDJMonth5 = 0: chkCDJMonth6 = 0: chkCDJMonth7 = 0: chkCDJMonth8 = 0: chkCDJMonth9 = 0: chkCDJMonth10 = 0: chkCDJMonth11 = 0: chkCDJMonth12 = 0
    chkSJMonth1 = 0: chkSJMonth2 = 0: chkSJMonth3 = 0: chkSJMonth4 = 0: chkSJMonth5 = 0: chkSJMonth6 = 0: chkSJMonth7 = 0: chkSJMonth8 = 0: chkSJMonth9 = 0: chkSJMonth10 = 0: chkSJMonth11 = 0: chkSJMonth12 = 0
    chkCRJMonth1 = 0: chkCRJMonth2 = 0: chkCRJMonth3 = 0: chkCRJMonth4 = 0: chkCRJMonth5 = 0: chkCRJMonth6 = 0: chkCRJMonth7 = 0: chkCRJMonth8 = 0: chkCRJMonth9 = 0: chkCRJMonth10 = 0: chkCRJMonth11 = 0: chkCRJMonth12 = 0
    chkGJMonth1 = 0: chkGJMonth2 = 0: chkGJMonth3 = 0: chkGJMonth4 = 0: chkGJMonth5 = 0: chkGJMonth6 = 0: chkGJMonth7 = 0: chkGJMonth8 = 0: chkGJMonth9 = 0: chkGJMonth10 = 0: chkGJMonth11 = 0: chkGJMonth12 = 0

    gconDMIS.Execute ("INSERT INTO AMIS_AcctngPeriod ([Yeer],[APJMonth1],[APJMonth2],[APJMonth3],[APJMonth4],[APJMonth5]" & _
                      ",[APJMonth6],[APJMonth7],[APJMonth8],[APJMonth9],[APJMonth10],[APJMonth11],[APJMonth12],[CDJMonth1]" & _
                      ",[CDJMonth2],[CDJMonth3],[CDJMonth4],[CDJMonth5],[CDJMonth6],[CDJMonth7],[CDJMonth8],[CDJMonth9]" & _
                      ",[CDJMonth10],[CDJMonth11],[CDJMonth12],[SJMonth1],[SJMonth2],[SJMonth3],[SJMonth4],[SJMonth5]" & _
                      ",[SJMonth6],[SJMonth7],[SJMonth8],[SJMonth9],[SJMonth10],[SJMonth11],[SJMonth12],[CRJMonth1]" & _
                      ",[CRJMonth2],[CRJMonth3],[CRJMonth4],[CRJMonth5],[CRJMonth6],[CRJMonth7],[CRJMonth8],[CRJMonth9]" & _
                      ",[CRJMonth10],[CRJMonth11],[CRJMonth12],[GJMonth1],[GJMonth2],[GJMonth3],[GJMonth4],[GJMonth5]" & _
                      ",[GJMonth6],[GJMonth7],[GJMonth8],[GJMonth9],[GJMonth10],[GJMonth11],[GJMonth12]) Values " & _
                    " (" & NumericVal(txtPeriodYear.Text) & "," & chkAPJMonth1.Value & "," & chkAPJMonth2.Value & _
                      "," & chkAPJMonth3.Value & "," & chkAPJMonth4.Value & "," & chkAPJMonth5.Value & "," & chkAPJMonth6.Value & _
                      "," & chkAPJMonth7.Value & "," & chkAPJMonth8.Value & "," & chkAPJMonth9.Value & "," & chkAPJMonth10.Value & _
                      "," & chkAPJMonth11.Value & "," & chkAPJMonth12.Value & "," & chkCDJMonth1.Value & "," & chkCDJMonth2.Value & _
                      "," & chkCDJMonth3.Value & "," & chkCDJMonth4.Value & "," & chkCDJMonth5.Value & "," & chkCDJMonth6.Value & _
                      "," & chkCDJMonth7.Value & "," & chkCDJMonth8.Value & "," & chkCDJMonth9.Value & "," & chkCDJMonth10.Value & _
                      "," & chkCDJMonth11.Value & "," & chkCDJMonth12.Value & "," & chkSJMonth1.Value & "," & chkSJMonth2.Value & _
                      "," & chkSJMonth3.Value & "," & chkSJMonth4.Value & "," & chkSJMonth5.Value & "," & chkSJMonth6.Value & _
                      "," & chkSJMonth7.Value & "," & chkSJMonth8.Value & "," & chkSJMonth9.Value & "," & chkSJMonth10.Value & _
                      "," & chkSJMonth11.Value & "," & chkSJMonth12.Value & "," & chkCRJMonth1.Value & "," & chkCRJMonth2.Value & _
                      "," & chkCRJMonth3.Value & "," & chkCRJMonth4.Value & "," & chkCRJMonth5.Value & "," & chkCRJMonth6.Value & _
                      "," & chkCRJMonth7.Value & "," & chkCRJMonth8.Value & "," & chkCRJMonth9.Value & "," & chkCRJMonth10.Value & _
                      "," & chkCRJMonth11.Value & "," & chkCRJMonth12.Value & "," & chkGJMonth1.Value & "," & chkGJMonth2.Value & _
                      "," & chkGJMonth3.Value & "," & chkGJMonth4.Value & "," & chkGJMonth5.Value & "," & chkGJMonth6.Value & _
                      "," & chkGJMonth7.Value & "," & chkGJMonth8.Value & "," & chkGJMonth9.Value & "," & chkGJMonth10.Value & _
                      "," & chkGJMonth11.Value & "," & chkGJMonth12.Value & ")")
    rsRefresh
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Sub UpdateAcctngYear()
    Screen.MousePointer = 11
    gconDMIS.Execute ("Update AMIS_AcctngPeriod SET " & _
                      "[APJMonth1] = " & chkAPJMonth1.Value & ",[APJMonth2] = " & chkAPJMonth2.Value & ",[APJMonth3] = " & chkAPJMonth3.Value & _
                      ",[APJMonth4] = " & chkAPJMonth4.Value & ",[APJMonth5] = " & chkAPJMonth5.Value & ",[APJMonth6] = " & chkAPJMonth6.Value & _
                      ",[APJMonth7] = " & chkAPJMonth7.Value & ",[APJMonth8] = " & chkAPJMonth8.Value & ",[APJMonth9] = " & chkAPJMonth9.Value & _
                      ",[APJMonth10] = " & chkAPJMonth10.Value & ",[APJMonth11] = " & chkAPJMonth11.Value & ",[APJMonth12] = " & chkAPJMonth12.Value & _
                      ",[CDJMonth1] = " & chkCDJMonth1.Value & ",[CDJMonth2] = " & chkCDJMonth2.Value & ",[CDJMonth3] = " & chkCDJMonth3.Value & _
                      ",[CDJMonth4] = " & chkCDJMonth4.Value & ",[CDJMonth5] = " & chkCDJMonth5.Value & ",[CDJMonth6] = " & chkCDJMonth6.Value & _
                      ",[CDJMonth7] = " & chkCDJMonth7.Value & ",[CDJMonth8] = " & chkCDJMonth8.Value & ",[CDJMonth9] = " & chkCDJMonth9.Value & _
                      ",[CDJMonth10] = " & chkCDJMonth10.Value & ",[CDJMonth11] = " & chkCDJMonth11.Value & ",[CDJMonth12] = " & chkCDJMonth12.Value & _
                      ",[SJMonth1] = " & chkSJMonth1.Value & ",[SJMonth2] = " & chkSJMonth2.Value & ",[SJMonth3] = " & chkSJMonth3.Value & _
                      ",[SJMonth4] = " & chkSJMonth4.Value & ",[SJMonth5] = " & chkSJMonth5.Value & ",[SJMonth6] = " & chkSJMonth6.Value & _
                      ",[SJMonth7] = " & chkSJMonth7.Value & ",[SJMonth8] = " & chkSJMonth8.Value & ",[SJMonth9] = " & chkSJMonth9.Value & _
                      ",[SJMonth10] = " & chkSJMonth10.Value & ",[SJMonth11] = " & chkSJMonth11.Value & ",[SJMonth12] = " & chkSJMonth12.Value & _
                      ",[CRJMonth1] = " & chkCRJMonth1.Value & ",[CRJMonth2] = " & chkCRJMonth2.Value & ",[CRJMonth3] = " & chkCRJMonth3.Value & _
                      ",[CRJMonth4] = " & chkCRJMonth4.Value & ",[CRJMonth5] = " & chkCRJMonth5.Value & ",[CRJMonth6] = " & chkCRJMonth6.Value & _
                      ",[CRJMonth7] = " & chkCRJMonth7.Value & ",[CRJMonth8] = " & chkCRJMonth8.Value & ",[CRJMonth9] = " & chkCRJMonth9.Value & _
                      ",[CRJMonth10] = " & chkCRJMonth10.Value & ",[CRJMonth11] = " & chkCRJMonth11.Value & ",[CRJMonth12] = " & chkCRJMonth12.Value & _
                      ",[GJMonth1] = " & chkGJMonth1.Value & ",[GJMonth2] = " & chkGJMonth2.Value & ",[GJMonth3] = " & chkGJMonth3.Value & _
                      ",[GJMonth4] = " & chkGJMonth4.Value & ",[GJMonth5] = " & chkGJMonth5.Value & ",[GJMonth6] = " & chkGJMonth6.Value & _
                      ",[GJMonth7] = " & chkGJMonth7.Value & ",[GJMonth8] = " & chkGJMonth8.Value & ",[GJMonth9] = " & chkGJMonth9.Value & _
                      ",[GJMonth10] = " & chkGJMonth10.Value & ",[GJMonth11] = " & chkGJMonth11.Value & ",[GJMonth12] = " & chkGJMonth12.Value & _
                    " WHERE Yeer = " & txtPeriodYear.Text)
    rsRefresh
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub txtPeriodYear_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        rsRefresh
        StoreMemvars
    End If
End Sub


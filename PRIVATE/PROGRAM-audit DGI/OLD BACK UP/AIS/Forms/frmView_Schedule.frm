VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#10.4#0"; "CALEND~1.OCX"
Begin VB.Form frmAISView_Schedule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Exam Schedule"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11280
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      Height          =   3105
      Left            =   0
      TabIndex        =   1
      Top             =   4050
      Width           =   10845
      Begin MSComctlLib.ListView ListView1 
         Height          =   1665
         Left            =   210
         TabIndex        =   2
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2937
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   10845
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2820
         Left            =   3690
         TabIndex        =   5
         Top             =   420
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   53542913
         CurrentDate     =   39132
      End
      Begin XtremeCalendarControl.CalendarControl CalendarControl1 
         Height          =   3285
         Left            =   5910
         TabIndex        =   4
         Top             =   420
         Width           =   4635
         _Version        =   655364
         _ExtentX        =   8176
         _ExtentY        =   5794
         _StockProps     =   64
      End
      Begin XtremeCalendarControl.DatePicker DatePicker1 
         Height          =   2475
         Left            =   300
         TabIndex        =   3
         Top             =   360
         Width           =   3285
         _Version        =   655364
         _ExtentX        =   5794
         _ExtentY        =   4366
         _StockProps     =   64
      End
   End
End
Attribute VB_Name = "frmAISView_Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

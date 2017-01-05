VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#10.4#0"; "CODEJO~1.OCX"
Begin VB.Form frmCRIS_CalendarSales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13350
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCalendar 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8775
      Left            =   0
      ScaleHeight     =   8775
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin XtremeCalendarControl.DatePicker dtCal 
         Height          =   2565
         Left            =   30
         TabIndex        =   2
         Top             =   645
         Width           =   2835
         _Version        =   655364
         _ExtentX        =   5001
         _ExtentY        =   4524
         _StockProps     =   64
         AutoSize        =   0   'False
         ShowWeekNumbers =   -1  'True
         Show3DBorder    =   0
         MaxSelectionCount=   0
         BoldDaysPerIdleStep=   1
         BoldDaysIdleStepTime_ms=   100
      End
      Begin XtremeShortcutBar.ShortcutCaption sCap 
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   4695
         _Version        =   655364
         _ExtentX        =   8281
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Pick Your Date"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption captionCalendar 
         Height          =   360
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4680
         _Version        =   655364
         _ExtentX        =   8255
         _ExtentY        =   635
         _StockProps     =   14
         Caption         =   "Calendar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
   End
   Begin XtremeCalendarControl.CalendarControl cCalSales 
      Height          =   8730
      Left            =   3060
      TabIndex        =   0
      Top             =   0
      Width           =   10305
      _Version        =   655364
      _ExtentX        =   18177
      _ExtentY        =   15399
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmCRIS_CalendarSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

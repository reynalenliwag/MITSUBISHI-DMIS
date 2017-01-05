VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMS_Dtrweekly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Time Record"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2955
   Icon            =   "frmdtrweekly.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2955
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1140
      MouseIcon       =   "frmdtrweekly.frx":1FF7A
      MousePointer    =   99  'Custom
      Picture         =   "frmdtrweekly.frx":200CC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1020
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1980
      MouseIcon       =   "frmdtrweekly.frx":2056B
      MousePointer    =   99  'Custom
      Picture         =   "frmdtrweekly.frx":206BD
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   1020
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   150
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52101121
      CurrentDate     =   40330
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   345
      Left            =   840
      TabIndex        =   3
      Top             =   570
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52101121
      CurrentDate     =   40359
   End
   Begin Crystal.CrystalReport rptformula 
      Left            =   300
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   5
      Top             =   600
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "From :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   -60
      TabIndex        =   4
      Top             =   180
      Width           =   825
   End
End
Attribute VB_Name = "frmHRMS_Dtrweekly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim fdate As Date
    Dim tdate As Date
    
    fdate = CDate(dtpFromDate.Value)
    tdate = CDate(dtpToDate.Value)
    
    rptformula.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptformula.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptformula.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptformula.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"
    rptformula.Formulas(8) = "mindate = '" & fdate & "'"
    rptformula.Formulas(7) = "maxdate = '" & tdate & "'"
        
    PrintSQLReport rptformula, HRMS_REPORT_PATH & "attendance.rpt", "{hrms_attend.datetoday} >= date(" & YEAR(fdate) & "," & MONTH(fdate) & "," & Day(fdate) & ") AND {hrms_attend.datetoday} <= date(" & YEAR(tdate) & "," & MONTH(tdate) & "," & Day(tdate) & ")", HRMS_REPORT_Connection, 1

End Sub

Private Sub Form_Load()

Screen.MousePointer = 11
CenterMe frmMain, Me, 1

dtpFromDate.Value = firstDay(LOGDATE)
dtpToDate.Value = LOGDATE

Screen.MousePointer = 0

End Sub

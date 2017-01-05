VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERInfoMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport rptJoannPearl 
      Left            =   150
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Individual Development Plan"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox picOtherInfo 
      Height          =   6660
      Left            =   60
      ScaleHeight     =   6600
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   60
      Width           =   5775
      Begin VB.CommandButton cmdExitOtherInfo 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   2805
         Picture         =   "OTHERInfoMain.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Exit Window"
         Top             =   5940
         Width           =   1500
      End
      Begin VB.CommandButton cmdShowOtherInfo 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   1185
         Picture         =   "OTHERInfoMain.frx":0366
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "View Selected Detail"
         Top             =   5940
         Width           =   1635
      End
      Begin VB.OptionButton optJoannPearl 
         Caption         =   "Print Employee Development Plan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Print Employee Development Plan"
         Top             =   5490
         Width           =   5655
      End
      Begin VB.OptionButton optTrainingPlan 
         Caption         =   "Training Plan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "View Training Plan"
         Top             =   3390
         Width           =   5655
      End
      Begin VB.OptionButton optPerformanceEvaluation 
         Caption         =   "Performance Evaluation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "View Performance Evaluation"
         Top             =   3810
         Width           =   5655
      End
      Begin VB.OptionButton optDependents 
         Caption         =   "Dependents"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "View Dependents"
         Top             =   480
         Width           =   5655
      End
      Begin VB.OptionButton optParents 
         Caption         =   "Parents"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "View Parents"
         Top             =   900
         Width           =   5655
      End
      Begin VB.OptionButton optEducation 
         Caption         =   "Education"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "View Education"
         Top             =   1320
         Width           =   5655
      End
      Begin VB.OptionButton optCSEligibility 
         Caption         =   "CS Eligibility"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "View CS Eligibility"
         Top             =   1740
         Width           =   5655
      End
      Begin VB.OptionButton optExamsPassed 
         Caption         =   "Exams Passed"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "View Exams Passed"
         Top             =   2160
         Width           =   5655
      End
      Begin VB.OptionButton optPastEmployment 
         Caption         =   "Past Employment"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "View Past Employment"
         Top             =   2580
         Width           =   5655
      End
      Begin VB.OptionButton optTraining 
         Caption         =   "Training"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "View Training"
         Top             =   3000
         Width           =   5655
      End
      Begin VB.OptionButton optPersonnelAction 
         Caption         =   "Personnel Action"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "View Personnel Action"
         Top             =   4230
         Width           =   5655
      End
      Begin VB.OptionButton optMemo 
         Caption         =   "Memo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "View Memo"
         Top             =   4650
         Width           =   5655
      End
      Begin VB.OptionButton optOtherInfo 
         Caption         =   "Other Info"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "View Other Information"
         Top             =   5070
         Width           =   5655
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00DEDFDE&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Information"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   5655
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   -30
         X2              =   5805
         Y1              =   390
         Y2              =   405
      End
   End
   Begin wizButton.cmd cmdOtherInfo 
      Height          =   6810
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   12012
      TX              =   "Other Info"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "OTHERInfoMain.frx":067F
   End
End
Attribute VB_Name = "frmOTHERInfoMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExitOtherInfo_Click()
    If MsgBox("Exit Viewing Other Info?", vbQuestion + vbYesNo, "Exit Window") = vbYes Then
        frmHRMSEmpInfo.optViewOtherInfo.Value = False
        Unload Me
    End If
End Sub

Private Sub cmdShowOtherInfo_Click()
    On Error GoTo Errorcode
    If optDependents.Value = True Then
        Screen.MousePointer = 11
        Unload frmOTHERINFODependents
        Screen.MousePointer = 0
        frmOTHERINFODependents.Show vbModal
    End If
    If optParents.Value = True Then
        Screen.MousePointer = 11
        Unload frmOTHERINFOParents
        Screen.MousePointer = 0
        frmOTHERINFOParents.Show vbModal
    End If
    If optEducation.Value = True Then
        Screen.MousePointer = 11
        Unload frmOTHERINFOEducation
        Screen.MousePointer = 0
        frmOTHERINFOEducation.Show vbModal
    End If
    If optCSEligibility.Value = True Then
        Screen.MousePointer = 11
        Unload frmOTHERINFOCSEligibility
        Screen.MousePointer = 0
        frmOTHERINFOCSEligibility.Show vbModal
    End If
    If optExamsPassed.Value = True Then
        Screen.MousePointer = 11
        Unload frmOTHERINFOExamsPassed
        Screen.MousePointer = 0
        frmOTHERINFOExamsPassed.Show vbModal
    End If
    If optPastEmployment.Value = True Then
        Screen.MousePointer = 11
        Unload frmOTHERINFOPastEmployment
        Screen.MousePointer = 0
        frmOTHERINFOPastEmployment.Show vbModal
    End If
    If optTraining.Value = True Then
        Screen.MousePointer = 11
        Unload frmOTHERINFOTraining
        Screen.MousePointer = 0
        frmOTHERINFOTraining.Show vbModal
    End If
    If optTrainingPlan.Value = True Then
        Screen.MousePointer = 11
        Unload frmOTHERINFOTrainings
        Screen.MousePointer = 0
        frmOTHERINFOTrainings.Show vbModal
    End If
    If optPerformanceEvaluation.Value = True Then
        Screen.MousePointer = 11
        Unload frmOTHERINFOPersonalAction
        Screen.MousePointer = 0
        frmOTHERINFOPerformanceEvaluation.Show vbModal
    End If
    If optPersonnelAction.Value = True Then
        Screen.MousePointer = 11
        Unload frmOTHERINFOPersonalAction
        Screen.MousePointer = 0
        frmOTHERINFOPersonalAction.Show vbModal
    End If
    If optMemo.Value = True Then
        Screen.MousePointer = 11
        Unload frmOTHERINFOMemorandum
        Screen.MousePointer = 0
        frmOTHERINFOMemorandum.Show vbModal
    End If
    If optOtherInfo.Value = True Then
        'Screen.MousePointer = 11
        'Unload frmOTHERINFOotherinfo
        'Screen.MousePointer = 0
        'frmOTHERINFOotherinfo.Show vbModal
    End If
    If optJoannPearl.Value = True Then
        PrintSQLReport rptJoannPearl, HRMS_REPORT_PATH & "\IndivDevPlan.rpt", "{EmpInfo.EmpNo} = " & EMPLOYEE_NO, DMIS_REPORT_Connection, 1
    End If
    Exit Sub

Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 0
End Sub


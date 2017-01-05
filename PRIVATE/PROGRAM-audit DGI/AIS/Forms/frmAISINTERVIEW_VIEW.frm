VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAISINTERVIEW_VIEW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Interview Schedule"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAISINTERVIEW_VIEW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8610
   Begin VB.PictureBox picUPDATE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3885
      Left            =   1560
      ScaleHeight     =   3855
      ScaleWidth      =   5535
      TabIndex        =   20
      Top             =   1290
      Visible         =   0   'False
      Width           =   5565
      Begin VB.TextBox txtNOTES 
         Appearance      =   0  'Flat
         Height          =   1365
         Left            =   1590
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1530
         Width           =   3765
      End
      Begin VB.ComboBox cboREMARKS 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "frmAISINTERVIEW_VIEW.frx":058A
         Left            =   1590
         List            =   "frmAISINTERVIEW_VIEW.frx":0597
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   3825
      End
      Begin VB.CommandButton cmdCANCEL 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4500
         Picture         =   "frmAISINTERVIEW_VIEW.frx":05B3
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancel"
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton cmdUPDATE 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3660
         Picture         =   "frmAISINTERVIEW_VIEW.frx":0B2F
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Update Changes"
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label lblAPP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   1590
         TabIndex        =   26
         Top             =   780
         Width           =   3795
      End
      Begin VB.Label lblAPP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1590
         TabIndex        =   25
         Top             =   420
         Width           =   1785
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   1020
         TabIndex        =   24
         Top             =   1530
         Width           =   480
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   660
         TabIndex        =   23
         Top             =   870
         Width           =   840
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   750
         TabIndex        =   22
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applicant ID"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   420
         TabIndex        =   21
         Top             =   510
         Width           =   1050
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   405
         Left            =   0
         TabIndex        =   27
         Top             =   -180
         Width           =   5715
         _Version        =   655364
         _ExtentX        =   10081
         _ExtentY        =   714
         _StockProps     =   14
         Caption         =   "       "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.99
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         ForeColor       =   0
      End
   End
   Begin VB.Frame fmeAPP 
      Caption         =   "List Of Applicants"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   90
      TabIndex        =   7
      Top             =   2070
      Width           =   8445
      Begin MSComctlLib.ListView lsvDISP 
         Height          =   2175
         Left            =   120
         TabIndex        =   0
         Top             =   330
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   3836
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
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Remarks"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Note"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Double Click To Edit Result"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   2610
         Width           =   2475
      End
   End
   Begin VB.CommandButton cmdEXIT 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   7710
      Picture         =   "frmAISINTERVIEW_VIEW.frx":11CF
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit Window"
      Top             =   5100
      Width           =   795
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6930
      Picture         =   "frmAISINTERVIEW_VIEW.frx":1721
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Schedule of Interview"
      Top             =   5100
      Width           =   795
   End
   Begin Crystal.CrystalReport rptDISPLAY 
      Left            =   390
      Top             =   5250
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "Schedule"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   90
      TabIndex        =   9
      Top             =   60
      Width           =   8415
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   2040
         TabIndex        =   19
         Top             =   1470
         Width           =   2115
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Interview"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   16
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1500
         TabIndex        =   15
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1710
         TabIndex        =   14
         Top             =   1590
         Width           =   210
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   2085
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   2040
         TabIndex        =   12
         Top             =   660
         Width           =   6285
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interview Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   1830
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   2040
         TabIndex        =   10
         Top             =   1050
         Width           =   2115
      End
   End
   Begin VB.Label lblSCHED_ID 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   9000
      TabIndex        =   18
      Top             =   150
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblEXAMID 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   9000
      TabIndex        =   17
      Top             =   1590
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "frmAISINTERVIEW_VIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function DisplayList(SCHED_ID As Integer)
    Dim rsTmp                                                         As ADODB.Recordset
    Dim ITEM                                                          As ListItem

    lsvDISP.Enabled = False

    Set rsTmp = gconDMIS.Execute("Select * From HRMS_APPLICANT_INTERVIEW_SCHEDULE Where INT_ID = " & SCHED_ID & "")
    lsvDISP.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvDISP.ListItems.Add(, , rsTmp!APPLICANT_ID)
            ITEM.SubItems(1) = FindApplicantName(rsTmp!APPLICANT_ID)
            ITEM.SubItems(2) = Null2String(rsTmp!REMARKS)
            ITEM.SubItems(3) = Null2String(rsTmp!NOTES)

            rsTmp.MoveNext
        Loop
    End If
    lsvDISP.Enabled = True
End Function

Sub CleanUpdateForm()
    txtNotes.Text = ""
    lblAPP(0).Caption = ""
    lblAPP(1).Caption = ""
End Sub

Sub ShowUpdate(COND As Boolean)
    cmdPrint.Enabled = COND
    cmdExit.Enabled = COND
    fmeAPP.Enabled = COND
    picUPDATE.Visible = Not COND
End Sub

Private Sub cmdCancel_Click()
    Call ShowUpdate(True)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:18
Private Sub cmdPRINT_Click()
    Dim FILTER                                                        As String
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "ACESS_PRINT", "APPLICANT INFO") = False Then Exit Sub
    rptDISPLAY.Formulas(0) = "DateOfExam = '" & lblINFO(0).Caption & "'"
    rptDISPLAY.Formulas(1) = "InterviewDescription = '" & lblINFO(1).Caption & "'"
    rptDISPLAY.Formulas(2) = "FTime = '" & lblINFO(2).Caption & "'"
    rptDISPLAY.Formulas(3) = "TTime = '" & lblINFO(3).Caption & "'"



    'FILTER = "{HRMS_EXAM_SCHEDULE.SCHED_ID} = " & CLng(lblSCHED_ID.Caption) & ""
    'Call PrintSQLReport(rptDISPLAY, AIS_REPORT_PATH & "ExamSchedule.rpt", FILTER, AIS_REPORT_Connection, 1)

    FILTER = "{HRMS_INTERVIEW_SCHEDULE.INT_ID} = " & CLng(lblSCHED_ID.Caption) & ""
    Call PrintSQLReport(rptDISPLAY, AIS_REPORT_PATH & "InterviewSchedule.rpt", FILTER, AIS_REPORT_Connection, 1)

    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:18
Private Sub cmdUPDATE_Click()
    Dim REMARKS                                                       As String
    Dim NOTES                                                         As String

    On Error GoTo Errorcode:
    frmMain.MousePointer = 11

    If MsgBox("Update Applicant Interview", vbQuestion + vbYesNo + vbDefaultButton1, "Are You Sure") = vbYes Then
        REMARKS = N2Str2Null(cboREMARKS)
        NOTES = N2Str2Null(txtNotes)

        gconDMIS.Execute ("Update HRMS_APPLICANT_INTERVIEW_SCHEDULE Set Remarks = " & REMARKS & _
                          ",Notes = " & NOTES & " Where Applicant_ID = " & lblAPP(0).Caption & _
                        " And INT_ID = " & CLng(lblSCHED_ID.Caption) & "")

        Call ShowUpdate(True)

        Call DisplayList(lblSCHED_ID.Caption)
    End If
    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISINTERVIEW.Enabled = True
    On Error Resume Next
    frmAISINTERVIEW.SetFocus
End Sub

Private Sub lsvDISP_DblClick()
    Dim INDEX                                                         As Long

    If Not lsvDISP.ListItems.count = 0 Then
        INDEX = lsvDISP.SelectedItem.INDEX
        Call CleanUpdateForm

        With lsvDISP
            Call ShowUpdate(False)
            lblAPP(0).Caption = .ListItems(INDEX).Text
            lblAPP(1).Caption = .ListItems(INDEX).SubItems(1)

            cboREMARKS.ListIndex = 0
            On Error Resume Next
            cboREMARKS.SetFocus
        End With
    End If
End Sub


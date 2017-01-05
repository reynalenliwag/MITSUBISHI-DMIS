VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Report_TestDriveEvaluation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Drive Evaluation Report"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_TestDriveEvalution.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2445
   ScaleWidth      =   4905
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
      Height          =   825
      Left            =   2460
      MouseIcon       =   "Report_TestDriveEvalution.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "Report_TestDriveEvalution.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Close Window"
      Top             =   1560
      Width           =   885
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Monthly Evaluation"
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
      Left            =   2670
      TabIndex        =   10
      Top             =   1230
      Width           =   1965
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Select Date Ranged"
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
      Left            =   570
      TabIndex        =   9
      Top             =   1230
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   990
      Top             =   1950
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "MMPC Monthly Inventory Control"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
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
      Height          =   825
      Left            =   1590
      MouseIcon       =   "Report_TestDriveEvalution.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "Report_TestDriveEvalution.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Print Report"
      Top             =   1560
      Width           =   885
   End
   Begin VB.PictureBox picMonthly 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   90
      ScaleHeight     =   1215
      ScaleWidth      =   5085
      TabIndex        =   5
      Top             =   0
      Width           =   5085
      Begin VB.TextBox txtYear 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   555
         Left            =   3510
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "9999"
         Top             =   210
         Width           =   1005
      End
      Begin VB.ComboBox cboMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   465
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2670
         TabIndex        =   8
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.PictureBox picRange 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   -90
      Width           =   4995
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   405
         Left            =   1470
         TabIndex        =   1
         Top             =   270
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   52953089
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   405
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   52953089
         CurrentDate     =   38216
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   690
         TabIndex        =   4
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   810
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmSMIS_Report_TestDriveEvaluation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim FILTER                                                        As String
    On Error GoTo ErrorCode:

    If picMonthly.Visible = True Then
        If Len(txtYear.Text) = 4 Or txtYear.Text <> "" Then
            If gconDMIS.Execute("select COUNT(*) from CRIS_TestDriveSchedules WHERE year(ENDDATETIME) = '" & txtYear.Text & "' and month(ENDDATETIME) = " & What_month(cboMonth)).Fields(0).Value > 0 Then
                Screen.MousePointer = 11
                rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptReleased.Formulas(2) = "Manth = ' As of " & cboMonth & " of " & txtYear & "'"
                FILTER = "(year({CTD.ENDDATETIME}) = " & txtYear.Text & " AND month({CTD.ENDDATETIME}) = " & What_month(cboMonth) & ")"
                rptReleased.WindowTitle = "Test Drive Evaluation Report"
                PrintSQLReport rptReleased, SMIS_REPORT_PATH & "TDEval.rpt", FILTER, DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "No Record for " & cboMonth.Text & " " & txtYear.Text
                Exit Sub
            End If
        End If
    Else

        If gconDMIS.Execute("select COUNT(*) from CRIS_TestDriveSchedules WHERE ENDDATETIME between '" & dtpFrom.Value & "' and '" & dtpTo.Value & "'").Fields(0).Value > 0 Then

            Screen.MousePointer = 11
            rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptReleased.Formulas(2) = "Manth = ' For The Period from :" & FormatDateTime(dtpFrom, vbShortDate) & " To " & FormatDateTime(dtpTo, vbShortDate) & "'"

            FILTER = "((year({CTD.ENDDATETIME}) >=" & Year(dtpFrom) & " AND month({CTD.ENDDATETIME}) >= " & Month(dtpFrom) & " AND Day({CTD.ENDDATETIME}) >= " & Day(dtpFrom) & ")"
            FILTER = FILTER & " AND " & " (year({CTD.ENDDATETIME}) <= " & Year(dtpTo) & " AND month({CTD.ENDDATETIME}) <= " & Month(dtpTo) & " AND Day({CTD.ENDDATETIME}) <= " & Day(dtpTo) & " ))"


            rptReleased.WindowTitle = "Test Drive Evaluation Report"
            PrintSQLReport rptReleased, SMIS_REPORT_PATH & "TDEval.rpt", FILTER, DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record Found In The Range "
            Exit Sub
        End If
    End If


    If Option1.Value = True Then
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "TEST DRIVE EVALUATION", "", "", "", "FROM " & dtpFrom & " " & "TO " & dtpTo, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Else
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
          Call NEW_LogAudit("V", "TEST DRIVE EVALUATION", "", "", "", cboMonth & " " & txtYear, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    End If

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (TEST DRIVE EVALUATION)"
            Call frmALL_AuditInquiry.DisplayHistory("", "TEST DRIVE EVALUATION", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))
    dtpFrom.Value = firstDay(LOGDATE)
    dtpTo.Value = LOGDATE
    txtYear.Text = Year(LOGDATE)
    Option2.Value = True
    Screen.MousePointer = 0
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        picMonthly.Visible = False
        picRange.Visible = True
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        picMonthly.Visible = True
        picRange.Visible = False
    End If

End Sub

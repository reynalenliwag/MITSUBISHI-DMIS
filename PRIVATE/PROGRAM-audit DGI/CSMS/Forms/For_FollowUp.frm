VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSFor_followUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "For Follow-Up"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   ForeColor       =   &H00DEDFDE&
   Icon            =   "For_FollowUp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   3735
   Begin MSComCtl2.DTPicker txtDate 
      Height          =   375
      Left            =   1380
      TabIndex        =   8
      Top             =   90
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52690945
      CurrentDate     =   39622
   End
   Begin VB.TextBox txtNDays 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2940
      TabIndex        =   3
      ToolTipText     =   "Input starting date of the report"
      Top             =   1050
      Width           =   645
   End
   Begin VB.OptionButton Option3 
      Caption         =   "INPUT NUMBER OF DAYS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   270
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.OptionButton Option2 
      Caption         =   "AFTER 3 DAYS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   1
      Top             =   810
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "AFTER 3 MONTHS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   510
      Value           =   -1  'True
      Width           =   3135
   End
   Begin Crystal.CrystalReport rptTechnician 
      Left            =   90
      Top             =   1890
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Technician Efficiency Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
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
      Height          =   825
      Left            =   2820
      MouseIcon       =   "For_FollowUp.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "For_FollowUp.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close Window"
      Top             =   1500
      Width           =   795
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
      Left            =   2040
      MouseIcon       =   "For_FollowUp.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "For_FollowUp.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   1500
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   30
      TabIndex        =   5
      Top             =   210
      Width           =   1245
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2100
      TabIndex        =   4
      Top             =   3120
      Width           =   495
   End
End
Attribute VB_Name = "frmCSMSFor_followUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRO_DET                                           As ADODB.Recordset
Dim rsREPOR                                            As ADODB.Recordset

Sub PrintFollowupInExcel()
    Dim rsREPOR                                        As New ADODB.Recordset

    Set rsREPOR = gconDMIS.Execute("")
    Set rsREPOR = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "TRANSACTIONS FOR FOLLOW UP") = False Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo ERRORCODE
    Dim Selected_Date                                  As String


    '    If MsgBox("Print in Excel", vbQuestion + vbYesNo, "CSMS") = vbYes Then
    '
    '        Exit Sub
    '    End If

    If Option1.Value = True Then
        Selected_Date = DateSerial(Year(txtDate.Value), Month(txtDate.Value) - 3, Day(txtDate.Value))
    ElseIf Option2.Value = True Then
        Selected_Date = DateSerial(Year(txtDate.Value), Month(txtDate.Value), Day(txtDate.Value) - 3)
    Else
        Selected_Date = DateSerial(Year(txtDate.Value), Month(txtDate.Value), Day(txtDate.Value) - NumericVal(txtNDays.Text))
    End If

    Set rsREPOR = New ADODB.Recordset
    Set rsREPOR = gconDMIS.Execute("Select * from CSMS_RepOr Where Dte_rel = '" & Selected_Date & "'")
    If Not rsREPOR.EOF And Not rsREPOR.BOF Then
        Screen.MousePointer = 11
        If Option1.Value = True Then
            rptTechnician.ReportTitle = "CUSTOMER TRANSACTIONS FOR FOLLOW-UP (AFTER 3 MONTHS)"
            rptTechnician.WindowTitle = "CUSTOMER TRANSACTIONS FOR FOLLOW-UP (AFTER 3 MONTHS)"
        ElseIf Option2.Value = True Then
            rptTechnician.ReportTitle = "CUSTOMER TRANSACTIONS FOR FOLLOW-UP (AFTER 3 DAYS)"
            rptTechnician.WindowTitle = "CUSTOMER TRANSACTIONS FOR FOLLOW-UP (AFTER 3 DAYS)"
        Else
            rptTechnician.ReportTitle = "CUSTOMER TRANSACTIONS FOR FOLLOW-UP (AFTER " & NumericVal(txtNDays.Text) & " DAYS)"
            rptTechnician.WindowTitle = "CUSTOMER TRANSACTIONS FOR FOLLOW-UP (AFTER " & NumericVal(txtNDays.Text) & " DAYS)"
        End If

        'JUN 02/05/2008
        rptTechnician.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
        rptTechnician.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
        rptTechnician.Formulas(2) = "Printedby = '" & LOGNAME & "'"

        PrintSQLReport rptTechnician, CSMS_REPORT_PATH & "For_FollowUp.rpt", "{REPOR.Dte_Rel} = date(" & Year(Selected_Date) & "," & Month(Selected_Date) & "," & Day(Selected_Date) & ")", CSMS_REPORT_CONNECTION, 1
        'LogAudit "V", "FOR FOLLOW UP REPORT", txtDate

        'NEW LOG AUDIT-----------------------------------------------------
        Dim rTYPE                                      As String
        If Option1.Value = True Then rTYPE = "After 3 Months"
        If Option2.Value = True Then rTYPE = "After 3 Days"
        If Option3.Value = True Then rTYPE = "After " & txtNDays & " Days"
        Call NEW_LogAudit("V", "TRANSACTIONS FOR FOLLOW UP", "", "", "", "Choosen Date: " & txtDate & "-" & rTYPE, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        Screen.MousePointer = 0
    Else
        ShowNoRecord
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0

    Exit Sub

ERRORCODE:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (TRANSACTIONS FOR FOLLOW UP)"
            Call frmALL_AuditInquiry.DisplayHistory("", "TRANSACTIONS FOR FOLLOW UP", "PRINTING")
            'End If

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtDate.Value = LOGDATE
    txtNDays.Text = 0
    txtNDays.Enabled = False
    Screen.MousePointer = 0
End Sub

Private Sub Option1_Click()
    If Option3.Value = True Then
        txtNDays.Enabled = True
    Else
        txtNDays.Enabled = False
    End If
End Sub

Private Sub Option2_Click()
    If Option3.Value = True Then
        txtNDays.Enabled = True
    Else
        txtNDays.Enabled = False
    End If
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then
        txtNDays.Enabled = True
    Else
        txtNDays.Enabled = False
    End If
End Sub

Private Sub txtdate_LostFocus()
    '    If IsDate(txtDate) = False Then
    '        txtDate.Text = Format(txtDate.Text, "SHORT DATE")
    '    Else
    '        txtDate.Text = Date
    '    End If
End Sub

Private Sub txtNDays_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub


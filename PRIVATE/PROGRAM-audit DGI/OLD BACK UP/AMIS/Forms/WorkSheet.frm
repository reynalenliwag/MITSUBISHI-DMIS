VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISWorkSheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Work Sheet"
   ClientHeight    =   3555
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4620
   ForeColor       =   &H00FFFFFF&
   Icon            =   "WorkSheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4620
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
      Height          =   885
      Left            =   2310
      MouseIcon       =   "WorkSheet.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "WorkSheet.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Close Window"
      Top             =   2640
      Width           =   885
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
      Height          =   885
      Left            =   1440
      MouseIcon       =   "WorkSheet.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "WorkSheet.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Print Report"
      Top             =   2640
      Width           =   885
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Work Sheet for the Period"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   285
      Left            =   930
      TabIndex        =   5
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Frame picPeriod 
      Height          =   585
      Left            =   210
      TabIndex        =   6
      Top             =   1950
      Width           =   4215
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   780
         TabIndex        =   8
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   52101121
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   2730
         TabIndex        =   10
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   52101121
         CurrentDate     =   38216
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   210
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   2220
         TabIndex        =   9
         Top             =   210
         Width           =   435
      End
   End
   Begin VB.CheckBox chkInclude 
      Caption         =   "Include Trial Balance of Journals"
      Enabled         =   0   'False
      ForeColor       =   &H00701E2A&
      Height          =   525
      Left            =   1170
      TabIndex        =   4
      Top             =   1125
      Width           =   2535
   End
   Begin VB.OptionButton optBalances 
      Caption         =   "Work Sheet Of Balances"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   930
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.OptionButton optTotals 
      Caption         =   "Work Sheet Of Totals"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   930
      TabIndex        =   3
      Top             =   870
      Width           =   2775
   End
   Begin Crystal.CrystalReport rptAMISTrialBalance 
      Left            =   300
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Work Sheet - Trial Balance of Journals"
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
   Begin MSComCtl2.DTPicker dtpAsOF 
      Height          =   405
      Left            =   1440
      TabIndex        =   1
      Top             =   90
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52101121
      CurrentDate     =   38216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "As Of:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   630
      TabIndex        =   0
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frmAMISWorkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                  As ADODB.Recordset

Function Setacctname(XXX As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount where AcctCode = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        Setacctname = Null2String(rsChartAccount!Description)
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200714:33
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:


    Dim DEBIT_BALANCE                             As Double
    Dim CREDIT_BALANCE                            As Double
    Dim TOTAL_DEBIT_BALANCE                       As Double
    Dim TOTAL_CREDIT_BALANCE                      As Double
    If IsDate(dtpAsOF) = False Then
        MsgSpeechBox "Error In As of date"
        Exit Sub
    End If
    If Option1.Value = True Then
        Set rsJournal_HD = New ADODB.Recordset
        rsJournal_HD.Open "select * from AMIS_Journal_Det where jtype <> 'CLO' and (jdate >= '" & CDate(dtpFrom) & "' AND jdate <= '" & CDate(dtpTo) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
            Screen.MousePointer = 11
            Set rsJournal_HD = New ADODB.Recordset
            rsJournal_HD.Open "select SUM(DEBIT) AS DEBIT_TOTAL, SUM(CREDIT) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype <> 'CLO' and Status = 'P' AND (jdate >= '" & CDate(dtpFrom) & "' AND jdate <= '" & CDate(dtpTo) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
                rsJournal_HD.MoveFirst
                TOTAL_DEBIT_BALANCE = 0
                TOTAL_CREDIT_BALANCE = 0
                gconDMIS.Execute "update AMIS_ChartAccount Set Debit_Total = 0,Credit_Total = 0,DebitBalance = 0,CreditBalance = 0"
                Screen.MousePointer = 11
                Do While Not rsJournal_HD.EOF
                    If N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) > N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) Then
                        DEBIT_BALANCE = N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) - N2Str2Zero(rsJournal_HD!CREDIT_TOTAL)
                        CREDIT_BALANCE = 0
                    Else
                        If N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) > N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) Then
                            CREDIT_BALANCE = N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) - N2Str2Zero(rsJournal_HD!DEBIT_TOTAL)
                            DEBIT_BALANCE = 0
                        Else
                            CREDIT_BALANCE = 0: DEBIT_BALANCE = 0
                        End If
                    End If
                    TOTAL_DEBIT_BALANCE = TOTAL_DEBIT_BALANCE + DEBIT_BALANCE
                    TOTAL_CREDIT_BALANCE = TOTAL_CREDIT_BALANCE + CREDIT_BALANCE
                    gconDMIS.Execute "update AMIS_ChartAccount Set" & _
                                     " Debit_Total = " & N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) & "," & _
                                     " Credit_Total = " & N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) & "," & _
                                     " DebitBalance = " & DEBIT_BALANCE & "," & _
                                     " CreditBalance = " & CREDIT_BALANCE & _
                                     " Where AcctCode = '" & Null2String(rsJournal_HD!Acct_code) & "'"
                    rsJournal_HD.MoveNext
                Loop
                Screen.MousePointer = 0
            End If
            Screen.MousePointer = 0
            ShowWorkSheetReport "WorkSheetOfBalances", "FinancialStatement", "", "Work Sheet of Balances", "From: " & dtpFrom & " To: " & dtpTo, True, TOTAL_DEBIT_BALANCE, TOTAL_CREDIT_BALANCE
            'ShowReport "TrialBalanceForWorkSheets", "FinancialStatement", "{AMIS_Journal_Det.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {AMIS_Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Work Sheet - Trial Balance Of Journals", "From: " & dtpFrom & " To: " & dtpTo, True
            ShowReport "TrialBalanceForWorkSheets", "FinancialStatement", "{Journal_Det.jdate} >= CDate('" & dtpFrom & "') AND {Journal_Det.jdate} <= CDate('" & dtpTo & "')", "Work Sheet - Trial Balance Of Journals", "From: " & dtpFrom & " To: " & dtpTo, True
            'Unload Me
        Else
            ShowNoRecord
        End If
        Call NEW_LogAudit("V", "WorkSheet of Balance", "", "", "", dtpFrom & " " & dtpFrom, "", "")
    Else
        Set rsJournal_HD = New ADODB.Recordset
        'rsJournal_Hd.Open "select * from AMIS_Journal_Det where (jdate <= '" & CDate(dtpAsOF) & "') and year(jdate) = " & Year(dtpTo), gconDMIS, adOpenForwardOnly, adLockReadOnly
        rsJournal_HD.Open "select * from AMIS_Journal_Det where jtype <> 'CLO' and (jdate <= '" & CDate(dtpAsOF) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
            If optBalances.Value = True Then
                Screen.MousePointer = 11
                Set rsJournal_HD = New ADODB.Recordset
                'rsJournal_Hd.Open "select SUM(DEBIT) AS DEBIT_TOTAL, SUM(CREDIT) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "') and year(jdate) = " & Year(dtpTo) & " group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                rsJournal_HD.Open "select SUM(DEBIT) AS DEBIT_TOTAL, SUM(CREDIT) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype <> 'CLO' and Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
                    rsJournal_HD.MoveFirst
                    TOTAL_DEBIT_BALANCE = 0
                    TOTAL_CREDIT_BALANCE = 0
                    gconDMIS.Execute "update AMIS_ChartAccount Set Debit_Total = 0,Credit_Total = 0,DebitBalance = 0,CreditBalance = 0"
                    Do While Not rsJournal_HD.EOF
                        'MsgBox "CODE : " & Null2String(rsJournal_HD!ACCT_CODE) & vbCrLf & _
                         '       "NAME : " & SetAcctName(Null2String(rsJournal_HD!ACCT_CODE)) & vbCrLf & _
                         '       "DEBIT: " & N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) & " CREDIT: " & N2Str2Zero(rsJournal_HD!CREDIT_TOTAL)
                        If N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) > N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) Then
                            DEBIT_BALANCE = N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) - N2Str2Zero(rsJournal_HD!CREDIT_TOTAL)
                            CREDIT_BALANCE = 0
                        Else
                            If N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) > N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) Then
                                CREDIT_BALANCE = N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) - N2Str2Zero(rsJournal_HD!DEBIT_TOTAL)
                                DEBIT_BALANCE = 0
                            Else
                                CREDIT_BALANCE = 0: DEBIT_BALANCE = 0
                            End If
                        End If
                        TOTAL_DEBIT_BALANCE = TOTAL_DEBIT_BALANCE + DEBIT_BALANCE
                        TOTAL_CREDIT_BALANCE = TOTAL_CREDIT_BALANCE + CREDIT_BALANCE
                        gconDMIS.Execute "update AMIS_ChartAccount Set" & _
                                         " Debit_Total = " & N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) & "," & _
                                         " Credit_Total = " & N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) & "," & _
                                         " DebitBalance = " & DEBIT_BALANCE & "," & _
                                         " CreditBalance = " & CREDIT_BALANCE & _
                                         " Where AcctCode = '" & Null2String(rsJournal_HD!Acct_code) & "'"
                        rsJournal_HD.MoveNext
                    Loop
                End If
                Screen.MousePointer = 0
                ShowWorkSheetReport "WorkSheetOfBalances", "FinancialStatement", "", "Work Sheet of Balances", "AS OF: " & dtpAsOF, True, TOTAL_DEBIT_BALANCE, TOTAL_CREDIT_BALANCE
            Else
                ShowReport "WorkSheetOfTotals", "FinancialStatement", "{Journal_Det.jdate} <= date(" & Year(dtpAsOF) & "," & Month(dtpAsOF) & "," & Day(dtpAsOF) & ")", "Work Sheet of Totals", "AS OF: " & dtpAsOF, True
                If chkInclude.Value = 1 Then ShowReport "TrialBalanceForWorkSheets", "FinancialStatement", "{Journal_Det.jdate} <= date(" & Year(dtpAsOF) & "," & Month(dtpAsOF) & "," & Day(dtpAsOF) & ")", "Work Sheet - Trial Balance Of Journals", "AS OF: " & dtpAsOF, True
                Call NEW_LogAudit("V", "WorkSheet of total", "", "", "", dtpFrom & " " & dtpFrom, "", "")
            End If
            'Unload Me
        Else
            ShowNoRecord
        End If

    End If
    Exit Sub
Errorcode:
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
        If Option1.Value = True Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (WorkSheet of Balance)"
            Call frmALL_AuditInquiry.DisplayHistory("", "WorkSheet of Balance", "PRINTING")
        Else
            frmALL_AuditInquiry.Caption = "Audit Inquiry (WorkSheet of total)"
            Call frmALL_AuditInquiry.DisplayHistory("", "WorkSheet of total", "PRINTING")
        End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    dtpAsOF = LOGDATE
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
    Screen.MousePointer = 0
End Sub


Private Sub optTotals_Click()
    If optTotals.Value = True Then
        chkInclude.Enabled = True
    Else
        chkInclude.Value = 0
        chkInclude.Enabled = False
    End If
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
End Sub

Private Sub optBalances_Click()
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        picPeriod.Enabled = True
        dtpAsOF.Enabled = False
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
    Else
        picPeriod.Enabled = False
        dtpAsOF.Enabled = True
        dtpFrom.Enabled = False
        dtpTo.Enabled = False
    End If
End Sub

Public Sub ShowWorkSheetReport(ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean, SUM_DEBIT_WS As Variant, SUM_CREDIT_WS As Variant)
    Screen.MousePointer = 11
    Dim rsProfile                                 As ADODB.Recordset
    Dim CrystalRpt                                As Crystal.CrystalReport
    Set CrystalRpt = Me.rptAMISTrialBalance
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rtp"
        CrystalRpt.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        CrystalRpt.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        If WithDate = True Then CrystalRpt.Formulas(2) = "ReportDate = '" & REPORT_DATE & "'"
        CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
        CrystalRpt.Formulas(3) = "SUM_DEBIT =" & SUM_DEBIT_WS
        CrystalRpt.Formulas(4) = "SUM_CREDIT =" & SUM_CREDIT_WS
        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
        CrystalRpt.PageZoom 89
    End If
    Screen.MousePointer = 0
End Sub


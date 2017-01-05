VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISExpenseReport 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1680
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   4830
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ExpenseReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1680
   ScaleWidth      =   4830
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
      Left            =   2535
      MouseIcon       =   "ExpenseReport.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ExpenseReport.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   735
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
      Height          =   825
      Left            =   1665
      MouseIcon       =   "ExpenseReport.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ExpenseReport.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   735
      Width           =   885
   End
   Begin Crystal.CrystalReport rptAMISIncomeStatement 
      Left            =   900
      Top             =   1020
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Income Statements"
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   405
      Left            =   780
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
      Format          =   52232193
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3030
      TabIndex        =   3
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
      Format          =   52232193
      CurrentDate     =   38216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   2550
      TabIndex        =   2
      Top             =   150
      Width           =   435
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   6
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISExpenseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_Det                                 As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200713:02
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:


    Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo    As String
    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
    End If
    Dim vCUMULATIVE_GROSSSALES, vCUMULATIVE_DISCOUNTSRETURNS, vCUMULATIVE_NETSALES As Double
    Dim vCURRENT_GROSSSALES, vCURRENT_DISCOUNTSRETURNS, vCURRENT_NETSALES As Double
    Dim vPREVIOUS_GROSSSALES, vPREVIOUS_DISCOUNTSRETURNS, vPREVIOUS_NETSALES As Double

    Dim rsJournal_Det                             As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') and year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND (AMIS_ChartAccount.Headers=" & CASH_SALES & " OR AMIS_ChartAccount.Headers=" & CHARGE_SALES & ")")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCUMULATIVE_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') and YEAR(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vCUMULATIVE_NETSALES = vCUMULATIVE_GROSSSALES - vCUMULATIVE_DISCOUNTSRETURNS
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers=" & CASH_SALES & " OR AMIS_ChartAccount.Headers=" & CHARGE_SALES & ")")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vCURRENT_NETSALES = vCURRENT_GROSSSALES - vCURRENT_DISCOUNTSRETURNS
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (AMIS_ChartAccount.Headers=" & CASH_SALES & " OR AMIS_ChartAccount.Headers=" & CHARGE_SALES & ")")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPREVIOUS_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPREVIOUS_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vPREVIOUS_NETSALES = vPREVIOUS_GROSSSALES - vPREVIOUS_DISCOUNTSRETURNS
    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.Formulas(5) = "ToJDate = '" & CDate(dtpTo) & "'"
        If REPORT_EXPENSETYPE = "ADMIN" Then
            rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF ADMINISTRATIVE EXPENSES"
            rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF ADMINISTRATIVE EXPENSES"
        Else
            rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES"
            rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES"
        End If
    End If
    rptAMISIncomeStatement.Formulas(0) = "CUMULATIVE_NETSALES = " & NumericVal(vCUMULATIVE_NETSALES)
    rptAMISIncomeStatement.Formulas(1) = "CURRENT_NETSALES = " & NumericVal(vCURRENT_NETSALES)
    rptAMISIncomeStatement.Formulas(2) = "PREVIOUS_NETSALES = " & NumericVal(vPREVIOUS_NETSALES)
    ReportFolder = "ExpenseReport\"
    If REPORT_EXPENSETYPE = "ADMIN" Then
        PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfAdministrativeExpensesCumulative.rpt", "({Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")) and year({Journal_Hd.jdate}) = " & Year(dtpTo), DMIS_REPORT_Connection, 1
        LogAudit "V", "SCHEDULE OF ADMINISTRATIVE EXPENSES CUMULATIVE"
    Else
        PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfSellingExpensesCumulative.rpt", "({Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")) and year({Journal_Hd.jdate}) = " & Year(dtpTo), DMIS_REPORT_Connection, 1
        LogAudit "V", "SCHEDULE OF SELLING EXPENSES CUMULATIVE"
    End If

    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    If REPORT_EXPENSETYPE = "ADMIN" Then
        Me.Caption = "SCHEDULE OF ADMINSTRATIVE EXPENSE"
    Else
        Me.Caption = "SCHEDULE OF SELLING EXPENSE"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub


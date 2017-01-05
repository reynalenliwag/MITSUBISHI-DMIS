VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISBalanceSheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Sheet"
   ClientHeight    =   1575
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   3135
   ForeColor       =   &H00F5F5F5&
   Icon            =   "BalanceSheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   3135
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
      Left            =   1680
      MouseIcon       =   "BalanceSheet.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "BalanceSheet.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   675
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
      Left            =   810
      MouseIcon       =   "BalanceSheet.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "BalanceSheet.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   675
      Width           =   885
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   1050
      TabIndex        =   1
      Top             =   120
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
      Format          =   51773441
      CurrentDate     =   38216
   End
   Begin Crystal.CrystalReport rptAMISBalanceSheet 
      Left            =   60
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Financial Statements - Balance Sheet"
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   405
      Left            =   2820
      TabIndex        =   4
      Top             =   1740
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
      Format          =   51773441
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   765
   End
End
Attribute VB_Name = "frmAMISBalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                  As ADODB.Recordset
Dim rsJournal_Det                                 As ADODB.Recordset

Dim V_NetIncomeOrLoss                             As Double
Dim V_ProvisionForBonus                           As Double
Dim V_ProvisionForTax                             As Double
Dim V_Total_Current_Asset                         As Double
Dim V_Net_Propert_Equipment                       As Double
Dim V_Other_Assets                                As Double
Dim V_Propert_Equipment                           As Double
Dim V_AccumDepreciation                           As Double
Dim V_TaxCredit                                   As Double

Sub ShowBalanceSheetReport(ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean)
    Screen.MousePointer = 11
    Dim rsProfile                                 As ADODB.Recordset
    Dim CrystalRpt                                As Crystal.CrystalReport
    Set CrystalRpt = frmMain.rptMain
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        'CrystalRpt.Reset
        CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt"
        CrystalRpt.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        CrystalRpt.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        If WithDate = True Then CrystalRpt.Formulas(2) = "ReportDate = '" & REPORT_DATE & "'"
        CrystalRpt.Formulas(3) = "NetIncomeOrLoss = " & V_NetIncomeOrLoss
        CrystalRpt.Formulas(4) = "NetIncomeOrLoss2 = " & V_NetIncomeOrLoss
        CrystalRpt.Formulas(5) = "ProvisionForBonus = " & V_ProvisionForBonus
        CrystalRpt.Formulas(6) = "ProvisionForTax = " & V_ProvisionForTax
        CrystalRpt.Formulas(7) = "TOTAL_CURRENT_ASSET = " & V_Total_Current_Asset
        CrystalRpt.Formulas(8) = "NET_PROPERTY_EQUIPMENT = " & V_Net_Propert_Equipment
        CrystalRpt.Formulas(9) = "OTHER_ASSETS = " & V_Other_Assets
        CrystalRpt.Formulas(10) = "Tax_Credits = " & V_TaxCredit
        CrystalRpt.Formulas(11) = "CurrentMonthYear = '" & Format(dtpTo, "MM/DD/YYYY") & "'"
        CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
        CrystalRpt.PageZoom 89
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpTo = LOGDATE
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:


    Dim V_GrossSales, V_SalesDiscountsAndReturns, V_CostOfSales As Double
    Dim V_LessSellingExpense, V_LessAdminExpense, V_LessOtherExpense, V_AddOtherIncome As Double

    Dim Cummulative_Cash_GrossSales, Cummulative_Charge_GrossSales, Cummulative_Cash_SalesDiscountsAndReturns, Cummulative_Charge_SalesDiscountsAndReturns As Double
    Dim Cummulative_Cash_CostOfSales, Cummulative_Charge_CostOfSales, Cummulative_LessSellingExpense, Cummulative_LessAdminExpense As Double
    Dim Cummulative_LessOtherExpense, Cummulative_AddOtherIncome As Double
    If IsDate(dtpTo) = False Then
        MsgSpeechBox "Error In Date"
        Exit Sub
    End If
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_Det where (jdate <= '" & CDate(dtpTo) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        '================ CUMMULATIVE ================
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & CASH_SALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_GrossSales = N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        'Set rsJOURNAL_DET = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & CHARGE_SALES)
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_GrossSales = N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        V_GrossSales = Cummulative_Cash_GrossSales + Cummulative_Charge_GrossSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & CASH_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & CHARGE_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        V_SalesDiscountsAndReturns = Cummulative_Cash_SalesDiscountsAndReturns + Cummulative_Charge_SalesDiscountsAndReturns
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='63' OR AMIS_ChartAccount.Headers=" & CASH_COSTOFSALES & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & CHARGE_COSTOFSALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        V_CostOfSales = Cummulative_Cash_CostOfSales + Cummulative_Charge_CostOfSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "'" & _
                                             " AND (AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode <> " & ADMIN_EXPENSE & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessSellingExpense = N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        V_LessSellingExpense = Cummulative_LessSellingExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode = " & ADMIN_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessAdminExpense = N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        V_LessAdminExpense = Cummulative_LessAdminExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & OTHER_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessOtherExpense = N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        V_LessOtherExpense = Cummulative_LessOtherExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & OTHER_INCOME)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_AddOtherIncome = N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        V_AddOtherIncome = Cummulative_AddOtherIncome

        V_NetIncomeOrLoss = (((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense)) - V_LessOtherExpense + V_AddOtherIncome

        'V_NetIncomeOrLoss = ((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense)
        '        If V_NetIncomeOrLoss > 0 Then
        '            V_NetIncomeOrLoss = Round((((((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense))) - V_LessOtherExpense) + V_AddOtherIncome, 2)
        '            If V_NetIncomeOrLoss > 0 Then V_ProvisionForBonus = (V_NetIncomeOrLoss * 0.2)
        '            If V_NetIncomeOrLoss > 0 Then V_ProvisionForTax = ((V_NetIncomeOrLoss - V_ProvisionForBonus) * 0.32)
        '            If V_NetIncomeOrLoss > 0 Then V_NetIncomeOrLoss = V_NetIncomeOrLoss - (V_ProvisionForBonus + V_ProvisionForTax)
        '        Else
        '            V_NetIncomeOrLoss = Round(((((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense)) - V_LessOtherExpense) + Abs(V_AddOtherIncome), 2)
        '            If V_NetIncomeOrLoss > 0 Then V_ProvisionForBonus = (V_NetIncomeOrLoss * 0.2)
        '            If V_NetIncomeOrLoss > 0 Then V_ProvisionForTax = ((V_NetIncomeOrLoss - V_ProvisionForBonus) * 0.32)
        '            If V_NetIncomeOrLoss > 0 Then V_NetIncomeOrLoss = V_NetIncomeOrLoss - (V_ProvisionForBonus + V_ProvisionForTax)
        '        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Total_Current_Asset from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & CURRENT_ASSET)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            V_Total_Current_Asset = N2Str2Zero(rsJournal_Det!Total_Current_Asset)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as TaxCredit from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Titles=" & TAX_CREDITS)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            V_TaxCredit = N2Str2Zero(rsJournal_Det!TaxCredit)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Property_Equipment from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Titles=" & PROPERTY_EQUIPMENT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            V_Propert_Equipment = N2Str2Zero(rsJournal_Det!PROPERTY_EQUIPMENT)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as AccumDepreciation from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Titles=" & ACCUMULATED_DEPRECIATION)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            V_AccumDepreciation = N2Str2Zero(rsJournal_Det!AccumDepreciation)
        End If
        V_Net_Propert_Equipment = V_Propert_Equipment + V_AccumDepreciation
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Other_Assets from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Titles=" & OTHER_ASSET)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            V_Other_Assets = N2Str2Zero(rsJournal_Det!Other_Assets)
        End If
        ShowBalanceSheetReport "BalanceSheet", "FinancialStatement\FinancialStatements\", "({Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & "))", "BALANCE SHEETS", "AS OF: " & Format(dtpTo, "long date"), True
        Screen.MousePointer = 0
    Else
        ShowNoRecord
    End If
    LogAudit "V", "BALANCE SHEET", dtpTo
    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub dtpTo_Change()
    dtpFrom.Value = firstDay(dtpTo)
End Sub

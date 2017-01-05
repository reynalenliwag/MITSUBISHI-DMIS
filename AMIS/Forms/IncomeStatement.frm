VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISIncomeStatement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income Statement"
   ClientHeight    =   1560
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4830
   ForeColor       =   &H00FFFFFF&
   Icon            =   "IncomeStatement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
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
      Left            =   2400
      MouseIcon       =   "IncomeStatement.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "IncomeStatement.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   645
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
      Left            =   1530
      MouseIcon       =   "IncomeStatement.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "IncomeStatement.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   645
      Width           =   885
   End
   Begin Crystal.CrystalReport rptAMISIncomeStatement 
      Left            =   960
      Top             =   990
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
      Format          =   131334145
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3030
      TabIndex        =   3
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
      Format          =   131334145
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
      Top             =   180
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
      Top             =   180
      Width           =   435
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   6
      Top             =   3000
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISIncomeStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                            As ADODB.Recordset
Dim rsJournal_Det                                           As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdPrint_Click()
'On Error GoTo ErrorCode:



    Dim Prev_dtpFrom, Prev_dtpTo                            As String
    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
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
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("select * from AMIS_Journal_HD where (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Dim rsProfile                                       As ADODB.Recordset
        rptAMISIncomeStatement.Reset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then
            rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENTS"
            rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENTS"
        End If
        '================ CUMMULATIVE ================
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='41'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='51'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='52'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='61' OR AMIS_ChartAccount.Headers='63')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='62'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND ((AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode <> '40')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode = '40'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='91'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='81'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        '=========================================
        '================ CURRENT ================
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='41'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='51'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='52'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='61' OR AMIS_ChartAccount.Headers='63')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='62'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "')" & _
                                             " AND ((AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.DepartmentCode <> '40')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode = '40'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='91'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='81'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        '=========================================
        '================ PREVIOUS ================
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='41'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='51'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='52'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='61' OR AMIS_ChartAccount.Headers='63')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='62'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "')" & _
                                             " AND ((AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.DepartmentCode <> '40')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode = '40'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='91'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='81'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        '=========================================
        PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatements.rpt", "{Journal_Hd.jtype} = 'CLO' AND {Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") and year({Journal_Hd.jdate}) = " & Year(dtpTo), DMIS_REPORT_Connection, 1
    Else
        ShowNoRecord
    End If
    LogAudit "V", "INCOME STATEMENT", dtpFrom & "-" & dtpTo
    Exit Sub
ErrorCode:
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
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAMISRange = Nothing
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub


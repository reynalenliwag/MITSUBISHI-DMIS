VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E6BE8522-29DC-4EDD-813C-BAA34BBA1069}#2.0#0"; "WIZMACFORM.OCX"
Begin VB.Form frmAMISIncomeStatementByProduct 
   BorderStyle     =   0  'None
   Caption         =   "Income Statement"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4830
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "IncomeStatementByProduct.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin wizMacForm.wizMacApp wizMacApp1 
      Height          =   320
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   556
      Buttons         =   2
   End
   Begin VB.CheckBox chkInclude 
      Caption         =   "Include Income Statement Details"
      Height          =   255
      Left            =   1170
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2460
      MouseIcon       =   "IncomeStatementByProduct.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "IncomeStatementByProduct.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1140
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print"
      Height          =   825
      Left            =   1470
      MouseIcon       =   "IncomeStatementByProduct.frx":089E
      MousePointer    =   99  'Custom
      Picture         =   "IncomeStatementByProduct.frx":09F0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1140
      Width           =   885
   End
   Begin Crystal.CrystalReport rptAMISIncomeStatement 
      Left            =   870
      Top             =   1500
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
      TabIndex        =   0
      Top             =   360
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
      Format          =   19660801
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3030
      TabIndex        =   1
      Top             =   360
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
      Format          =   19660801
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   420
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2550
      TabIndex        =   6
      Top             =   420
      Width           =   435
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISIncomeStatementByProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD As ADODB.Recordset
Dim rsJournal_Det As ADODB.Recordset

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim PRODUCT As String
If dtpFrom > dtpTo Then
   MsgSpeechBox "Error In From and To date"
   Exit Sub
End If
'If REPORT_RANGETYPE = "JVS" Then
   Set rsJournal_HD = New ADODB.Recordset
   Set rsJournal_HD = gconAMIS.Execute("select * from Journal_HD where (jdate >= #" & dtpFrom & "# AND jdate <= #" & dtpTo & "#)")
   If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
      Dim rsProfile As ADODB.Recordset
      Set rsProfile = New ADODB.Recordset
      Set rsProfile = gconAMIS.Execute("Select * from Profile")
      If Not (rsProfile.EOF And rsProfile.BOF) Then
         rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
         rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
         rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - BY PRODUCT"
      End If
      '================ CUMMULATIVE ================
      PRODUCT = "'10'"
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='41' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='42' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='61' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='62' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='51' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='52' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='71' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='72' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='91' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as AddOtherIncome from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='81' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
      End If
      '=========================================
      '================ CURRENT ================
      PRODUCT = "'30'"
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='41' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='42' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='61' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='62' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='51' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='52' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='71' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='72' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='91' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as AddOtherIncome from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='81' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
      End If
      '=========================================
      '================ PREVIOUS ================
      PRODUCT = "'20'"
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='41' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='42' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='61' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='62' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='51' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='52' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='71' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='72' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='91' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
      End If
      Set rsJournal_Det = New ADODB.Recordset
      Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as AddOtherIncome from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= #" & dtpFrom & "# AND Journal_Det.jdate <= #" & dtpTo & "#) AND ChartAccount.Headers='81' AND ChartAccount.DepartmentCode = " & PRODUCT)
      If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
         rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
      End If
      '=========================================
      PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\IncomeStatementByProduct.rpt", "{Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
   Else
      ShowNoRecord
   End If
'End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
dtpTo = LOGDATE
wizMacApp1.MacCaption = Me.Caption
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

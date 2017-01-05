VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmHRMSProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Payroll"
   ClientHeight    =   7890
   ClientLeft      =   1545
   ClientTop       =   3180
   ClientWidth     =   7650
   ControlBox      =   0   'False
   ForeColor       =   &H00D8E9EC&
   Icon            =   "progress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7890
   ScaleWidth      =   7650
   Begin FlexCell.Grid Grid1 
      Height          =   6195
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   10927
      Appearance      =   0
      BackColor2      =   12907725
      BackColorBkg    =   -2147483645
      Cols            =   5
      DefaultFontName =   "Courier New"
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      GridColor       =   12632256
      Rows            =   30
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6570
      Picture         =   "progress.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Generation of Payroll Done "
      Top             =   7080
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Tax is Base on Annualized Computation"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7680
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Tax is Base on BIR Tax Table Bracket"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7680
      TabIndex        =   10
      Top             =   1050
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5700
      MouseIcon       =   "progress.frx":0762
      MousePointer    =   99  'Custom
      Picture         =   "progress.frx":08B4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Generate Payroll Now"
      Top             =   7080
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   735
      Left            =   6570
      MouseIcon       =   "progress.frx":0C02
      MousePointer    =   99  'Custom
      Picture         =   "progress.frx":0D54
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancel"
      Top             =   7080
      Width           =   855
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   90
      ScaleHeight     =   1155
      ScaleWidth      =   7455
      TabIndex        =   3
      Top             =   6390
      Width           =   7455
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   5115
         TabIndex        =   4
         Top             =   750
         Width           =   5115
         Begin VB.Label labName 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   5
            Top             =   -30
            Width           =   4395
         End
      End
      Begin wizProgBar.Prg gauProgress 
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   300
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   556
         Picture         =   "progress.frx":1092
         BackColor       =   14215660
         ForeColor       =   255
         BorderStyle     =   2
         BarPicture      =   "progress.frx":10AE
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   5445
         TabIndex        =   6
         Top             =   660
         Width           =   5445
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   7
            Top             =   0
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   609
            TX              =   "cmd1"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "progress.frx":10CA
         End
      End
      Begin VB.Label lblPercent 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   30
         Width           =   5595
      End
   End
   Begin VB.Label labCutOff 
      Caption         =   "Label1"
      Height          =   315
      Left            =   5040
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   2835
   End
End
Attribute VB_Name = "frmHRMSProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemporaryRecordSet                                              As ADODB.Recordset
Dim GRP_SSS                                                           As Integer
Dim GRP_PAGIBIG                                                       As Integer
Dim GRP_PHILHEALTH                                                    As Integer
Dim GRP_TAX                                                           As Integer
Dim GRP_LOAN                                                          As Integer
Dim GRP_OTHER                                                         As Integer
Dim SSS_BASIS                                                         As Integer
Dim PAGIBIG_BASIS                                                     As Integer
Dim PHEALTH_BASIS                                                     As Integer
Dim TAX_BASIS                                                         As Integer
Dim TAX_COMP                                                          As Integer
Dim WORKING_DAYS                                                      As Integer
Dim WORKING_HOURS                                                     As Integer
Dim AVERAGE_MONTH                                                     As Integer
Dim vPayrollGeneratingStatus                                          As String
Dim I                                                                 As Integer
Dim CNT                                                               As Integer

Dim XdedSalLoan                                                       As Double
Dim XdedCalLoan                                                       As Double
Dim XdedMPL                                                           As Double
Dim XdedHLL                                                           As Double
Dim XdedBLL                                                           As Double
Dim XdedOTHER                                                         As Double

Dim date_from                                                         As String
Dim date_to                                                           As String

Function GetFromSSSTable(vSWELDOTRIENTA As Currency, EMPE_OR_EMPR As String) As Double
    Set rsTemporaryRecordSet = New ADODB.Recordset
    Set rsTemporaryRecordSet = gconDMIS.Execute("SELECT * FROM HRMS_SSSTABLE WHERE " & vSWELDOTRIENTA & " >= RANGE1 AND " & vSWELDOTRIENTA & " < RANGE2 ")
    If Not (rsTemporaryRecordSet.BOF And rsTemporaryRecordSet.EOF) Then
        If EMPE_OR_EMPR = "EMPLOYER" Then
            GetFromSSSTable = rsTemporaryRecordSet!Owner_SSS
        Else
            GetFromSSSTable = rsTemporaryRecordSet!Emp_SSS
        End If
    End If
    Set rsTemporaryRecordSet = Nothing
End Function

Function GetFromPagIbigTable(vSWELDOTRIENTA As Currency, EMPE_OR_EMPR As String)
    If COMPANY_CODE = "HAI" Then
        Dim rsTemporaryRecordSet                                      As New ADODB.Recordset
        Set rsTemporaryRecordSet = gconDMIS.Execute("SELECT * FROM HRMS_PAGIBIGTABLE WHERE " & vSWELDOTRIENTA & " >= [FROM] AND " & vSWELDOTRIENTA & " < [TO]")
        If Not (rsTemporaryRecordSet.BOF And rsTemporaryRecordSet.EOF) Then
            Dim Perc                                                  As Double
            Perc = N2Str2Zero(rsTemporaryRecordSet!Percent)
            If EMPE_OR_EMPR = "EMPLOYEE" Then
                GetFromPagIbigTable = vSWELDOTRIENTA * Perc
            End If
            If EMPE_OR_EMPR = "EMPLOYER" Then
                GetFromPagIbigTable = (vSWELDOTRIENTA * Perc)
            End If
        End If
        Set rsTemporaryRecordSet = Nothing
    Else
        If vSWELDOTRIENTA <> 0 Then
            If EMPE_OR_EMPR = "EMPLOYER" Then
                GetFromPagIbigTable = 100
            End If
            If EMPE_OR_EMPR = "EMPLOYEE" Then
                GetFromPagIbigTable = 100
            End If
        End If
    End If
End Function

Function GetFromPhilHealthTable(vSWELDOTRIENTA As Currency, EMPR_OR_EMPE As String)
    Dim rsTemporaryRecordSet                                          As New ADODB.Recordset
    Set rsTemporaryRecordSet = gconDMIS.Execute("SELECT * FROM HRMS_PHICTABLE WHERE " & vSWELDOTRIENTA & " >= RANGE1 AND " & vSWELDOTRIENTA & " < RANGE2 ")
    If Not (rsTemporaryRecordSet.BOF And rsTemporaryRecordSet.EOF) Then
        If EMPR_OR_EMPE = "EMPLOYEE" Then
            GetFromPhilHealthTable = NumericVal(rsTemporaryRecordSet!Emp_MCR)
        Else
            GetFromPhilHealthTable = NumericVal(rsTemporaryRecordSet!Owner_MCR)
        End If
    End If
    Set rsTemporaryRecordSet = Nothing
End Function

Function GetFromTaxTable(TAXCODE As String, EMPSAL_GROSS As Variant, EMP_PAYROLL_TYPE As String)
    GetFromTaxTable = 0
    Dim RSTAX                                                         As New ADODB.Recordset
    Dim COLNO                                                         As Integer
    Dim RESULT_TAX                                                    As Double
    Set rsTemporaryRecordSet = New ADODB.Recordset
    Set rsTemporaryRecordSet = gconDMIS.Execute("SELECT * FROM HRMS_TAXTABLEDETAILS WHERE TAXBASIS = '" & EMP_PAYROLL_TYPE & "' AND TAXCODE = '" & TAXCODE & "'")
    If Not (rsTemporaryRecordSet.BOF And rsTemporaryRecordSet.EOF) Then
        If 2 > EMPSAL_GROSS Then
            RESULT_TAX = rsTemporaryRecordSet!Col1
            COLNO = 1
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col2 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col3 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col2
            COLNO = 2
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col3 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col4 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col3
            COLNO = 3
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col4 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col5 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col4
            COLNO = 4
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col5 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col6 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col5
            COLNO = 5
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col6 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col7 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col6
            COLNO = 6
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col7 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col8 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col7
            COLNO = 7
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col8 Then
            RESULT_TAX = rsTemporaryRecordSet!Col8
            COLNO = 8
        End If
        Set RSTAX = gconDMIS.Execute("SELECT * FROM HRMS_TAXTABLE WHERE TAXBASIS = '" & EMP_PAYROLL_TYPE & "'")
        If Not (RSTAX.BOF And RSTAX.EOF) Then
            If COLNO = 1 Then
                GetFromTaxTable = 1
            End If
            If COLNO = 2 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per2)
            End If
            If COLNO = 3 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per3) + RSTAX!EXp3
            End If
            If COLNO = 4 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per4) + RSTAX!EXp4
            End If
            If COLNO = 5 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per5) + RSTAX!EXp5
            End If
            If COLNO = 6 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per6) + RSTAX!EXp6
            End If
            If COLNO = 7 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per7) + RSTAX!EXp7
            End If
            If COLNO = 8 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per8) + RSTAX!EXp8
            End If
        End If
    End If
    Set rsTemporaryRecordSet = Nothing
End Function

Function FindPrevMonth(vMONTH As String) As String
    If vMONTH = "January" Then FindPrevMonth = "12"
    If vMONTH = "February" Then FindPrevMonth = "1"
    If vMONTH = "March" Then FindPrevMonth = "2"
    If vMONTH = "April" Then FindPrevMonth = "3"
    If vMONTH = "May" Then FindPrevMonth = "4"
    If vMONTH = "June" Then FindPrevMonth = "5"
    If vMONTH = "July" Then FindPrevMonth = "6"
    If vMONTH = "August" Then FindPrevMonth = "7"
    If vMONTH = "September" Then FindPrevMonth = "8"
    If vMONTH = "October" Then FindPrevMonth = "9"
    If vMONTH = "November" Then FindPrevMonth = "10"
    If vMONTH = "December" Then FindPrevMonth = "11"
End Function

Function GetPreviousGross(vEmployeeNo As String) As Currency
    Dim rsTemporaryRecordSet                                          As New ADODB.Recordset
    Dim PREV_FROM                                                     As Date
    Dim PREV_TO                                                       As Date

    If frmHRMSGenerate.cboQuensina.Text = "1st Cut-Off" Then
        PREV_FROM = DateSerial(frmHRMSGenerate.cboYear, What_month(frmHRMSGenerate.cboMonth), 6)
        PREV_TO = DateSerial(frmHRMSGenerate.cboYear, What_month(frmHRMSGenerate.cboMonth), 20)
    End If
    If frmHRMSGenerate.cboQuensina.Text = "2nd Cut-Off" Then
        PREV_FROM = DateSerial(frmHRMSGenerate.cboYear, What_month(frmHRMSGenerate.cboMonth), 5)
        PREV_TO = DateSerial(frmHRMSGenerate.cboYear, FindPrevMonth(frmHRMSGenerate.cboMonth), 21)
    End If

    Set rsTemporaryRecordSet = gconDMIS.Execute("select GROSS FROM HRMS_PAYROLL WHERE EMPNO = '" & vEmployeeNo & _
                                                "' AND PAYDATEFROM = '" & PREV_FROM & _
                                                "' AND PAYDATETO = '" & PREV_TO & "'")
    If Not (rsTemporaryRecordSet.BOF And rsTemporaryRecordSet.EOF) Then
        GetPreviousGross = NumericVal(rsTemporaryRecordSet!GROSS)
    Else
        GetPreviousGross = 0
    End If
    Set rsTemporaryRecordSet = Nothing
End Function

Function getAssumedBasic(GENDATE As Date, Salari As Double) As Double
    Dim AssumedBasic                                                  As Double
    If labCutOff.Caption = "1st Cut-Off" Then
        AssumedBasic = Salari * (12 - MONTH(GENDATE)) + (Salari / 2)
    Else
        AssumedBasic = Salari * (12 - MONTH(GENDATE))
    End If
    getAssumedBasic = AssumedBasic
End Function

Function getAssumedPH(GENDATE As Date, Salari As Double) As Double
    getAssumedPH = GetFromPhilHealthTable(CCur(Salari), "EMPLOYEE") * (12 - MONTH(GENDATE))
End Function

Function getAssumedSSS(GENDATE As Date, Salari As Double) As Double
    getAssumedSSS = GetFromSSSTable(CCur(Salari), "EMPLOYEE") * (12 - MONTH(GENDATE))
End Function

Function getAssumedPagIbig(GENDATE As Date, Salari As Double) As Double
    getAssumedPagIbig = GetFromPagIbigTable(CCur(Salari), "EMPLOYEE") * (12 - MONTH(GENDATE))
End Function

Function GETDEDUCTION_DESCRIPTION(xxxDedCode As String) As String
    Dim rsDedDesc                                                     As ADODB.Recordset
    Set rsDedDesc = gconDMIS.Execute("SELECT Description  FROM HRMS_DeductionCode WHERE Code='" & xxxDedCode & "'")
    If Not rsDedDesc.EOF Or Not rsDedDesc.BOF Then
        GETDEDUCTION_DESCRIPTION = Null2String(rsDedDesc!Description)
    End If
End Function

Function GETLOAN_DESCRIPTION(xxxLoanCode As String) As String
    Dim rsLoanDesc                                                    As ADODB.Recordset
    Set rsLoanDesc = gconDMIS.Execute("SELECT Description  FROM HRMS_LoanCode WHERE Code='" & xxxLoanCode & "'")
    If Not rsLoanDesc.EOF Or Not rsLoanDesc.BOF Then
        GETLOAN_DESCRIPTION = Null2String(rsLoanDesc!Description)
    End If
End Function

Function GETHASHVALUE(xAccountNo As String, netsal As Double) As Double

    xAccountNo = Repleys(xAccountNo)

    Dim firstvalue
    Dim secondvalue
    Dim thirdvalue

    Dim firstvalue_sal
    Dim secondvalue_sal
    Dim thirdvalue_sal

    Dim X                                                             As Integer
    Dim count                                                         As Integer
    count = 0
    Dim matt(15)                                                      As Integer

    For X = 1 To Len(xAccountNo)
        If IsNumeric(Mid(xAccountNo, X, 1)) Then
            count = count + 1
            'MsgBox Mid(xAccountNo, x, 1)
            matt(count) = Mid(xAccountNo, X, 1)
        End If
    Next
    firstvalue = CInt(CStr(matt(5)) & CStr(matt(6)))
    secondvalue = CInt(CStr(matt(7)) & CStr(matt(8)))
    thirdvalue = CInt(CStr(matt(9)) & CStr(matt(10)))

    '    firstvalue = NumericVal(Mid(xAccountNo, 7, 2))
    '    secondvalue = NumericVal(Mid(xAccountNo, 9, 2))
    '    thirdvalue = NumericVal(Mid(xAccountNo, 11, 1)) & NumericVal(Mid(xAccountNo, 13, 1))


    firstvalue_sal = NumericVal(firstvalue) * netsal
    secondvalue_sal = NumericVal(secondvalue) * netsal
    thirdvalue_sal = NumericVal(thirdvalue) * netsal

    GETHASHVALUE = Round((firstvalue_sal + secondvalue_sal + thirdvalue_sal), 2)
End Function

Function ComputeTotalCommission(LEVEL As String, EMPNO As String, CUTOFF As String, PAYMONTH As Integer, PAYYEAR As Integer) As Double
    ComputeTotalCommission = 0
    Dim rsCommission                                                  As ADODB.Recordset
    Set rsCommission = New ADODB.Recordset
    Set rsCommission = gconDMIS.Execute("SELECT SUM(AMOUNT) AS TOTALCOMMISSION " & _
                                      " FROM HRMS_COMMISSION " & _
                                      " WHERE EMPLEVEL = " & LEVEL & _
                                      " AND EMPNO = '" & EMPNO & "'" & _
                                      " AND CUT_OFF = " & CUTOFF & _
                                      " AND PAY_MONTH = " & PAYMONTH & _
                                      " AND PAY_YEAR = " & PAYYEAR & "")
                                     

    If Not rsCommission.EOF And Not rsCommission.BOF Then
        ComputeTotalCommission = N2Str2Zero(rsCommission!TOTALCOMMISSION)
    End If
    ComputeTotalCommission = Round(ComputeTotalCommission, 2)
    Set rsCommission = Nothing
End Function

Function ComputeTotalCommissionTax(LEVEL As String, EMPNO As String, CUTOFF As String, PAYMONTH As Integer, PAYYEAR As Integer) As Double
    ComputeTotalCommissionTax = 0
    Dim rsCommission                                                  As ADODB.Recordset
    Set rsCommission = New ADODB.Recordset
    Set rsCommission = gconDMIS.Execute("SELECT SUM(TAX) AS TOTALTAX " & _
                                      " FROM HRMS_COMMISSION " & _
                                      " WHERE EMPLEVEL = " & LEVEL & _
                                      " AND EMPNO = '" & EMPNO & "'" & _
                                      " AND CUT_OFF = " & CUTOFF & _
                                      " AND PAY_MONTH = " & PAYMONTH & _
                                      " AND PAY_YEAR = " & PAYYEAR & "")
                                      

    If Not rsCommission.EOF And Not rsCommission.BOF Then
        ComputeTotalCommissionTax = N2Str2Zero(rsCommission!TOTALTAX)
    End If
    ComputeTotalCommissionTax = Round(ComputeTotalCommissionTax, 2)
    Set rsCommission = Nothing
End Function

Function ComputeTotalNonTaxAdj(LEVEL As String, EMPNO As String, CUTOFF As String, PAYMONTH As Integer, PAYYEAR As Integer) As Double
    ComputeTotalNonTaxAdj = 0
    Dim rsAdjustment                                                  As ADODB.Recordset
    Set rsAdjustment = New ADODB.Recordset
    Set rsAdjustment = gconDMIS.Execute("SELECT SUM(AMOUNT) AS NONTAXADJ " & _
                                      " FROM HRMS_ADJUSTMENT " & _
                                      " WHERE EMPLEVEL = " & LEVEL & _
                                      " AND EMPNO = '" & EMPNO & "'" & _
                                      " AND TYPE = 'NT' " & _
                                      " AND CUT_OFF = " & CUTOFF & _
                                      " AND PAY_MONTH = " & PAYMONTH & _
                                      " AND PAY_YEAR = " & PAYYEAR & "")

    If Not rsAdjustment.EOF And Not rsAdjustment.BOF Then
        ComputeTotalNonTaxAdj = N2Str2Zero(rsAdjustment!NONTAXADJ)
    End If
    ComputeTotalNonTaxAdj = Round(ComputeTotalNonTaxAdj, 2)
    Set rsAdjustment = Nothing
End Function

Function ComputeTotalTaxAdj(LEVEL As String, EMPNO As String, CUTOFF As String, PAYMONTH As Integer, PAYYEAR As Integer) As Double
    ComputeTotalTaxAdj = 0
    Dim rsAdjustment                                                  As ADODB.Recordset
    Set rsAdjustment = New ADODB.Recordset
    Set rsAdjustment = gconDMIS.Execute("SELECT SUM(AMOUNT) AS TAXADJ " & _
                                      " FROM HRMS_ADJUSTMENT " & _
                                      " WHERE EMPLEVEL = " & LEVEL & _
                                      " AND EMPNO = '" & EMPNO & "'" & _
                                      " AND TYPE = 'T' " & _
                                      " AND CUT_OFF = " & CUTOFF & _
                                      " AND PAY_MONTH = " & PAYMONTH & _
                                      " AND PAY_YEAR = " & PAYYEAR & "")

    If Not rsAdjustment.EOF And Not rsAdjustment.BOF Then
        ComputeTotalTaxAdj = N2Str2Zero(rsAdjustment!TAXADJ)
    End If
    ComputeTotalTaxAdj = Round(ComputeTotalTaxAdj, 2)
    Set rsAdjustment = Nothing
End Function

Function ComputeTotalOT(LEVEL As String, EMPNO As String, CUTOFF As String, PAYMONTH As Integer, PAYYEAR As Integer) As Double
    ComputeTotalOT = 0
    Dim rsOvertime                                                    As ADODB.Recordset
    Set rsOvertime = New ADODB.Recordset
    Set rsOvertime = gconDMIS.Execute("SELECT SUM(AMOUNT) AS TOTALOVERTIME " & _
                                    " FROM HRMS_OVERTIME " & _
                                    " WHERE EMPLEVEL = " & LEVEL & _
                                    " AND EMPNO = '" & EMPNO & "'" & _
                                    " AND CUT_OFF = " & CUTOFF & _
                                    " AND PAY_MONTH = " & PAYMONTH & _
                                    " AND PAY_YEAR = " & PAYYEAR & "")
                                    

    If Not rsOvertime.EOF And Not rsOvertime.BOF Then
        ComputeTotalOT = N2Str2Zero(rsOvertime!TOTALOVERTIME)
    End If
    ComputeTotalOT = Round(ComputeTotalOT, 2)
    Set rsOvertime = Nothing
End Function

Function ComputeTotalAdvance(LEVEL As String, EMPNO As String, CUTOFF As String, PAYMONTH As Integer, PAYYEAR As Integer) As Double
    ComputeTotalAdvance = 0
    Dim rsSalaryAdvance                                               As ADODB.Recordset
    Set rsSalaryAdvance = New ADODB.Recordset
    Set rsSalaryAdvance = gconDMIS.Execute("SELECT SUM(AMOUNT) AS TOTALAMOUNT " & _
                                         " FROM HRMS_ADVANCE " & _
                                         " WHERE EMPNO = '" & EMPNO & "'" & _
                                         " AND CUT_OFF = " & CUTOFF & _
                                         " AND PAY_MONTH = " & PAYMONTH & _
                                         " AND PAY_YEAR = " & PAYYEAR & "")
                                         

    If Not rsSalaryAdvance.BOF And Not rsSalaryAdvance.EOF Then
        ComputeTotalAdvance = N2Str2Zero(rsSalaryAdvance!TotalAmount)
    End If
    ComputeTotalAdvance = Round(ComputeTotalAdvance, 2)
    Set rsSalaryAdvance = Nothing
End Function

Function ComputeDeductionAbsences(LEVEL As String, EMPNO As String, CUTOFF As String, PAYMONTH As Integer, PAYYEAR As Integer) As Double
    ComputeDeductionAbsences = 0
    Dim rsDeductions                                                  As ADODB.Recordset
    Set rsDeductions = New ADODB.Recordset
    Set rsDeductions = gconDMIS.Execute("SELECT * " & _
                                      " FROM HRMS_DEDUCTIONS " & _
                                      " WHERE EMPLEVEL = " & LEVEL & _
                                      " AND EMPNO = '" & EMPNO & "'" & _
                                      " AND (PARTICULAR = 'WD' OR PARTICULAR = 'HD')" & _
                                      " AND CUT_OFF = " & CUTOFF & _
                                      " AND PAY_MONTH = " & PAYMONTH & _
                                      " AND PAY_YEAR = " & PAYYEAR & "")
                                      

    If Not rsDeductions.EOF And Not rsDeductions.BOF Then
        rsDeductions.MoveFirst
        Do While Not rsDeductions.EOF
            ComputeDeductionAbsences = ComputeDeductionAbsences + N2Str2Zero(rsDeductions!AMOUNT)

            gconDMIS.Execute ("INSERT INTO HRMS_PAYROLL_DET " & _
                              "(EMPLEVEL, EMPNO, TRANTYPE, DET_AMOUNT, ISADD, PAYPERIOD_FROM, PAYPERIOD_TO, DET_CODE, DET_DESC, CUT_OFF, PAY_MONTH, PAY_YEAR) values(" _
                            & LEVEL & _
                              ",'" & EMPNO & _
                              "','" & "D" & _
                              "'," & NumericVal(rsDeductions!AMOUNT) & _
                              "," & 0 & _
                              ",'" & GENFROM & _
                              "','" & GENTO & _
                              "','" & Null2String(rsDeductions!PARTICULAR) & _
                              "','" & GETDEDUCTION_DESCRIPTION(Null2String(rsDeductions!PARTICULAR)) & _
                              "','" & CUTOFF & _
                              "'," & PAYMONTH & _
                              "," & PAYYEAR & ")")

            rsDeductions.MoveNext
        Loop
    End If
    ComputeDeductionAbsences = Round(ComputeDeductionAbsences, 2)
    Set rsDeductions = Nothing
End Function

Function ComputeDeductionLateUndertime(LEVEL As String, EMPNO As String, CUTOFF As String, PAYMONTH As Integer, PAYYEAR As Integer) As Double
    Dim DET_DESC                                                      As String
    ComputeDeductionLateUndertime = 0
    Dim rsDeductions                                                  As ADODB.Recordset
    Set rsDeductions = New ADODB.Recordset
    Set rsDeductions = gconDMIS.Execute("SELECT * " & _
                                      " FROM HRMS_DEDUCTIONS " & _
                                      " WHERE EMPLEVEL = " & LEVEL & _
                                      " AND EMPNO = '" & EMPNO & "'" & _
                                      " AND (PARTICULAR = 'LT' OR PARTICULAR = 'UT')" & _
                                      " AND CUT_OFF = " & CUTOFF & _
                                      " AND PAY_MONTH = " & PAYMONTH & _
                                      " AND PAY_YEAR = " & PAYYEAR & "")
                                      
                                      

    If Not rsDeductions.EOF And Not rsDeductions.BOF Then
        rsDeductions.MoveFirst
        Do While Not rsDeductions.EOF
            ComputeDeductionLateUndertime = ComputeDeductionLateUndertime + N2Str2Zero(rsDeductions!AMOUNT)

            gconDMIS.Execute ("INSERT INTO HRMS_PAYROLL_DET " & _
                              "(EMPLEVEL, EMPNO, TRANTYPE, DET_AMOUNT, ISADD, PAYPERIOD_FROM, PAYPERIOD_TO, DET_CODE, DET_DESC, CUT_OFF, PAY_MONTH, PAY_YEAR) values(" _
                            & LEVEL & _
                              ",'" & EMPNO & _
                              "','" & "D" & _
                              "'," & NumericVal(rsDeductions!AMOUNT) & _
                              "," & 0 & _
                              ",'" & GENFROM & _
                              "','" & GENTO & _
                              "','" & Null2String(rsDeductions!PARTICULAR) & _
                              "','" & GETDEDUCTION_DESCRIPTION(Null2String(rsDeductions!PARTICULAR)) & _
                              "','" & CUTOFF & _
                              "'," & PAYMONTH & _
                              "," & PAYYEAR & ")")

            rsDeductions.MoveNext
        Loop
    End If
    ComputeDeductionLateUndertime = Round(ComputeDeductionLateUndertime, 2)
    Set rsDeductions = Nothing
    DET_DESC = ""
End Function

Function ComputeDeductionOthers(LEVEL As String, EMPNO As String, CUTOFF As String, PAYMONTH As Integer, PAYYEAR As Integer) As Double
    ComputeDeductionOthers = 0
    Dim rsDeductions                                                  As ADODB.Recordset
    Set rsDeductions = New ADODB.Recordset
    Set rsDeductions = gconDMIS.Execute("SELECT * " & _
                                      " FROM HRMS_DEDUCTIONS " & _
                                      " WHERE EMPLEVEL = " & LEVEL & _
                                      " AND EMPNO = '" & EMPNO & "'" & _
                                      " AND (PARTICULAR <> 'LT' AND PARTICULAR <> 'UT' AND PARTICULAR <> 'HD' AND PARTICULAR <> 'WD')" & _
                                      " AND CUT_OFF = " & CUTOFF & _
                                      " AND PAY_MONTH = " & PAYMONTH & _
                                      " AND PAY_YEAR = " & PAYYEAR & "")
                                      
    If Not rsDeductions.EOF And Not rsDeductions.BOF Then
        rsDeductions.MoveFirst
        Do While Not rsDeductions.EOF
            ComputeDeductionOthers = ComputeDeductionOthers + N2Str2Zero(rsDeductions!AMOUNT)

            gconDMIS.Execute ("INSERT INTO HRMS_PAYROLL_DET " & _
                              "(EMPLEVEL, EMPNO, TRANTYPE, DET_AMOUNT, ISADD, PAYPERIOD_FROM, PAYPERIOD_TO, DET_CODE, DET_DESC, CUT_OFF, PAY_MONTH, PAY_YEAR) values(" _
                            & LEVEL & _
                              ",'" & EMPNO & _
                              "','" & "D" & _
                              "'," & NumericVal(rsDeductions!AMOUNT) & _
                              "," & 0 & _
                              ",'" & GENFROM & _
                              "','" & GENTO & _
                              "','" & Null2String(rsDeductions!PARTICULAR) & _
                              "','" & GETDEDUCTION_DESCRIPTION(Null2String(rsDeductions!PARTICULAR)) & _
                              "','" & CUTOFF & _
                              "'," & PAYMONTH & _
                              "," & PAYYEAR & ")")

            rsDeductions.MoveNext
        Loop
    End If
    ComputeDeductionOthers = Round(ComputeDeductionOthers, 2)
    Set rsDeductions = Nothing
End Function

Function ComputeAllowance(LEVEL As String, EMPNO As String, CUTOFF As String, PAYMONTH As Integer, PAYYEAR As Integer, ALLOWANCE As Double) As Double
    ComputeAllowance = 0
    Dim VTOTALMIN                                                     As Integer
    Dim VALLOWANCEPERMINUTE                                           As Double

    VTOTALMIN = 0
    VALLOWANCEPERMINUTE = 0

    Dim rsDeductions                                                  As ADODB.Recordset
    Set rsDeductions = New ADODB.Recordset
    Set rsDeductions = gconDMIS.Execute("SELECT ISNULL(SUM(NOMIN),0) AS TOTALM " & _
                                      " FROM HRMS_DEDUCTIONS " & _
                                      " WHERE EMPNO ='" & EMPNO & "'" & _
                                      " AND CUT_OFF = " & CUTOFF & _
                                      " AND PAY_MONTH = " & PAYMONTH & _
                                      " AND PAY_YEAR = " & PAYYEAR & "")
                                     

    If Not (rsDeductions.EOF And rsDeductions.BOF) Then
        If ALLOWANCE > 0 Then
            VTOTALMIN = N2Str2Zero(rsDeductions!TOTALM)
            VALLOWANCEPERMINUTE = (((ALLOWANCE * 12) / 314) / 8) / 60
            ComputeAllowance = (ALLOWANCE / 2) - (VALLOWANCEPERMINUTE * (VTOTALMIN))
        End If
    End If
    ComputeAllowance = Round(ComputeAllowance, 2)
    Set rsDeductions = Nothing
End Function

Private Function SetDeductionBasis(DEDTYPE As String, EmpPayrollEmpNo As String, EmpPayrollGroup As Integer, EmpDedGroup As Integer, EmpSalGross As Double, EmpSalBasic As Double, EmpUTAndAbs As Double, EmpOTAndHol As Double) As Double
    Dim rsPreviousPayroll                                             As ADODB.Recordset
    Dim PrevSalBasis                                                  As Double
    PrevSalBasis = 0
    SetDeductionBasis = 0
    If EmpDedGroup = 1 And labCutOff.Caption = "1st Cut-Off" Then
        Select Case EmpPayrollGroup
            Case 1: SetDeductionBasis = EmpSalGross + EmpSalBasic
            Case 2: SetDeductionBasis = EmpSalBasic * 2
            Case 3: SetDeductionBasis = (EmpSalBasic * 2) - EmpUTAndAbs
            Case 4: SetDeductionBasis = ((EmpSalBasic * 2) + EmpOTAndHol) - EmpUTAndAbs
            Case Else
                SetDeductionBasis = 0
        End Select
    End If

    If EmpDedGroup = 2 And labCutOff.Caption = "2nd Cut-Off" Then
        Set rsPreviousPayroll = New ADODB.Recordset
        Select Case EmpPayrollGroup
            Case 1
                Set rsPreviousPayroll = gconDMIS.Execute("SELECT SUM(GROSS) AS TOTAL_BASIS FROM HRMS_PAYROLL WHERE EMPNO = '" & EmpPayrollEmpNo & "' AND PAY_MONTH = " & PAY_MONTH & " AND PAY_YEAR = " & PAY_YEAR)
            Case 2
                Set rsPreviousPayroll = gconDMIS.Execute("SELECT SUM(RATE) AS TOTAL_BASIS FROM HRMS_PAYROLL  WHERE EMPNO = '" & EmpPayrollEmpNo & "' AND PAY_MONTH = " & PAY_MONTH & " AND PAY_YEAR = " & PAY_YEAR)
            Case 3
                Set rsPreviousPayroll = gconDMIS.Execute("SELECT SUM(RATE) - (SUM(UNDERTIME) + SUM(ABSENT)) AS TOTAL_BASIS FROM HRMS_PAYROLL WHERE EMPNO = '" & EmpPayrollEmpNo & "' AND PAY_MONTH = " & PAY_MONTH & " AND PAY_YEAR = " & PAY_YEAR)
            Case 4
                Set rsPreviousPayroll = gconDMIS.Execute("SELECT (SUM(RATE) + SUM(OVERTIME) + SUM(HOLIDAY) + SUM(TAXABLEADJ)) - (SUM(UNDERTIME) + SUM(ABSENT)) AS TOTAL_BASIS FROM HRMS_PAYROLL WHERE EMPNO = '" & EmpPayrollEmpNo & "' AND PAY_MONTH = " & PAY_MONTH & " AND PAY_YEAR = " & PAY_YEAR)
        End Select

        If Not rsPreviousPayroll.EOF And Not rsPreviousPayroll.BOF Then
            PrevSalBasis = Round(NumericVal(rsPreviousPayroll!total_basis), 2)
        End If
        Select Case EmpPayrollGroup
            Case 1: SetDeductionBasis = EmpSalGross + PrevSalBasis
            Case 2: SetDeductionBasis = EmpSalBasic + PrevSalBasis
            Case 3: SetDeductionBasis = EmpSalBasic - EmpUTAndAbs + PrevSalBasis
            'Case 4: SetDeductionBasis = (EmpSalBasic + EmpOTAndHol) - EmpUTAndAbs + PrevSalBasis
            Case 4: SetDeductionBasis = ((EmpSalBasic * 2) + EmpOTAndHol) - EmpUTAndAbs
        
        End Select
    End If
    If EmpDedGroup = 3 Then
        Select Case EmpPayrollGroup
            Case 1: SetDeductionBasis = EmpSalGross
            Case 2: SetDeductionBasis = EmpSalBasic
            Case 3: SetDeductionBasis = EmpSalBasic - EmpUTAndAbs
            Case 4: SetDeductionBasis = (EmpSalBasic + EmpOTAndHol) - EmpUTAndAbs
        End Select

    End If
End Function

Sub StoreDeductionSetValues(vGROUP As Variant)
    If Not IsNumeric(vGROUP) Then
        vGROUP = 1
    End If

    GRP_SSS = 0
    GRP_PAGIBIG = 0
    GRP_PHILHEALTH = 0
    GRP_TAX = 0
    GRP_LOAN = 0
    GRP_OTHER = 0
    SSS_BASIS = 0
    PAGIBIG_BASIS = 0
    PHEALTH_BASIS = 0
    TAX_BASIS = 0
    TAX_COMP = 0

    WORKING_DAYS = 314

    Dim rsTemporaryRecordSet                                          As New ADODB.Recordset
    Set rsTemporaryRecordSet = gconDMIS.Execute("SELECT * FROM HRMS_SETUPDEDUCTION WHERE DEDUCTION_SET = " & vGROUP & "")
    If Not (rsTemporaryRecordSet.BOF And rsTemporaryRecordSet.EOF) Then
        GRP_SSS = NumericVal(rsTemporaryRecordSet!SSS)
        GRP_PAGIBIG = NumericVal(rsTemporaryRecordSet!PAGIBIG)
        GRP_PHILHEALTH = NumericVal(rsTemporaryRecordSet!PHILHEALTH)
        GRP_TAX = NumericVal(rsTemporaryRecordSet!TAX)
        GRP_LOAN = NumericVal(rsTemporaryRecordSet!LOAN)
        GRP_OTHER = NumericVal(rsTemporaryRecordSet!Others)

        SSS_BASIS = NumericVal(rsTemporaryRecordSet!SSS_BASIS)
        PHEALTH_BASIS = NumericVal(rsTemporaryRecordSet!PHILHEALTH_BASIS)
        PAGIBIG_BASIS = NumericVal(rsTemporaryRecordSet!PAGIBIG_BASIS)
        TAX_BASIS = NumericVal(rsTemporaryRecordSet!TAX_BASIS)
        TAX_COMP = NumericVal(rsTemporaryRecordSet!TAX_COMPUTED)
        WORKING_DAYS = NumericVal(rsTemporaryRecordSet!WORKING_DAY)
        WORKING_HOURS = NumericVal(rsTemporaryRecordSet!WORKING_HOURS)
        AVERAGE_MONTH = NumericVal(rsTemporaryRecordSet!AVERAGE_MONTH)
        
        'date_from = (rsTemporaryRecordSet!adj_from)
        'date_to = (rsTemporaryRecordSet!adj_to)
    End If
    Set rsTemporaryRecordSet = Nothing
End Sub

Sub Process_Loans(rsLoanMas As ADODB.Recordset, EMPLIVIL, vEmployeeNo)
    Dim RSLOAN                                         As ADODB.Recordset
    Dim VTRANNO                                        As Integer
    Dim vtxtpaytype                                    As String
    Dim VACTNO                                         As String
    Dim VLOANDESCRIPTION                               As String
    Dim VLOANCODE                                      As String
    Dim VTEMPBALANCE                                   As Double
    Dim PAYMENT                                        As Double
    Dim VLOANBAL                                       As Double
    Dim VSMONDED                                       As Double
    Dim dedSalLoan                                     As Double
    Dim dedCalLoan                                     As Double
    Dim dedMPL                                         As Double
    Dim dedHLL                                         As Double
    Dim dedOTHER                                       As Double
    Dim rsLOANX                                        As ADODB.Recordset
    'VTRANNO = N2Str2Null(rsLoanMas!tranno)
    VTRANNO = N2Str2Zero(rsLoanMas!TRANNO)
    VSMONDED = N2Str2Zero(rsLoanMas!SMONTHLYDED)
    VLOANBAL = Round(N2Str2Zero(rsLoanMas!LoanBalance), 2)
    VACTNO = N2Str2Null(rsLoanMas!acctno)

Call gconDMIS.Execute("DELETE FROM HRMS_LOANMASDET WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(vEmployeeNo) & " AND TRANNO = " & N2Str2Null(VTRANNO) & " AND CUT_OFF = '" & CUTTOFF_CODE & "' AND PAY_MONTH = " & PAY_MONTH & " AND PAY_YEAR = " & PAY_YEAR & "")


    Set RSLOAN = New ADODB.Recordset
    Set RSLOAN = gconDMIS.Execute("SELECT SUM(AMOUNT) AS TOTALPAYMENT " & _
                                " FROM HRMS_LOANMASDET " & _
                                " WHERE EMPLEVEL = " & EMPLIVIL & _
                                " AND EMPNO = " & N2Str2Null(vEmployeeNo) & _
                                " AND TRANNO = " & N2Str2Null(VTRANNO) & "")

    If Not (RSLOAN.EOF And RSLOAN.BOF) Then
        gconDMIS.Execute "UPDATE HRMS_LOANMAS SET " & _
                       " LOANBALANCE = BEG_BAL-" & N2Str2Zero(RSLOAN!totalpayment) & _
                       " WHERE TRANNO = " & VTRANNO


        VTEMPBALANCE = VLOANBAL + N2Str2Zero(RSLOAN!totalpayment)
'        Call gconDMIS.Execute("DELETE FROM HRMS_LOANMASDET WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(vEmployeeNo) & " AND TRANNO = " & N2Str2Null(VTRANNO) & " AND CUT_OFF = '" & CUTTOFF_CODE & "' AND PAY_MONTH = " & PAY_MONTH & " AND PAY_YEAR = " & PAY_YEAR & "")

        Set rsLOANX = gconDMIS.Execute("SELECT LOANBALANCE FROM HRMS_LOANMAS WHERE TRANNO=" & VTRANNO)
        If Not rsLOANX.EOF Or Not rsLOANX.BOF Then
            VLOANBAL = Round(N2Str2Zero(rsLOANX!LoanBalance), 2)
        End If


    End If



    If VLOANBAL = 0 Then
        PAYMENT = (VSMONDED)
        GoTo skip1
    End If
    If VLOANBAL <= (VSMONDED) Then
        PAYMENT = VLOANBAL
    Else
        PAYMENT = (VSMONDED)
    End If
    'skip1:
    If N2Str2Zero(rsLoanMas!LoanBalance) >= 0 Then
        If Null2String(rsLoanMas!LOANTYPE) = "SSAL" Then    'SSS SALARY LOAN
            dedSalLoan = dedSalLoan + PAYMENT
            XdedSalLoan = dedSalLoan
            VLOANCODE = "'SSAL'"
        ElseIf Null2String(rsLoanMas!LOANTYPE) = "CSAL" Then    'SSS CALAMITY LOAN
            dedCalLoan = dedCalLoan + PAYMENT
            XdedCalLoan = dedCalLoan
            VLOANCODE = "'CSAL'"
        ElseIf Null2String(rsLoanMas!LOANTYPE) = "PSAL" Then    'PAG-IBIG SALARY LOAN
            dedMPL = dedCalLoan + PAYMENT
            XdedMPL = dedMPL
            VLOANCODE = "'PSAL'"
        ElseIf Null2String(rsLoanMas!LOANTYPE) = "HDMF" Then    'PAG-IBIG HDMF
            dedHLL = dedHLL + PAYMENT
            XdedHLL = dedHLL
            VLOANCODE = "'HDMF'"
        Else                                          'OTHER TYPE OF LOAN
            dedOTHER = dedOTHER + PAYMENT
            XdedOTHER = XdedOTHER + dedOTHER
            VLOANCODE = N2Str2Null(rsLoanMas!LOANTYPE)
        End If
        If vPayrollGeneratingStatus = "P" Then        'PAYROLL LOAN DEDUCTION:DATABASE UPDATES INSEERTS
            Dim VPAYTYPE                               As String
            VLOANDESCRIPTION = ""
            VLOANDESCRIPTION = "'LOAN PAYMENT " & Null2String(rsLoanMas!LOANTYPE) & "'"

            vtxtpaytype = Null2String(rsLoanMas!DEDUCTION_OPTION)

            gconDMIS.Execute "INSERT INTO HRMS_LOANMASDET " & _
                             "(EMPLEVEL, EMPNO, ACCTNO, AMOUNT,PAYTYPE,LOANDESCRIPTION, DEYT, LOANTYPE, TRANNO, CUT_OFF, PAY_MONTH, PAY_YEAR)" & _
                           " VALUES (" & EMPLIVIL & _
                             ", " & N2Str2Null(vEmployeeNo) & _
                             ", " & VACTNO & _
                             ", " & PAYMENT & _
                             ", " & vtxtpaytype & _
                             ", " & VLOANDESCRIPTION & _
                             ", '" & GENTO & _
                             "', " & VLOANCODE & _
                             ", " & N2Str2Null(VTRANNO) & _
                             ",'" & CUTTOFF_CODE & _
                             "'," & PAY_MONTH & _
                             "," & PAY_YEAR & ")"

            gconDMIS.Execute ("INSERT INTO HRMS_PAYROLL_DET (EMPLEVEL,EMPNO,TRANTYPE,DET_AMOUNT,ISADD,PAYPERIOD_FROM,PAYPERIOD_TO,DET_CODE,DET_DESC,TRANNO, CUT_OFF ,PAY_MONTH ,PAY_YEAR) VALUES(" _
                            & EMPLIVIL & ",'" & vEmployeeNo & "','L'," & PAYMENT & ",0,'" & GENFROM & "','" & GENTO & "'," & VLOANCODE & ",'" & GETLOAN_DESCRIPTION(Null2String(rsLoanMas!LOANTYPE)) & "'," & N2Str2Null(VTRANNO) & ",'" & CUTTOFF_CODE & "'," & PAY_MONTH & "," & PAY_YEAR & ")")

            Dim RSLOAN2                                As ADODB.Recordset
            Set RSLOAN2 = New ADODB.Recordset
            Set RSLOAN2 = gconDMIS.Execute("SELECT SUM(AMOUNT) AS TOTALPAYMENT2 " & _
                                         " FROM HRMS_LOANMASDET " & _
                                         " WHERE EMPLEVEL = " & EMPLIVIL & _
                                         " AND EMPNO = " & N2Str2Null(vEmployeeNo) & _
                                         " AND TRANNO = " & N2Str2Null(VTRANNO) & "")

            gconDMIS.Execute "UPDATE HRMS_LOANMAS SET " & _
                           " LOANBALANCE = BEG_BAL-" & N2Str2Zero(RSLOAN2!totalpayment2) & _
                           " WHERE TRANNO = " & VTRANNO
        End If

    End If
skip1:
End Sub

Sub FillGrid()
    Grid1.Rows = 1
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO " & _
                                    " WHERE  (DATEHIRED <= '" & Format(GENTO, "SHORT DATE") & "') " & _
                                    " AND (RESIGNED IS NULL OR RESIGNED >= '" & Format(GENFROM, "SHORT DATE") & "')AND Includthispayroll = 'Y' AND " & PROCESS_OPTION & "" & _
                                    " ORDER BY LASTNAME ASC")
    Grid1.Rows = 1
    If Not rsEMPINFO2.EOF And Not rsEMPINFO2.BOF Then
        rsEMPINFO2.MoveFirst
        While Not rsEMPINFO2.EOF
            Grid1.AddItem Null2String(rsEMPINFO2!EMPNO) & Chr(9) & Null2String(rsEMPINFO2!lastname) & ", " & Null2String(rsEMPINFO2!FIRSTNAME) & Chr(9) & "DELETE"
            rsEMPINFO2.MoveNext
        Wend
        
    Else
       MsgBox "No Employee! Please use the Employee Payroll Setup first to proceed!", vbInformation, "HRMS"
       Exit Sub
    End If
    Set rsEMPINFO2 = Nothing
End Sub

Sub InitGrid()
    With Grid1
        .Cols = 4
        .Column(0).Width = 50
        .Column(1).Width = 80
        .Column(2).Width = 250
        .Column(3).Width = 80
        .Cell(0, 0).Text = "L/N"
        .Cell(0, 1).Text = "EMPNO"
        .Cell(0, 2).Text = "EMPLOYEE NAME"
        .Cell(0, 3).Text = "OPTION"
        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(1).Alignment = cellCenterCenter
        .Column(3).Alignment = cellCenterCenter
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDone_Click()
    Unload Me
    frmHRMSGenerate.cmdCancel.Value = True
End Sub

Private Sub cmdGo_Click()
    If Function_Access(LOGID, "Acess_Process", "PROCESS GENERATE PAYROLL") = False Then Exit Sub
    If MsgQuestionBox("Is this the Final Payroll?", "Post Payroll") = True Then
        vPayrollGeneratingStatus = "P"
    Else
        vPayrollGeneratingStatus = "U"
    End If

    cmdGO.Visible = False
    cmdCancel.Visible = False

    Dim rsEmpInfo                                      As ADODB.Recordset
    Dim rsDailyMonitoring                              As ADODB.Recordset
    Dim rsPAYROLL                                      As ADODB.Recordset
    Dim rsLoanMas                                      As ADODB.Recordset
    Dim rsATMdet                                       As ADODB.Recordset
    Dim rsSSS                                          As ADODB.Recordset
    Dim rsPH                                           As ADODB.Recordset
    Dim rsPagIbig                                      As ADODB.Recordset
    Dim rsTIN                                          As ADODB.Recordset
    Dim rsPrevPayroll                                  As ADODB.Recordset
    Dim rsAllPrevPayroll                               As ADODB.Recordset

    Dim EMPBASICSALARY                                 As Double
    Dim EmpDailyRate                                   As Double
    Dim TotOvertime                                    As Double
    Dim TotHoliday                                     As Double
    Dim TotCommission                                  As Double
    Dim TotCommissionTax                               As Double
    Dim TotTaxableAdj                                  As Double
    Dim TotNonTaxableAdj                               As Double
    Dim TotSalaryAdvance                               As Double
    Dim TotAbsent                                      As Double
    Dim TotUndertime                                   As Double
    Dim TotOthers                                      As Double
    Dim TotTelBill                                     As Double
    Dim NUMDAYS                                        As Integer
    Dim SUWELDO                                        As Double
    Dim SUWELDOKINSE                                   As Double
    Dim SUWELDOTRIENTA                                 As Double
    Dim dedPAGIBIG                                     As Double
    Dim dedEmpPAGIBIG                                  As Double
    Dim dedSSS                                         As Double
    Dim dedEmpSSS                                      As Double
    Dim dedPhilHealth                                  As Double
    Dim dedEmpPhilhealth                               As Double
    Dim dedTIN                                         As Double
    Dim AssdedSSS                                      As Double
    Dim AssdedPhilHealth                               As Double
    Dim ASSUMED_BASIC                                  As Double
    Dim ASSUMED_COLA                                   As Double
    Dim ASSUMEDN0NTAXABLE                              As Double
    Dim ASSUMEDMONTHLY                                 As Double

    Dim dedSalLoan                                     As Double
    Dim dedCalLoan                                     As Double
    Dim dedPagSalLoan                                  As Double
    Dim dedMPL                                         As Double
    Dim dedHLL                                         As Double
    Dim dedBLL                                         As Double
    Dim dedOTHER                                       As Double

    Dim SalGross                                       As Double
    Dim TotTaxable                                     As Double
    Dim TotNonTaxable                                  As Double
    Dim TotTaxWheld                                    As Double
    Dim PersonalEx                                     As Double
    Dim NetTaxable                                     As Double
    Dim Taxdue                                         As Double
    Dim EMPLIVIL                                       As String
    Dim THIRTINT_MONTH_EXCESS                          As Double
    Dim DED_ADVANCE                                    As Double
    Dim PAYROLL_GROSS                                  As Double
    Dim DEDUCTION_BASIS                                As Double
    Dim vEmployeeNo                                    As String
    Dim vEmployeeTaxCode                               As String
    Dim vEmployeePayType                               As String
    'LOAN VAR
    Dim THEXFACTOR                                     As Integer
    Dim VTRANNO                                        As String
    Dim VSMONDED                                       As Double
    Dim VLOANBAL                                       As Double
    Dim VACTNO                                         As String
    Dim VLOANCODE                                      As String
    Dim PAYMENT                                        As Double
    Dim VTEMPBALANCE                                   As Double
    Dim VLOANDESCRIPTION                               As String
    'ALLOWANCE VAR
    Dim VEMPALLOWANCE                                  As Double
    Dim VCOMPUTEDALLOWANCE                             As Double

    I = 1
    While I < Grid1.Rows
        
        '*****************************************************

            If CUTTOFF_CODE = 1 Then
               CUTTOFF_CODE = 2
        
                If PAY_MONTH = 1 Then
                    PAY_MONTH = 12
                    PAY_YEAR = PAY_YEAR - 1
                Else
                    PAY_MONTH = PAY_MONTH - 1
                End If
        
            ElseIf CUTTOFF_CODE = 2 Then
                   CUTTOFF_CODE = 1
            End If

      '**************************************************
        
        Screen.MousePointer = 11
        Set rsEmpInfo = New ADODB.Recordset
        Set rsEmpInfo = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE EMPNO = '" & Grid1.Cell(I, 1).Text & "'")
        StoreDeductionSetValues rsEmpInfo!PayrollGroup

        vEmployeeNo = Null2String(rsEmpInfo!EMPNO)
        vEmployeeTaxCode = Null2String(rsEmpInfo!EXSTATUS)
        vEmployeePayType = Null2String(rsEmpInfo!EMPSTATUS)
        EMPLIVIL = N2Str2Null(rsEmpInfo!EMPLEVEL)
        EMPBASICSALARY = N2Str2Zero(rsEmpInfo!BASICSALARY)

        If Null2String(rsEmpInfo!EMPLEVEL) = "A" Or Null2String(rsEmpInfo!EMPLEVEL) = "C" Then
            GoTo skip2
        End If

        labName.Caption = RTrim(Null2String(rsEmpInfo!lastname)) & ", " & RTrim(Null2String(rsEmpInfo!FIRSTNAME)) & " " & RTrim(Null2String(rsEmpInfo!MIDDLENAME))

        NUMDAYS = 0

        'SELECT SALARY
        '=========================================================================================================
        If vEmployeePayType = "M" Then
            EmpDailyRate = Round((N2Str2Zero(rsEmpInfo!BASICSALARY) * 12) / WORKING_DAYS, 2)
        ElseIf vEmployeePayType = "D" Then
            EmpDailyRate = (N2Str2Zero(rsEmpInfo!BASICSALARY))
        End If

        If vEmployeePayType = "M" Then
            SUWELDOKINSE = EMPBASICSALARY / 2
            SUWELDOTRIENTA = EMPBASICSALARY
            If SUWELDOTRIENTA > 30000 Then
                THIRTINT_MONTH_EXCESS = SUWELDOTRIENTA - 30000
            Else
                THIRTINT_MONTH_EXCESS = 0
            End If
        Else
            Set rsDailyMonitoring = New ADODB.Recordset
            Set rsDailyMonitoring = gconDMIS.Execute("SELECT EMPNO, DEYT, PARTICULAR, ACTUAL, ENTERED " & _
                                                   " FROM HRMS_DAILYMONITORING " & _
                                                   " WHERE EMPLEVEL = " & EMPLIVIL & _
                                                   " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & _
                                                   " AND CUT_OFF = " & CUTTOFF_CODE & _
                                                   " AND PAY_MONTH = " & PAY_MONTH & _
                                                   " AND PAY_YEAR = " & PAY_YEAR & "")

            If Not rsDailyMonitoring.EOF And Not rsDailyMonitoring.BOF Then
                NUMDAYS = NumericVal(rsDailyMonitoring!entered)
            End If
            SUWELDOKINSE = NUMDAYS * EmpDailyRate
            SUWELDOTRIENTA = SUWELDOKINSE
            ASSUMEDMONTHLY = (EmpDailyRate * WORKING_DAYS) / 12
        End If

        'COMPUTE FOR OVERTIME
        '=========================================================================================================
        TotOvertime = 0
        TotOvertime = ComputeTotalOT(EMPLIVIL, vEmployeeNo, CUTTOFF_CODE, PAY_MONTH, PAY_YEAR)

        'COMPUTE FOR COMMISSION
        '=========================================================================================================
        TotCommission = 0
        TotCommissionTax = 0
        TotCommission = ComputeTotalCommission(EMPLIVIL, vEmployeeNo, CUTTOFF_CODE, PAY_MONTH, PAY_YEAR)
        TotCommissionTax = ComputeTotalCommissionTax(EMPLIVIL, vEmployeeNo, CUTTOFF_CODE, PAY_MONTH, PAY_YEAR)

        'COMPUTE FOR ADJUSTMENT
        '=========================================================================================================
        TotNonTaxableAdj = 0
        TotTaxableAdj = 0
        TotNonTaxableAdj = ComputeTotalNonTaxAdj(EMPLIVIL, vEmployeeNo, CUTTOFF_CODE, PAY_MONTH, PAY_YEAR)
        TotTaxableAdj = ComputeTotalTaxAdj(EMPLIVIL, vEmployeeNo, CUTTOFF_CODE, PAY_MONTH, PAY_YEAR)

        'COMPUTE FOR DEDUCTIONS
        '=========================================================================================================
        TotUndertime = 0
        TotAbsent = 0
        TotOthers = 0
        TotTelBill = 0

        TotUndertime = ComputeDeductionLateUndertime(EMPLIVIL, vEmployeeNo, CUTTOFF_CODE, PAY_MONTH, PAY_YEAR)
        TotAbsent = ComputeDeductionAbsences(EMPLIVIL, vEmployeeNo, CUTTOFF_CODE, PAY_MONTH, PAY_YEAR)
        TotOthers = ComputeDeductionOthers(EMPLIVIL, vEmployeeNo, CUTTOFF_CODE, PAY_MONTH, PAY_YEAR)

        'COMPUTE FOR GROSS
        '=========================================================================================================
        PAYROLL_GROSS = (SUWELDOKINSE) + TotCommission + TotOvertime + TotHoliday + TotNonTaxableAdj + TotTaxableAdj

        'COMPUTE FOR PREMIUM DEDUCTIONS (SSS/PHIC/PAGIBIG)
        '=========================================================================================================

        If GRP_SSS = 1 Or GRP_SSS = 2 Then
            dedSSS = GetFromSSSTable(SetDeductionBasis("SSS", vEmployeeNo, SSS_BASIS, GRP_SSS, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYEE")
            dedEmpSSS = GetFromSSSTable(SetDeductionBasis("SSS", vEmployeeNo, SSS_BASIS, GRP_SSS, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYER")
        Else
            dedSSS = GetFromSSSTable(SetDeductionBasis("SSS", vEmployeeNo, SSS_BASIS, GRP_SSS, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYEE")
            dedEmpSSS = GetFromSSSTable(SetDeductionBasis("SSS", vEmployeeNo, SSS_BASIS, GRP_SSS, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYER")
        End If

        If GRP_PHILHEALTH = 1 Or GRP_PHILHEALTH = 2 Then
            dedPhilHealth = GetFromPhilHealthTable(SetDeductionBasis("PHIC", vEmployeeNo, PHEALTH_BASIS, GRP_PHILHEALTH, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYEE")
            dedEmpPhilhealth = GetFromPhilHealthTable(SetDeductionBasis("PHIC", vEmployeeNo, PHEALTH_BASIS, GRP_PHILHEALTH, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYER")
        Else
            dedPhilHealth = GetFromPhilHealthTable(SetDeductionBasis("PHIC", vEmployeeNo, PHEALTH_BASIS, GRP_PHILHEALTH, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYEE")
            dedEmpPhilhealth = GetFromPhilHealthTable(SetDeductionBasis("PHIC", vEmployeeNo, PHEALTH_BASIS, GRP_PHILHEALTH, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYER")
        End If

        If GRP_PAGIBIG = 1 Or GRP_PAGIBIG = 2 Then
            dedPAGIBIG = GetFromPagIbigTable(SetDeductionBasis("PIF", vEmployeeNo, PAGIBIG_BASIS, GRP_PAGIBIG, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYEE")
            dedEmpPAGIBIG = GetFromPagIbigTable(SetDeductionBasis("PIF", vEmployeeNo, PAGIBIG_BASIS, GRP_PAGIBIG, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYER")
        Else
            dedPAGIBIG = GetFromPagIbigTable(SetDeductionBasis("PIF", vEmployeeNo, PAGIBIG_BASIS, GRP_PAGIBIG, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYEE") / 2
            dedEmpPAGIBIG = GetFromPagIbigTable(SetDeductionBasis("PIF", vEmployeeNo, PAGIBIG_BASIS, GRP_PAGIBIG, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj), "EMPLOYER") / 2
        End If

        'COMPUTE FOR TAX DEDUCTION
        '=========================================================================================================
        If GRP_TAX = 4 Then
            dedTIN = 0
        Else
            If TAX_COMP = 1 Then
                dedTIN = GetFromTaxTable(vEmployeeTaxCode, SetDeductionBasis("TAX", vEmployeeNo, TAX_BASIS, GRP_TAX, PAYROLL_GROSS, SUWELDOKINSE, TotUndertime + TotAbsent, TotOvertime + TotHoliday + TotTaxableAdj) - (dedSSS + dedPhilHealth + dedPAGIBIG), UCase(rsEmpInfo!payrolltype))
                If dedTIN = 1 Then
                    dedTIN = 0
                End If
            Else
                'COMPUTE TAX BASED ON ANNUALIZED INCOME
                TotTaxable = 0
                TotNonTaxable = 0
                TotTaxWheld = 0
                NetTaxable = 0
                Taxdue = 0
                PersonalEx = Personal_EX(Null2String(rsEmpInfo!EXSTATUS))
    
                '*** NEED TO REVISE THE ASSUMED PREMIUMS -- SHOULD BASE ON PAYROLL GROUP
                If Null2String(rsEmpInfo!EMPSTATUS) = "M" Then
                    ASSUMED_BASIC = Round(getAssumedBasic(CDate(GENTO), N2Str2Zero(EMPBASICSALARY)), 2)
                    ASSUMEDN0NTAXABLE = Round(getAssumedPH(CDate(GENTO), N2Str2Zero(EMPBASICSALARY)) + getAssumedSSS(CDate(GENTO), N2Str2Zero(EMPBASICSALARY)) + getAssumedPagIbig(CDate(GENTO), N2Str2Zero(EMPBASICSALARY)), 2)
                Else
                    ASSUMED_BASIC = Round(getAssumedBasic(CDate(GENTO), ASSUMEDMONTHLY), 2)
                    ASSUMEDN0NTAXABLE = Round(getAssumedPH(CDate(GENTO), ASSUMEDMONTHLY) + getAssumedSSS(CDate(GENTO), ASSUMEDMONTHLY) + getAssumedPagIbig(CDate(GENTO), ASSUMEDMONTHLY), 2)
                End If
    
                '*** NEED TO REVISE THE ANNUALIZED COMPUTATION -- SHOULD BASE ON PAYROLL GROUP
                Set rsAllPrevPayroll = New ADODB.Recordset
                If TAX_BASIS = 1 Then Set rsAllPrevPayroll = gconDMIS.Execute("Select (SUM(rate)+SUM(overtime)+SUM(holiday)+SUM(taxableadj)+SUM(commission)) as TOTALTAXABLE,(SUM(sssE)+SUM(philhealthE)+SUM(pagibig)) as TOTALNONTAXABLE,SUM(tax)+SUM(commissiontax) as TOTALTAXWHELD from HRMS_Payroll where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND paydateto < '" & CDate(GENFROM) & "' AND year(paydateto) = " & YEAR(CDate(GENFROM)))
                If TAX_BASIS = 2 Then Set rsAllPrevPayroll = gconDMIS.Execute("Select (SUM(rate)) as TOTALTAXABLE,(SUM(sssE)+SUM(philhealthE)+SUM(pagibig)) as TOTALNONTAXABLE,SUM(tax)+SUM(commissiontax) as TOTALTAXWHELD from HRMS_Payroll where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND paydateto < '" & CDate(GENFROM) & "' AND year(paydateto) = " & YEAR(CDate(GENFROM)))
                If TAX_BASIS = 3 Then Set rsAllPrevPayroll = gconDMIS.Execute("Select (SUM(rate))-(SUM(undertime)+SUM(absent)) as TOTALTAXABLE,(SUM(sssE)+SUM(philhealthE)+SUM(pagibig)) as TOTALNONTAXABLE,SUM(tax)+SUM(commissiontax) as TOTALTAXWHELD from HRMS_Payroll where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND paydateto < '" & CDate(GENFROM) & "' AND year(paydateto) = " & YEAR(CDate(GENFROM)))
                If TAX_BASIS = 4 Then Set rsAllPrevPayroll = gconDMIS.Execute("Select (SUM(rate)+SUM(overtime)+SUM(holiday))-(SUM(undertime)+SUM(absent)) as TOTALTAXABLE,(SUM(sssE)+SUM(philhealthE)+SUM(pagibig)) as TOTALNONTAXABLE,SUM(tax)+SUM(commissiontax) as TOTALTAXWHELD from HRMS_Payroll where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND paydateto < '" & CDate(GENFROM) & "' AND year(paydateto) = " & YEAR(CDate(GENFROM)))
                If Not rsAllPrevPayroll.EOF And Not rsAllPrevPayroll.BOF Then
                    TotTaxable = TotTaxable + Round(N2Str2Zero(rsAllPrevPayroll!TOTALTAXABLE), 2)
                    TotNonTaxable = TotNonTaxable + Round(N2Str2Zero(rsAllPrevPayroll!TOTALNONTAXABLE), 2)
                    TotTaxWheld = TotTaxWheld + Round(N2Str2Zero(rsAllPrevPayroll!TOTALTAXWHELD), 2)
                End If
    
                TotTaxable = Round((TotTaxable + TotCommission + THIRTINT_MONTH_EXCESS + ASSUMED_BASIC + ASSUMED_COLA) + ((SUWELDOKINSE + TotOvertime + TotHoliday + TotTaxableAdj) - (TotUndertime + TotAbsent)), 2)
                NetTaxable = Round(TotTaxable - (TotNonTaxable + ASSUMEDN0NTAXABLE + dedSSS + dedPhilHealth + dedPAGIBIG + PersonalEx), 2)
                Taxdue = Round(Tax_Due(NetTaxable) - (TotCommissionTax + TotTaxWheld), 2)
    
                If Taxdue <= 0 Then
                    dedTIN = 0
                Else
                    If CUTTOFF_CODE = 1 Then
                        dedTIN = Taxdue / (((12 - MONTH(GENTO)) * 2) + 2)
                    Else
                        dedTIN = Taxdue / (((12 - MONTH(GENTO)) * 2) + 1)
                    End If
    
                End If
            End If
        End If
        'COMPUTE FOR SALARY ADVANCES
        '=========================================================================================================
        TotSalaryAdvance = 0
        TotSalaryAdvance = ComputeTotalAdvance(EMPLIVIL, vEmployeeNo, CUTTOFF_CODE, PAY_MONTH, PAY_YEAR)

        'COMPUTE FOR EMPLOYEE LOANS
        '=========================================================================================================
        XdedSalLoan = 0
        XdedCalLoan = 0
        XdedMPL = 0
        XdedHLL = 0
        XdedBLL = 0
        XdedOTHER = 0
        PAYMENT = 0
       
        Set rsLoanMas = gconDMIS.Execute("SELECT * FROM HRMS_LOANMAS " & _
                                       " WHERE EMPLEVEL = " & EMPLIVIL & _
                                       " AND EMPNO = " & N2Str2Null(vEmployeeNo) & _
                                       " AND (LOANBALANCE >= 0) " & _
                                       " AND DATESTARTED <= '" & Format(GENTO, "SHORT DATE") & "'" & _
                                       " ORDER BY DATEGRANTED DESC")
        If Not rsLoanMas.EOF And Not rsLoanMas.BOF Then
            rsLoanMas.MoveFirst
            Do While Not rsLoanMas.EOF
                If COMPANY_CODE = "HARI" Then
                    If Null2String(rsLoanMas!ISACTIVE) <> "N" Then
                        If rsLoanMas("DEDUCTION_OPTION") = 3 Then
                            Process_Loans rsLoanMas, EMPLIVIL, vEmployeeNo
                        ElseIf rsLoanMas("DEDUCTION_OPTION") = CUTTOFF_CODE Then
                            Process_Loans rsLoanMas, EMPLIVIL, vEmployeeNo
                        End If
                    End If
                    rsLoanMas.MoveNext
                Else
                    If rsLoanMas("DEDUCTION_OPTION") = 3 Then
                        Process_Loans rsLoanMas, EMPLIVIL, vEmployeeNo
                    ElseIf rsLoanMas("DEDUCTION_OPTION") = CUTTOFF_CODE Then
                        Process_Loans rsLoanMas, EMPLIVIL, vEmployeeNo
                    End If
                    rsLoanMas.MoveNext

                End If
            Loop
        End If

        'COMPUTE FOR ALLOWANCE: HARI STANDARD
        '=========================================================================================================
        VCOMPUTEDALLOWANCE = 0
        VEMPALLOWANCE = 0
        VEMPALLOWANCE = N2Str2Zero(rsEmpInfo!ALLOWANCE)
        VCOMPUTEDALLOWANCE = ComputeAllowance(EMPLIVIL, vEmployeeNo, CUTTOFF_CODE, PAY_MONTH, PAY_YEAR, VEMPALLOWANCE)
        If Null2String(rsEmpInfo!payrolltype) = "Monthly Base" Then
            VCOMPUTEDALLOWANCE = VEMPALLOWANCE
            SUWELDOKINSE = SUWELDOKINSE * 2
        End If

        'SAVE PAYROLL
        '=========================================================================================================

        SUWELDO = Round((SUWELDOKINSE + TotOvertime + TotHoliday + TotTaxableAdj + TotNonTaxableAdj + TotCommission), 2) - _
                  Round((Round(dedPhilHealth, 2) + Round(dedSSS, 2) + Round(dedPAGIBIG, 2) + Round(dedTIN, 2) + Round(XdedSalLoan, 2) + Round(XdedOTHER, 2) + Round(XdedCalLoan, 2) + Round(XdedMPL, 2) + Round(XdedHLL, 2) + Round(dedBLL, 2) + Round(TotUndertime, 2) + Round(TotTelBill, 2) + Round(TotAbsent, 2) + Round(TotOthers, 2) + Round(TotSalaryAdvance, 2)), 2)

        '=========================================================================================================

        
        'REVERSE THE CUT OFF
        '=========================================================================================================
                If CUTTOFF_CODE = 1 Then
                   CUTTOFF_CODE = 2
            
                ElseIf CUTTOFF_CODE = 2 Then
                       CUTTOFF_CODE = 1
            
                    If PAY_MONTH = 12 Then
                       PAY_MONTH = 1
                       PAY_YEAR = PAY_YEAR + 1
                    Else
                       PAY_MONTH = PAY_MONTH + 1
                      End If
                End If
            '=========================================================================================================
        
        
        
        
        gconDMIS.Execute "INSERT INTO HRMS_PAYROLL " & _
                         "(EMPLEVEL, EMPNO,TAXCODE, RATE, MONTHLYRATE, DAILYRATE, NDAYS, OVERTIME, HOLIDAY, COMMISSION, COMMISSIONTAX, TAXABLEADJ, NONTAXABLEADJ, GROSS, UNDERTIME, SSSE, SSSR, PHILHEALTHE, PHILHEALTHR, PAGIBIG, TAX, SSSSALLOAN, SSSCALLOAN, PAGSALLOAN, PAGHDMFLOAN, OTHERLOAN, ABSENT, TELBILL, OTHERS, PAYDATEFROM, PAYDATETO, NETPAY, PAYROLLSTATUS, ADVANCE, CUT_OFF, PAY_MONTH, PAY_YEAR, ALLOWANCE)" & _
                       " VALUES (" & EMPLIVIL & _
                         "," & N2Str2Null(rsEmpInfo!EMPNO) & ", " & N2Str2Null(rsEmpInfo!EXSTATUS) & _
                         ", " & Round(SUWELDOKINSE, 2) & ", " & Round(EMPBASICSALARY, 2) & _
                         ", " & Round(EmpDailyRate, 2) & ", " & NUMDAYS & _
                         ", " & Round(TotOvertime, 2) & ", " & Round(TotHoliday, 2) & _
                         ", " & Round(TotCommission, 2) & ", " & Round(TotCommissionTax, 2) & _
                         ", " & Round(TotTaxableAdj, 2) & ", " & Round(TotNonTaxableAdj, 2) & _
                         ", " & Round(SUWELDOKINSE + TotOvertime + TotHoliday + TotTaxableAdj + TotNonTaxableAdj + TotCommission, 2) & _
                         ", " & Round(TotUndertime, 2) & ", " & Round(dedSSS, 2) & _
                         ", " & Round(dedEmpSSS, 2) & ", " & Round(dedPhilHealth, 2) & _
                         ", " & Round(dedEmpPhilhealth, 2) & ", " & Round(dedPAGIBIG, 2) & _
                         ", " & Round(dedTIN, 2) & ", " & Round(XdedSalLoan, 2) & _
                         ", " & Round(XdedCalLoan, 2) & ", " & Round(XdedMPL, 2) & _
                         ", " & Round(XdedHLL, 2) & ", " & XdedOTHER & _
                         ", " & Round(TotAbsent, 2) & ", " & Round(TotTelBill, 2) & _
                         ", " & Round(TotOthers, 2) & ", '" & GENFROM & _
                         "', '" & GENTO & "', " & Round(SUWELDO, 2) & _
                         ",  '" & vPayrollGeneratingStatus & "'," & TotSalaryAdvance & _
                         ",  '" & CUTTOFF_CODE & "'," & PAY_MONTH & _
                         "," & PAY_YEAR & "," & VCOMPUTEDALLOWANCE & ")"

        'CHECK IF FINAL PAYROLL THEN IF YES SAVE PREMIUM DEDUCTIONS IN DATABASE
        '=========================================================================================================
        If vPayrollGeneratingStatus = "P" Then
            'SSS
            Set rsSSS = New ADODB.Recordset
            Set rsSSS = gconDMIS.Execute("SELECT * FROM HRMS_SSS " & _
                                       " WHERE EMPLEVEL = " & EMPLIVIL & _
                                       " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO))

            If rsSSS.EOF And rsSSS.BOF Then
                gconDMIS.Execute "INSERT INTO HRMS_SSS " & _
                    "(EMPLEVEL, EMPNO, SSSNO, DATESTART, EMPLOYEESHARE, EMPLOYERSHARE, LASTDATECONT)" & _
                    " VALUES (" & EMPLIVIL & _
                    ", " & N2Str2Null(rsEmpInfo!EMPNO) & _
                    ", " & N2Str2Null(rsEmpInfo!SSSNO) & _
                    ", '" & GENFROM & _
                    "', " & Round(dedSSS, 2) & _
                    ", " & Round(dedEmpSSS, 2) & _
                    ", '" & GENTO & "')"
            Else
                gconDMIS.Execute "UPDATE HRMS_SSS SET" & _
                               " EMPLOYEESHARE = " & Round(dedSSS, 2) & "," & _
                               " EMPLOYERSHARE = " & Round(dedEmpSSS, 2) & "," & _
                               " LASTDATECONT = '" & GENTO & "'" & _
                               " WHERE EMPLEVEL = " & EMPLIVIL & _
                               " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO)
            End If

            Set rsSSS = New ADODB.Recordset
            Set rsSSS = gconDMIS.Execute("SELECT * FROM HRMS_SSS " & _
                                       " WHERE EMPLEVEL = " & EMPLIVIL & _
                                       " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO))

            If Not rsSSS.EOF And Not rsSSS.BOF Then
                gconDMIS.Execute "INSERT INTO HRMS_SSSDET " & _
                                 "(EMPLEVEL, AYDI, DEYT, EMPNO, EMPLOYEEAMOUNT, EMPLOYERAMOUNT, CUT_OFF, PAY_MONTH, PAY_YEAR)" & _
                               " VALUES (" & EMPLIVIL & _
                                 "," & rsSSS!aydi & _
                                 ",'" & GENTO & _
                                 "'," & N2Str2Null(rsEmpInfo!EMPNO) & _
                                 "," & Round(dedSSS, 2) & _
                                 "," & Round(dedEmpSSS, 2) & _
                                 ",'" & CUTTOFF_CODE & _
                                 "'," & PAY_MONTH & _
                                 "," & PAY_YEAR & ")"
            End If

            'PHIC
            Set rsPH = New ADODB.Recordset
            Set rsPH = gconDMIS.Execute("SELECT * FROM HRMS_PHILHEALTH " & _
                                      " WHERE EMPLEVEL = " & EMPLIVIL & _
                                      " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO))

            If rsPH.EOF And rsPH.BOF Then
                gconDMIS.Execute "INSERT INTO HRMS_PHILHEALTH " & _
                    "(EMPLEVEL, EMPNO, PHNO, DATESTART, EMPLOYEESHARE, EMPLOYERSHARE, LASTDATECONT)" & _
                    " VALUES (" & EMPLIVIL & _
                    ", " & N2Str2Null(rsEmpInfo!EMPNO) & _
                    ", " & N2Str2Null(rsEmpInfo!SSSNO) & _
                    ", '" & GENFROM & _
                    "', " & Round(dedPhilHealth, 2) & _
                    ", " & Round(dedEmpPhilhealth, 2) & _
                    ", '" & GENTO & "')"
            Else
                gconDMIS.Execute "UPDATE HRMS_PHILHEALTH SET" & _
                    " EMPLOYEESHARE = " & Round(dedPhilHealth, 2) & "," & _
                    " EMPLOYERSHARE = " & Round(dedEmpPhilhealth, 2) & "," & _
                    " LASTDATECONT = '" & GENTO & "'" & _
                    " WHERE EMPLEVEL = " & EMPLIVIL & _
                    " AND PHNO = " & N2Str2Null(rsEmpInfo!SSSNO)
            End If

            Set rsPH = New ADODB.Recordset
            Set rsPH = gconDMIS.Execute("SELECT * FROM HRMS_PHILHEALTH WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO))
            If Not rsPH.EOF And Not rsPH.BOF Then
                gconDMIS.Execute "INSERT INTO HRMS_PHILHEALTHDET " & _
                    "(EMPLEVEL, AYDI, DEYT, EMPNO, EMPLOYEEAMOUNT, EMPLOYERAMOUNT, CUT_OFF, PAY_MONTH, PAY_YEAR)" & _
                    " VALUES (" & EMPLIVIL & _
                    ", " & rsPH!aydi & _
                    ", '" & GENTO & _
                    "', " & N2Str2Null(rsEmpInfo!EMPNO) & _
                    ", " & Round(dedPhilHealth, 2) & _
                    ", " & Round(dedEmpPhilhealth, 2) & _
                    ", '" & CUTTOFF_CODE & _
                    "', " & PAY_MONTH & _
                    ", " & PAY_YEAR & ")"
            End If

            'PAGIBIG
            Set rsPagIbig = New ADODB.Recordset
            Set rsPagIbig = gconDMIS.Execute("SELECT * FROM HRMS_PAGIBIG WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO))
            If rsPagIbig.EOF And rsPagIbig.BOF Then
                gconDMIS.Execute "INSERT INTO HRMS_PAGIBIG " & _
                    "(EMPLEVEL, EMPNO, PAGIBIGNO, DATESTART, EMPLOYEESHARE, EMPLOYERSHARE, LASTDATECONT)" & _
                    " VALUES (" & EMPLIVIL & _
                    ", " & N2Str2Null(rsEmpInfo!EMPNO) & _
                    ", " & N2Str2Null(rsEmpInfo!SSSNO) & _
                    ", '" & GENFROM & _
                    "', " & Round(dedPAGIBIG, 2) & _
                    ", " & Round(dedEmpPAGIBIG, 2) & _
                    ", '" & GENTO & "')"
            Else
                gconDMIS.Execute "UPDATE HRMS_PAGIBIG SET" & _
                    " EMPLOYEESHARE = " & Round(dedPAGIBIG, 2) & "," & _
                    " EMPLOYERSHARE = " & Round(dedEmpPAGIBIG, 2) & "," & _
                    " LASTDATECONT = '" & GENTO & "'" & _
                    " WHERE EMPLEVEL = " & EMPLIVIL & _
                    " AND PAGIBIGNO = " & N2Str2Null(rsEmpInfo!SSSNO)
            End If

            Set rsPagIbig = New ADODB.Recordset
            Set rsPagIbig = gconDMIS.Execute("SELECT * FROM HRMS_PAGIBIG WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO))
            If Not rsPagIbig.EOF And Not rsPagIbig.BOF Then
                gconDMIS.Execute "INSERT INTO HRMS_PAGIBIGDET " & _
                    "(EMPLEVEL, AYDI, DEYT, EMPNO, EMPLOYEEAMOUNT, EMPLOYERAMOUNT, CUT_OFF, PAY_MONTH, PAY_YEAR)" & _
                    " VALUES (" & EMPLIVIL & _
                    ", " & rsPagIbig!aydi & _
                    ", '" & GENTO & _
                    "', " & N2Str2Null(rsEmpInfo!EMPNO) & _
                    ", " & Round(dedPAGIBIG, 2) & _
                    ", " & Round(dedEmpPAGIBIG, 2) & _
                    ", '" & CUTTOFF_CODE & _
                    "', " & PAY_MONTH & _
                    ", " & PAY_YEAR & ")"
            End If
        End If

        '=========================================================================================================
        'CHECK IF FINAL PAYROLL THEN IF YES SAVE TAX DEDUCTION IN DATABASE
        If vPayrollGeneratingStatus = "P" Then
            Set rsTIN = New ADODB.Recordset
            Set rsTIN = gconDMIS.Execute("SELECT * FROM HRMS_TIN WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO))
            If rsTIN.EOF And rsTIN.BOF Then
                If IsEmpty(dedTIN) = True Then
                    dedTIN = 0
                End If
                gconDMIS.Execute "INSERT INTO HRMS_TIN " & _
                    "(EMPLEVEL, EMPNO, TINNO, DATESTART, DEDUCTION, LASTDATECONT)" & _
                    " VALUES (" & EMPLIVIL & _
                    ", " & N2Str2Null(rsEmpInfo!EMPNO) & _
                    ", " & N2Str2Null(rsEmpInfo!tinno) & _
                    ", '" & GENFROM & _
                    "', " & dedTIN & _
                    ", '" & GENTO & "')"
            Else
                If IsEmpty(dedTIN) = True Then
                    dedTIN = 0
                End If
                gconDMIS.Execute "UPDATE HRMS_TIN SET" & _
                    " DEDUCTION = " & dedTIN & "," & _
                    " LASTDATECONT = '" & GENTO & "'" & _
                    " WHERE EMPLEVEL = " & EMPLIVIL & _
                    " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO)
            End If

            Set rsTIN = New ADODB.Recordset
            Set rsTIN = gconDMIS.Execute("SELECT * FROM HRMS_TIN WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO))
            If Not (rsTIN.BOF And rsTIN.EOF) Then
                gconDMIS.Execute "INSERT INTO HRMS_TINDET " & _
                    "(EMPLEVEL, AYDI, EMPNO, DEYT, AMOUNT, CUT_OFF, PAY_MONTH, PAY_YEAR)" & _
                    " VALUES (" & EMPLIVIL & _
                    "," & rsTIN!aydi & _
                    ", " & N2Str2Null(rsEmpInfo!EMPNO) & _
                    ", '" & GENTO & _
                    "', " & dedTIN & _
                    ", '" & CUTTOFF_CODE & _
                    "', " & PAY_MONTH & _
                    ", " & PAY_YEAR & ")"
            End If
        End If

        '=========================================================================================================
        'CHECK IF FINAL PAYROLL THEN IF YES SAVE ATM DETAILS IN DATABASE
        If vPayrollGeneratingStatus = "P" Then
            Dim hash                                   As Double
            hash = 0
            hash = GETHASHVALUE(Null2String(N2Str2Null(rsEmpInfo!ACCOUNTNO)), Round(SUWELDO + VCOMPUTEDALLOWANCE, 2))

            gconDMIS.Execute "INSERT INTO HRMS_ATMDET " & _
                "(EMPLEVEL, ACCTNO, EMPNO, ATMID, DEYT, NETAMOUNT, CUT_OFF, PAY_MONTH, PAY_YEAR, HOR_HAS) " & _
                " VALUES (" & EMPLIVIL & _
                ", " & N2Str2Null(rsEmpInfo!ACCOUNTNO) & _
                ", " & N2Str2Null(rsEmpInfo!EMPNO) & _
                ", " & rsEmpInfo!ID & _
                ", '" & GENTO & _
                "', " & Round(SUWELDO + VCOMPUTEDALLOWANCE, 2) & _
                ", '" & CUTTOFF_CODE & _
                "', " & PAY_MONTH & _
                ", " & PAY_YEAR & _
                ", " & Round(hash, 2) & ")"

        End If
        '=========================================================================================================
skip2:
1000
        I = I + 1
        gauProgress.Value = (I / Grid1.Rows) * 100
        lblPercent.Caption = Int(gauProgress.Value) & "%"
        DoEvents
    Wend
    LogAudit "G", "GENERATE PAYROLL", GENFROM & "-" & GENTO
    labName.Caption = ""
    Screen.MousePointer = 0
    cmdDone.Visible = True
    Exit Sub

Errorcode:
    MsgBox Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    DrawXPCtl Me
    InitGrid
    FillGrid
    labName.Caption = ""
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub Grid1_DblClick()
    If Grid1.ActiveCell.Col = 3 Then
        Grid1.RemoveItem (Grid1.ActiveCell.Row)
        Grid1.Refresh
    End If
End Sub


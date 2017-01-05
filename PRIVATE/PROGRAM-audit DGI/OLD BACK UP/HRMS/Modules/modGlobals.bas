Attribute VB_Name = "modHRMSGlobals"
Option Explicit
Public OVERWRAYT                             As Boolean
Public LOGCODE, LOGNAME, LOGLEVEL, LOGTIME, LOGDATE As String
Attribute LOGNAME.VB_VarUserMemId = 1073741825
Attribute LOGLEVEL.VB_VarUserMemId = 1073741825
Attribute LOGTIME.VB_VarUserMemId = 1073741825
Attribute LOGDATE.VB_VarUserMemId = 1073741825
Public LEDGERSHOW, EMPINFOSHOW, HEADEMPINFOSHOW, CASHOW As Boolean
Attribute LEDGERSHOW.VB_VarUserMemId = 1073741830
Attribute EMPINFOSHOW.VB_VarUserMemId = 1073741830
Attribute HEADEMPINFOSHOW.VB_VarUserMemId = 1073741830
Attribute CASHOW.VB_VarUserMemId = 1073741830
Public GENFROM, GENTO, HEADOREMP, EMP_TYPE   As String
Attribute GENFROM.VB_VarUserMemId = 1073741834
Attribute GENTO.VB_VarUserMemId = 1073741834
Attribute HEADOREMP.VB_VarUserMemId = 1073741834
Attribute EMP_TYPE.VB_VarUserMemId = 1073741834
Public Maxwiz, AccessCNT                     As Long
Attribute Maxwiz.VB_VarUserMemId = 1073741840
Attribute AccessCNT.VB_VarUserMemId = 1073741840
Public wizVar                                As wizEnc
Attribute wizVar.VB_VarUserMemId = 1073741842
Public CryptVar                              As Crypto
Public EmpInfoEmpno                          As Object
Public FormYearlyRequest                     As String
Attribute FormYearlyRequest.VB_VarUserMemId = 1073741845
Public MODULENAME
Public PREPARED_BY, CHECKED_BY, GENERAL_MANAGER, APPROVED_BY, ACCOUNT_NO, BANK_NAME, BANK_LOCATION, BANK_MANAGER, SECRETARY, NOTED_BY As String
Attribute PREPARED_BY.VB_VarUserMemId = 1073741849
Public EMPLOYEE_NO                           As String
Attribute EMPLOYEE_NO.VB_VarUserMemId = 1073741858
Public COMPANY_TIN, COMPANY_NAME, COMPANY_ADDRESS
Public PREPARED_BY_DESIGNATION, APPROVED_BY_DESIGNATION As String
Attribute PREPARED_BY_DESIGNATION.VB_VarUserMemId = 1073741859
Attribute APPROVED_BY_DESIGNATION.VB_VarUserMemId = 1073741859
Public PROCESS_OPTION                        As String
Attribute PROCESS_OPTION.VB_VarUserMemId = 1073741861
Public IMPNO                                 As String
Attribute IMPNO.VB_VarUserMemId = 1073741862
Public OVERTIME_CODES                        As String
Attribute OVERTIME_CODES.VB_VarUserMemId = 1073741863
Public OVERTIME_RATE                         As Double
Attribute OVERTIME_RATE.VB_VarUserMemId = 1073741864
'UPDATE BY : MJP 10-06-07 --------------------------------------------------------------------------------
'NEW PAYROLL FORMAT
    Public PAYROLLCODE_FROM1                     As Integer
    Public PAYROLLCODE_FROM2                     As Integer
    Public PAYROLLCODE_TO1                       As Integer
    Public PAYROLLCODE_TO2                       As Integer
    Public PAYROLL_BASE                          As String
    Public PAYROLL_NO_OF_DAYS                    As Integer
    Public PAYROLL_CODE                          As Integer
'UPDATE BY : MJP 10-06-07 --------------------------------------------------------------------------------


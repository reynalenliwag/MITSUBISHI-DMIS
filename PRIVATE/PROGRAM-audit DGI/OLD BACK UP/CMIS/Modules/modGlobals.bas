Attribute VB_Name = "modCMISGlobals"
Option Explicit
Public LOGCODE, LOGNAME, LOGLEVEL, LOGDATE, LOGTIME As String
Attribute LOGNAME.VB_VarUserMemId = 1073741824
Attribute LOGLEVEL.VB_VarUserMemId = 1073741824
Attribute LOGDATE.VB_VarUserMemId = 1073741824
Attribute LOGTIME.VB_VarUserMemId = 1073741824

Public wizVar, CryptVar  As Object
Attribute wizVar.VB_VarUserMemId = 1073741830
Attribute CryptVar.VB_VarUserMemId = 1073741830
Public AccessCNT         As Integer
Attribute AccessCNT.VB_VarUserMemId = 1073741832
Public SUBTOTAL_ADDON    As Double
Attribute SUBTOTAL_ADDON.VB_VarUserMemId = 1073741833
Public BOOKTYPE          As String
Attribute BOOKTYPE.VB_VarUserMemId = 1073741834
Public VAT_OR As Integer
Public OR_VAT_NONVAT     As String
Attribute OR_VAT_NONVAT.VB_VarUserMemId = 1073741835
Public CANCEL_OR_VAT_NONVAT As String
Attribute CANCEL_OR_VAT_NONVAT.VB_VarUserMemId = 1073741836
Public BRANCH_CODE, OR_NUMBER_GLOBAL As String
Attribute BRANCH_CODE.VB_VarUserMemId = 1073741837
Attribute OR_NUMBER_GLOBAL.VB_VarUserMemId = 1073741837

Public RECEIPTS_AMOUNT, AMOUNT_TENDERED, CHANGE_DUE As Double
Attribute RECEIPTS_AMOUNT.VB_VarUserMemId = 1073741839
Attribute AMOUNT_TENDERED.VB_VarUserMemId = 1073741839
Attribute CHANGE_DUE.VB_VarUserMemId = 1073741839
Public MODE_OF_PAYMENT, INVOICE_DETAIL_TYPE, INVOICE_DETAIL_TRANNO As String
Attribute MODE_OF_PAYMENT.VB_VarUserMemId = 1073741842
Attribute INVOICE_DETAIL_TYPE.VB_VarUserMemId = 1073741842
Attribute INVOICE_DETAIL_TRANNO.VB_VarUserMemId = 1073741842

Public Const MAX_PETTYFUND = 0
Public Const MAX_LTOFUND = 0
Public Const CHANGE_FUND = 3000
Public CURRENT_CUTOFF_DATE As String
Attribute CURRENT_CUTOFF_DATE.VB_VarUserMemId = 1073741845
Public CASHPOSITION_CUTOFF_DATE As String
Attribute CASHPOSITION_CUTOFF_DATE.VB_VarUserMemId = 1073741846
Public IsLTOIsPettyCash  As String
Attribute IsLTOIsPettyCash.VB_VarUserMemId = 1073741847
Public CMIS_Report_Range As String
Attribute CMIS_Report_Range.VB_VarUserMemId = 1073741848
Public CMIS_Type_Of_Report As String
Attribute CMIS_Type_Of_Report.VB_VarUserMemId = 1073741849
Public ROSHOW, ESTISHOW  As String
Attribute ROSHOW.VB_VarUserMemId = 1073741850
Attribute ESTISHOW.VB_VarUserMemId = 1073741850
Public CURRENT_CUST_CODE As String
Attribute CURRENT_CUST_CODE.VB_VarUserMemId = 1073741852
Public CASH_OPTIONS      As String
Attribute CASH_OPTIONS.VB_VarUserMemId = 1073741853
Public TYPE_ON_HAND      As String
Attribute TYPE_ON_HAND.VB_VarUserMemId = 1073741854

Public PERIODMONTH, PERIODYEAR As Integer
Attribute PERIODMONTH.VB_VarUserMemId = 1073741855

Public PREPARED_BY, CHECKED_BY, APPROVED_BY, ACCOUNT_NO, BANK_NAME, BANK_LOCATION, BANK_MANAGER, SECRETARY, NOTED_BY As String
Public EMPLOYEE_NO       As String
Public PREPARED_BY_DESIGNATION, APPROVED_BY_DESIGNATION As String
Public Const MODULENAME = "CMIS"

Public PreparedBy, IssuedBy, CheckedBy, ApprovedBy, GeneralManager, NotedBy


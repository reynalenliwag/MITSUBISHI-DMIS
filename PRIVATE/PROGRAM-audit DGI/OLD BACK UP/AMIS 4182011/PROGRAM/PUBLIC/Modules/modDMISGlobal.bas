Attribute VB_Name = "modDMISGlobal"
Option Explicit
'UPDATED BY: JUN/ARNOLD-------------------------------------------------
'DATE UPDATED: 06-11-2009
'DESCRIPTION: DECLARE GLOBAL DUE TO PRINTING PURPOSE FOR CUSTOMER LEDGER
Public xBALANCE                                        As Double
Public BEG_BALANCE_DATE                                As Date
Public SelectEntity                                    As String
'UPDATED BY: JUN/ARNOLD-------------------------------------------------
Public rEndingBalance                                  As Double
Public QC_MODULE_ON                                    As String
Public LOGCODE As String, LOGNAME As String, LOGLEVEL As String, LOGTIME As String, LOGDATE As String
Attribute LOGNAME.VB_VarUserMemId = 1073741829
Attribute LOGLEVEL.VB_VarUserMemId = 1073741829
Attribute LOGTIME.VB_VarUserMemId = 1073741829
Attribute LOGDATE.VB_VarUserMemId = 1073741829
Public FROM_APPOINTMENT                                As String
Attribute FROM_APPOINTMENT.VB_VarUserMemId = 1073741834
Public RECEIVED_FROM_PO                                As String
Attribute RECEIVED_FROM_PO.VB_VarUserMemId = 1073741835
Public BIR_RELIEF_Connection                           As String
Attribute BIR_RELIEF_Connection.VB_VarUserMemId = 1073741836

Public wizVar, CryptVar                                As Object
Attribute wizVar.VB_VarUserMemId = 1073741837
Attribute CryptVar.VB_VarUserMemId = 1073741837
Public AccessCNT                                       As Integer
Attribute AccessCNT.VB_VarUserMemId = 1073741839
Public AcctCodeArray(5000), AcctNameArray(5000)        As String
Attribute AcctCodeArray.VB_VarUserMemId = 1073741840

Public BILANG, SEARCH_TAB                              As Long
Attribute BILANG.VB_VarUserMemId = 1073741842
Attribute SEARCH_TAB.VB_VarUserMemId = 1073741842
Public JOURNALTYPE                                     As String
Attribute JOURNALTYPE.VB_VarUserMemId = 1073741844
Public REPORT_RANGETYPE                                As String
Attribute REPORT_RANGETYPE.VB_VarUserMemId = 1073741845
Public REPORT_EXPENSETYPE                              As String
Attribute REPORT_EXPENSETYPE.VB_VarUserMemId = 1073741846
Public CUST_LEDGER_TYPE                                As String
Attribute CUST_LEDGER_TYPE.VB_VarUserMemId = 1073741847
'Public REPORT_AR As Strings
Public Report_AR                                       As String
Attribute Report_AR.VB_VarUserMemId = 1073741848
Public REPORT_AP                                       As String
Attribute REPORT_AP.VB_VarUserMemId = 1073741849
Public REFRESH_ACCOUNT                                 As Boolean
Attribute REFRESH_ACCOUNT.VB_VarUserMemId = 1073741850
Public CURRENT_CUSCODE                                 As String
Attribute CURRENT_CUSCODE.VB_VarUserMemId = 1073741851
Public CURRENT_VENDORCODE                              As String
Attribute CURRENT_VENDORCODE.VB_VarUserMemId = 1073741852
Public xSELECTED                                       As String

Public INVOICE_Type                                    As String
Attribute INVOICE_Type.VB_VarUserMemId = 1073741853
Public MYOB_JTYPE                                      As String
Attribute MYOB_JTYPE.VB_VarUserMemId = 1073741854
Public Const VAT_RATE = 12
Public EXTRACT_TYPE                                    As String
Attribute EXTRACT_TYPE.VB_VarUserMemId = 1073741855

Public CASH_SALES                                      As String
Attribute CASH_SALES.VB_VarUserMemId = 1073741856
Public CHARGE_SALES                                    As String
Attribute CHARGE_SALES.VB_VarUserMemId = 1073741857
Public CASH_DISCOUNT                                   As String
Attribute CASH_DISCOUNT.VB_VarUserMemId = 1073741858
Public CHARGE_DISCOUNT                                 As String
Attribute CHARGE_DISCOUNT.VB_VarUserMemId = 1073741859
Public CASH_COSTOFSALES                                As String
Attribute CASH_COSTOFSALES.VB_VarUserMemId = 1073741860
Public CHARGE_COSTOFSALES                              As String
Attribute CHARGE_COSTOFSALES.VB_VarUserMemId = 1073741861
Public OPERATIONAL_EXPENSE                             As String
Attribute OPERATIONAL_EXPENSE.VB_VarUserMemId = 1073741862
Public ADMIN_EXPENSE                                   As String
Attribute ADMIN_EXPENSE.VB_VarUserMemId = 1073741863
Public OTHER_INCOME                                    As String
Attribute OTHER_INCOME.VB_VarUserMemId = 1073741864
Public OTHER_EXPENSE                                   As String
Attribute OTHER_EXPENSE.VB_VarUserMemId = 1073741865
Public CURRENT_ASSET                                   As String
Attribute CURRENT_ASSET.VB_VarUserMemId = 1073741866
Public TAX_CREDITS                                     As String
Attribute TAX_CREDITS.VB_VarUserMemId = 1073741867
Public PROPERTY_EQUIPMENT                              As String
Attribute PROPERTY_EQUIPMENT.VB_VarUserMemId = 1073741868
Public ACCUMULATED_DEPRECIATION                        As String
Attribute ACCUMULATED_DEPRECIATION.VB_VarUserMemId = 1073741869
Public OTHER_ASSET                                     As String
Attribute OTHER_ASSET.VB_VarUserMemId = 1073741870

Public CUSCODE                                         As String
Attribute CUSCODE.VB_VarUserMemId = 1073741871
Public MODULENAME
Attribute MODULENAME.VB_VarUserMemId = 1073741872

Public INVOICE_DETAIL_TYPE, INVOICE_DETAIL_TRANNO      As String
Attribute INVOICE_DETAIL_TYPE.VB_VarUserMemId = 1073741873
Attribute INVOICE_DETAIL_TRANNO.VB_VarUserMemId = 1073741873

Public OVERWRAYT                                       As Boolean
Attribute OVERWRAYT.VB_VarUserMemId = 1073741875
Public GVD_DATABASE_PATH                               As String
Attribute GVD_DATABASE_PATH.VB_VarUserMemId = 1073741876
Public SKIN_PATH                                       As String
Attribute SKIN_PATH.VB_VarUserMemId = 1073741877
Public NEYM                                            As String
Attribute NEYM.VB_VarUserMemId = 1073741878
Public ADRES                                           As String
Attribute ADRES.VB_VarUserMemId = 1073741879
Public TELLNO                                          As String
Attribute TELLNO.VB_VarUserMemId = 1073741880
Public PURLASTNEYM                                     As String
Attribute PURLASTNEYM.VB_VarUserMemId = 1073741881
Public PURFIRSTNEYM                                    As String
Attribute PURFIRSTNEYM.VB_VarUserMemId = 1073741882
Public PURMIDDLE                                       As String
Attribute PURMIDDLE.VB_VarUserMemId = 1073741883
Public PRODUCTNO                                       As String
Attribute PRODUCTNO.VB_VarUserMemId = 1073741884
Public LASTNEYM                                        As String
Attribute LASTNEYM.VB_VarUserMemId = 1073741885
Public FIRSTNEYM                                       As String
Attribute FIRSTNEYM.VB_VarUserMemId = 1073741886
Public MIDDLE                                          As String
Attribute MIDDLE.VB_VarUserMemId = 1073741887
Public Add_o_Edit                                      As String
Attribute Add_o_Edit.VB_VarUserMemId = 1073741888
Public EMPINFOSHOW                                     As Boolean
Attribute EMPINFOSHOW.VB_VarUserMemId = 1073741889

Public SAECODE                                         As String
Attribute SAECODE.VB_VarUserMemId = 1073741890
Public LOGSAE                                          As String
Attribute LOGSAE.VB_VarUserMemId = 1073741891
Public SAENAME                                         As String
Attribute SAENAME.VB_VarUserMemId = 1073741892

Public INVOICENO                                       As String
Attribute INVOICENO.VB_VarUserMemId = 1073741893
Public InvoiceType                                     As String
Attribute InvoiceType.VB_VarUserMemId = 1073741894

Public rKeyPublicension(1000)                          As Integer
Attribute rKeyPublicension.VB_VarUserMemId = 1073741895
Public EncryptoFile(100000)                            As String
Attribute EncryptoFile.VB_VarUserMemId = 1073741896
Public CryptoKey                                       As Variant
Attribute CryptoKey.VB_VarUserMemId = 1073741897
Public Maxwiz                                          As Long
Attribute Maxwiz.VB_VarUserMemId = 1073741898
Public SEARCH_BY                                       As String
Attribute SEARCH_BY.VB_VarUserMemId = 1073741899
Public VInoArray(3000)                                 As String
Attribute VInoArray.VB_VarUserMemId = 1073741900
Public VICusNamArray(3000)                             As String
Attribute VICusNamArray.VB_VarUserMemId = 1073741901
Public CusVInoArray(3000)                              As String
Attribute CusVInoArray.VB_VarUserMemId = 1073741902
Public CusNamArray(3000)                               As String
Attribute CusNamArray.VB_VarUserMemId = 1073741903
Public CusNameArray(3000)                              As String
Attribute CusNameArray.VB_VarUserMemId = 1073741904
Public CusProdNoArray(3000)                            As String
Attribute CusProdNoArray.VB_VarUserMemId = 1073741905
Public CusProdNoArray2(3000)                           As String
Attribute CusProdNoArray2.VB_VarUserMemId = 1073741906
Public CusCodeArray(3000)                              As String
Attribute CusCodeArray.VB_VarUserMemId = 1073741907
Public BILANG2                                         As Long
Attribute BILANG2.VB_VarUserMemId = 1073741908
Public BILANG_CusName                                  As Long
Attribute BILANG_CusName.VB_VarUserMemId = 1073741909
Public BILANG_CusName2                                 As Long
Attribute BILANG_CusName2.VB_VarUserMemId = 1073741910
Public CUST_REPT_TYPE                                  As String
Attribute CUST_REPT_TYPE.VB_VarUserMemId = 1073741911
Public PARTS_ISSUED_TO_CUSTOMER_TYPE                   As String
Attribute PARTS_ISSUED_TO_CUSTOMER_TYPE.VB_VarUserMemId = 1073741912

Public Const WorkTimeStart                             As String = "8:00 AM"
Public Const WorkTimeEnd                               As String = "5:00 PM"
Public FILE_GRAPH                                      As String
Attribute FILE_GRAPH.VB_VarUserMemId = 1073741913
Public fromParts                                       As Boolean

'SIGNATORIES AND ADDRESSES
Public PREPARED_BY, CHECKED_BY, GENERAL_MANAGER, APPROVED_BY, ACCOUNT_NO, BANK_NAME, BANK_LOCATION, BANK_MANAGER, SECRETARY, NOTED_BY As String
Attribute PREPARED_BY.VB_VarUserMemId = 1073741914
Attribute CHECKED_BY.VB_VarUserMemId = 1073741914
Attribute GENERAL_MANAGER.VB_VarUserMemId = 1073741914
Attribute APPROVED_BY.VB_VarUserMemId = 1073741914
Attribute ACCOUNT_NO.VB_VarUserMemId = 1073741914
Attribute BANK_NAME.VB_VarUserMemId = 1073741914
Attribute BANK_LOCATION.VB_VarUserMemId = 1073741914
Attribute BANK_MANAGER.VB_VarUserMemId = 1073741914
Attribute SECRETARY.VB_VarUserMemId = 1073741914
Attribute NOTED_BY.VB_VarUserMemId = 1073741914
Public PreparedBy, ApprovedBy, CheckedBy, SalesDispatcher, GeneralManager, DeliveredBy, FinancingManager As String
Attribute PreparedBy.VB_VarUserMemId = 1073741924
Attribute ApprovedBy.VB_VarUserMemId = 1073741924
Attribute CheckedBy.VB_VarUserMemId = 1073741924
Attribute SalesDispatcher.VB_VarUserMemId = 1073741924
Attribute GeneralManager.VB_VarUserMemId = 1073741924
Attribute DeliveredBy.VB_VarUserMemId = 1073741924
Attribute FinancingManager.VB_VarUserMemId = 1073741924
'SIGNATORIES AND ADDRESSES
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'UPDATED BY: JUN
'DATE UPDATED: 09142008
'DESCRIPTION: SIGNATORIES DESIGNATION
'SIGNATORIES DESIGNATION
Public PreparedByDesig, CheckedByDesig, SalesApprovedDesig, SalesDispatcherDesig, GeneralManagerDesig, DeliveredByDesig, FinancingManagerDesig As String
Attribute PreparedByDesig.VB_VarUserMemId = 1073741931
Attribute CheckedByDesig.VB_VarUserMemId = 1073741931
Attribute SalesApprovedDesig.VB_VarUserMemId = 1073741931
Attribute SalesDispatcherDesig.VB_VarUserMemId = 1073741931
Attribute GeneralManagerDesig.VB_VarUserMemId = 1073741931
Attribute DeliveredByDesig.VB_VarUserMemId = 1073741931
Attribute FinancingManagerDesig.VB_VarUserMemId = 1073741931
'-------------------------------------------------------------------------------------------------------------------------------------------------------


Public gconINVENTORY                                   As ADODB.Connection
Attribute gconINVENTORY.VB_VarUserMemId = 1073741938

Public PROC_TYPE                                       As String
Attribute PROC_TYPE.VB_VarUserMemId = 1073741939
Public WAREHOUSETYPE As String, COUNTERTYPE As String, MAT_COUNTERTYPE As String, ORDERTYPE As String, VPAMCOR As String, BIR_YearEnd As String, ISSREPTYPE As String, ORDER_REPORT As String
Attribute WAREHOUSETYPE.VB_VarUserMemId = 1073741940
Attribute COUNTERTYPE.VB_VarUserMemId = 1073741940
Attribute MAT_COUNTERTYPE.VB_VarUserMemId = 1073741940
Attribute ORDERTYPE.VB_VarUserMemId = 1073741940
Attribute VPAMCOR.VB_VarUserMemId = 1073741940
Attribute BIR_YearEnd.VB_VarUserMemId = 1073741940
Attribute ISSREPTYPE.VB_VarUserMemId = 1073741940
Attribute ORDER_REPORT.VB_VarUserMemId = 1073741940
Public BIRDATA_Connection, BIR_DATABASE_PATH           As String
Attribute BIRDATA_Connection.VB_VarUserMemId = 1073741948
Attribute BIR_DATABASE_PATH.VB_VarUserMemId = 1073741948
Public CSMS_PARTSQUERY                                 As Boolean
Attribute CSMS_PARTSQUERY.VB_VarUserMemId = 1073741950

Public TOTJOBAMT, TOTJOBDISC, TOTJOBTAX                As Double
Attribute TOTJOBAMT.VB_VarUserMemId = 1073741951
Attribute TOTJOBDISC.VB_VarUserMemId = 1073741951
Attribute TOTJOBTAX.VB_VarUserMemId = 1073741951
Public TOTPARTSAMT, TOTPARTSDISC, TOTPARTSTAX          As Double
Attribute TOTPARTSAMT.VB_VarUserMemId = 1073741954
Attribute TOTPARTSDISC.VB_VarUserMemId = 1073741954
Attribute TOTPARTSTAX.VB_VarUserMemId = 1073741954
Public TOTMATAMT, TOTMATDISC, TOTMATTAX                As Double
Attribute TOTMATAMT.VB_VarUserMemId = 1073741957
Attribute TOTMATDISC.VB_VarUserMemId = 1073741957
Attribute TOTMATTAX.VB_VarUserMemId = 1073741957
Public TOTACCAMT, TOTACCDISC, TOTACCTAX                As Double
Attribute TOTACCAMT.VB_VarUserMemId = 1073741960
Attribute TOTACCDISC.VB_VarUserMemId = 1073741960
Attribute TOTACCTAX.VB_VarUserMemId = 1073741960
Public DNPIDFrom, DNPIDTo                              As Long
Attribute DNPIDFrom.VB_VarUserMemId = 1073741963
Attribute DNPIDTo.VB_VarUserMemId = 1073741963


Public gconBIR_RELIEF                                  As ADODB.Connection
Attribute gconBIR_RELIEF.VB_VarUserMemId = 1073741965
Public PARTSQUERY                                      As Integer
Attribute PARTSQUERY.VB_VarUserMemId = 1073741966
Public rKeyDimension(1000)                             As Integer
Attribute rKeyDimension.VB_VarUserMemId = 1073741967

Public VoiceMsg                                        As Boolean
Attribute VoiceMsg.VB_VarUserMemId = 1073741968
Public STOCK_TYPE                                      As String
Attribute STOCK_TYPE.VB_VarUserMemId = 1073741969
Public PRR_REPORT                                      As String
Attribute PRR_REPORT.VB_VarUserMemId = 1073741970
Public FORECASTING_BUTTON_CLICK                        As Integer
Attribute FORECASTING_BUTTON_CLICK.VB_VarUserMemId = 1073741971
Public PRR_BUTTON_CLICK                                As Integer
Attribute PRR_BUTTON_CLICK.VB_VarUserMemId = 1073741972


Public Const MAX_ISS_LINE = 14

Public Const PESO_VALUE_FOR_ONE = 300
Public Const PESO_VALUE_FOR_TWO = 3000
Public Const PESO_VALUE_FOR_THREE = 6000

Public Const RANK_FAST_MOVING = "A"
Public Const RANK_MEDIUM_MOVING = "B"
Public Const RANK_SLOW_MOVING = "C"
Public Const RANK_NON_MOVING = "D"
Public Const RANK_NEW_PARTS = "E"

Public Const PARTS_MARK_UP_FROM_DNP = 1.32
Public Const PARTS_SSTOCK_NO_MONTHS = 2

Public Y_REGRESSION_INTERVAL                           As Double
Attribute Y_REGRESSION_INTERVAL.VB_VarUserMemId = 1073741973
Public Const X_MEAN_INTERVAL = 1

Public Const HARI_LEAD_TIME = 1.25
Public Const HARI_ORDER_FREQUENCY = 1.5

Public Const M_COMPANY_CODE = "MULTIPLE"

Public ROSHOW                                          As Boolean
Attribute ROSHOW.VB_VarUserMemId = 1073741974
Public ESTISHOW                                        As Boolean
Attribute ESTISHOW.VB_VarUserMemId = 1073741975
Public ESTIKCNT                                        As Integer
Attribute ESTIKCNT.VB_VarUserMemId = 1073741976

Public RO_OR_ESTI_OR_PART                              As String
Attribute RO_OR_ESTI_OR_PART.VB_VarUserMemId = 1073741977

Public TOTJOBDISCVAL                                   As Double
Attribute TOTJOBDISCVAL.VB_VarUserMemId = 1073741978
Public TOTPARTSDISCVAL                                 As Double
Attribute TOTPARTSDISCVAL.VB_VarUserMemId = 1073741979
Public TOTMATDISCVAL                                   As Double
Attribute TOTMATDISCVAL.VB_VarUserMemId = 1073741980
Public TOTACCDISCVAL                                   As Double
Attribute TOTACCDISCVAL.VB_VarUserMemId = 1073741981



Public SearchBy, SEARCHCUSTOMERNAME, SEARCHPLATENO     As String
Attribute SearchBy.VB_VarUserMemId = 1073741982
Attribute SEARCHCUSTOMERNAME.VB_VarUserMemId = 1073741982
Attribute SEARCHPLATENO.VB_VarUserMemId = 1073741982

Public QUESTION_TEST                                   As String
Attribute QUESTION_TEST.VB_VarUserMemId = 1073741985

Public EDIT_RO                                         As String
Attribute EDIT_RO.VB_VarUserMemId = 1073741986



Public SUBTOTAL_ADDON                                  As Double
Attribute SUBTOTAL_ADDON.VB_VarUserMemId = 1073741987
Public BOOKTYPE                                        As String
Attribute BOOKTYPE.VB_VarUserMemId = 1073741988
Public VAT_OR                                          As Integer
Attribute VAT_OR.VB_VarUserMemId = 1073741989
Public OR_VAT_NONVAT                                   As String
Attribute OR_VAT_NONVAT.VB_VarUserMemId = 1073741990
Public CANCEL_OR_VAT_NONVAT                            As String
Attribute CANCEL_OR_VAT_NONVAT.VB_VarUserMemId = 1073741991
Public BRANCH_CODE, OR_NUMBER_GLOBAL                   As String
Attribute BRANCH_CODE.VB_VarUserMemId = 1073741992
Attribute OR_NUMBER_GLOBAL.VB_VarUserMemId = 1073741992

Public FINAL_CASH                                      As Double
Public RECEIPTS_BALANCE                                As Double
Public RECEIPTS_AMOUNT, AMOUNT_TENDERED, CHANGE_DUE    As Double
Attribute RECEIPTS_AMOUNT.VB_VarUserMemId = 1073741994
Attribute AMOUNT_TENDERED.VB_VarUserMemId = 1073741994
Attribute CHANGE_DUE.VB_VarUserMemId = 1073741994
Public MODE_OF_PAYMENT                                 As String
Attribute MODE_OF_PAYMENT.VB_VarUserMemId = 1073741997
Public TYPE_OF_PAYMENT                                 As String

Public Const MAX_PETTYFUND = 0
Public Const MAX_LTOFUND = 0
Public Const CHANGE_FUND = 3000

Public RE_OPEN_CUTOFF                                  As String
Attribute RE_OPEN_CUTOFF.VB_VarUserMemId = 1073741998
Public PREVIOUS_CUTOFF                                 As String
Attribute PREVIOUS_CUTOFF.VB_VarUserMemId = 1073741999

Public CURRENT_CUTOFF_DATE                             As String
Attribute CURRENT_CUTOFF_DATE.VB_VarUserMemId = 1073742000
Public CASHPOSITION_CUTOFF_DATE                        As String
Attribute CASHPOSITION_CUTOFF_DATE.VB_VarUserMemId = 1073742001
Public IsLTOIsPettyCash                                As String
Attribute IsLTOIsPettyCash.VB_VarUserMemId = 1073742002
Public CMIS_Report_Range                               As String
Attribute CMIS_Report_Range.VB_VarUserMemId = 1073742003
Public CMIS_Type_Of_Report                             As String
Attribute CMIS_Type_Of_Report.VB_VarUserMemId = 1073742004
Public CURRENT_CUST_CODE                               As String
Attribute CURRENT_CUST_CODE.VB_VarUserMemId = 1073742005
Public CASH_OPTIONS                                    As String
Attribute CASH_OPTIONS.VB_VarUserMemId = 1073742006
Public TYPE_ON_HAND                                    As String
Attribute TYPE_ON_HAND.VB_VarUserMemId = 1073742007

Public PERIODMONTH, PERIODYEAR                         As Integer
Attribute PERIODMONTH.VB_VarUserMemId = 1073742008
Attribute PERIODYEAR.VB_VarUserMemId = 1073742008

Public EMPLOYEE_NO                                     As String
Attribute EMPLOYEE_NO.VB_VarUserMemId = 1073742010
Public PREPARED_BY_DESIGNATION, APPROVED_BY_DESIGNATION As String
Attribute PREPARED_BY_DESIGNATION.VB_VarUserMemId = 1073742011
Attribute APPROVED_BY_DESIGNATION.VB_VarUserMemId = 1073742011

Public IssuedBy, NotedBy
Attribute IssuedBy.VB_VarUserMemId = 1073742013
Attribute NotedBy.VB_VarUserMemId = 1073742013

Public LEDGERSHOW, HEADEMPINFOSHOW, CASHOW             As Boolean
Attribute LEDGERSHOW.VB_VarUserMemId = 1073742015
Attribute HEADEMPINFOSHOW.VB_VarUserMemId = 1073742015
Attribute CASHOW.VB_VarUserMemId = 1073742015
Public GENFROM, GENTO, HEADOREMP, EMP_TYPE             As String
Attribute GENFROM.VB_VarUserMemId = 1073742018
Attribute GENTO.VB_VarUserMemId = 1073742018
Attribute HEADOREMP.VB_VarUserMemId = 1073742018
Attribute EMP_TYPE.VB_VarUserMemId = 1073742018

Public EmpInfoEmpno                                    As Object
Attribute EmpInfoEmpno.VB_VarUserMemId = 1073742022
Public FormYearlyRequest                               As String
Attribute FormYearlyRequest.VB_VarUserMemId = 1073742023

Public SQL_STATEMENT                                   As String
Attribute SQL_STATEMENT.VB_VarUserMemId = 1073742024
Public PROCESS_OPTION                                  As String
Attribute PROCESS_OPTION.VB_VarUserMemId = 1073742025
Public IMPNO                                           As String
Attribute IMPNO.VB_VarUserMemId = 1073742026

Public OVERTIME_CODES                                  As String
Attribute OVERTIME_CODES.VB_VarUserMemId = 1073742027
Public OVERTIME_RATE                                   As Double
Attribute OVERTIME_RATE.VB_VarUserMemId = 1073742028

Public PAYROLLCODE_FROM1                               As Integer
Attribute PAYROLLCODE_FROM1.VB_VarUserMemId = 1073742029
Public PAYROLLCODE_FROM2                               As Integer
Attribute PAYROLLCODE_FROM2.VB_VarUserMemId = 1073742030
Public PAYROLLCODE_TO1                                 As Integer
Attribute PAYROLLCODE_TO1.VB_VarUserMemId = 1073742031
Public PAYROLLCODE_TO2                                 As Integer
Attribute PAYROLLCODE_TO2.VB_VarUserMemId = 1073742032
Public PAYROLL_BASE                                    As String
Attribute PAYROLL_BASE.VB_VarUserMemId = 1073742033
Public PAYROLL_NO_OF_DAYS                              As Integer
Attribute PAYROLL_NO_OF_DAYS.VB_VarUserMemId = 1073742034
Public PAYROLL_CODE                                    As Integer
Attribute PAYROLL_CODE.VB_VarUserMemId = 1073742035

Public CUTTOFF_CODE                                    As String
Attribute CUTTOFF_CODE.VB_VarUserMemId = 1073742036
Public PAY_MONTH                                       As Integer
Attribute PAY_MONTH.VB_VarUserMemId = 1073742037
Public PAY_YEAR                                        As Integer
Attribute PAY_YEAR.VB_VarUserMemId = 1073742038
Public DEDUCTION_OPTION                                As String
Attribute DEDUCTION_OPTION.VB_VarUserMemId = 1073742039
Public RE_PROCESS_CUT_OFF                              As Boolean
Attribute RE_PROCESS_CUT_OFF.VB_VarUserMemId = 1073742040

Public xPAIDFOR                                        As String
Attribute xPAIDFOR.VB_VarUserMemId = 1073742041
Public vREFERENCENO                                    As String
Attribute vREFERENCENO.VB_VarUserMemId = 1073742042
Public xBankCode                                       As String
Attribute xBankCode.VB_VarUserMemId = 1073742043
Public SelectCustomer                                  As String
Attribute SelectCustomer.VB_VarUserMemId = 1073742044
Public TranType                                        As String
Public xPAYEE_NAME                                     As String


Function IS_IN_AMIS(DMISProcessType As String, DMISTranType As String, DMISTranno As String, Optional IF_VAT As Boolean) As Boolean
    Dim rsAMISJournal                                  As ADODB.Recordset
    IS_IN_AMIS = False
    Dim DMISInvoiceType                                As String
    If DMISProcessType = "RECEIVED" Then
        Set rsAMISJournal = New ADODB.Recordset
        Set rsAMISJournal = gconDMIS.Execute("Select InvoiceType,InvoiceNo from AMIS_Journal_HD Where Jtype = 'APJ' and InvoiceType = '" & DMISTranType & "' AND InvoiceNo = '" & DMISTranno & "'")
        If Not rsAMISJournal.EOF And Not rsAMISJournal.BOF Then
            IS_IN_AMIS = True
        End If
    End If
    If DMISProcessType = "PAYMENT" Then
        Set rsAMISJournal = New ADODB.Recordset
        If IF_VAT = True Then
            Set rsAMISJournal = gconDMIS.Execute("Select InvoiceType,InvoiceNo from AMIS_Journal_HD Where Jtype = 'CRJ' AND InvoiceNo = '" & DMISTranno & "'")
        Else
            Set rsAMISJournal = gconDMIS.Execute("Select InvoiceType,InvoiceNo from AMIS_Journal_HD Where Jtype = 'CRJ' AND left(InvoiceNo,2) = 'NV' AND right(InvoiceNo,6) = '" & DMISTranno & "'")
        End If
        If Not rsAMISJournal.EOF And Not rsAMISJournal.BOF Then
            IS_IN_AMIS = True
        End If
    End If
    If DMISProcessType = "INVOICE" Then
        If DMISTranType = "PARTS" Then DMISInvoiceType = "PI"
        If DMISTranType = "ACCESSORIES" Then DMISInvoiceType = "AI"
        If DMISTranType = "MATERIALS" Then DMISInvoiceType = "MI"
        If DMISTranType = "SERVICE" Then DMISInvoiceType = "SI"
        If DMISTranType = "VEHICLES" Then DMISInvoiceType = "VI"
        If DMISTranType = "" Then DMISInvoiceType = "NULL"
        Set rsAMISJournal = New ADODB.Recordset
        Set rsAMISJournal = gconDMIS.Execute("Select InvoiceType,InvoiceNo from AMIS_Journal_HD Where Jtype = 'APJ' and InvoiceType = '" & DMISInvoiceType & "' AND InvoiceNo = '" & DMISTranno & "'")
        If Not rsAMISJournal.EOF And Not rsAMISJournal.BOF Then
            IS_IN_AMIS = True
        End If
    End If
End Function

Function getJTYPE(XINVOICE As String, xInvoiceType As String, Optional xARTAG As String) As String
'Update by BTT:11/21/2008
    Dim RS                                             As New ADODB.Recordset
    If xARTAG = "" Then
        Set RS = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_HD WHERE INVOICENO ='" & XINVOICE & "' and INVOICETYPE ='" & xInvoiceType & "'")
    ElseIf Left(xARTAG, 5) = "11-02" Then
        Set RS = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_HD WHERE INVOICENO ='" & XINVOICE & "' and INVOICETYPE ='" & xInvoiceType & "' AND JTYPE='SJ'")
    Else
        Set RS = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_HD WHERE INVOICENO ='" & XINVOICE & "' and INVOICETYPE ='" & xInvoiceType & "' AND JTYPE='CRJ'")
    End If
    If Not (RS.EOF And RS.BOF) Then
        getJTYPE = Null2String(RS!jtype)
    Else
        getJTYPE = "NULL"
    End If
    Set RS = Nothing
End Function

Function FormExist(XXX As String)

    Dim FRM                                            As Form

    For Each FRM In Forms
        If (UCase(FRM.Name) = UCase(XXX)) Then
            FormExist = True
        End If
    Next
    Set FRM = Nothing
End Function

Sub SaveLogFile()
    Dim FileName                                       As String
    'FileName = LOGDATE & time &
    Open AMIS_REPORT_PATH & Format(LOGDATE, "MMDDYY") & Format(Time, "HHMMAM/PM") & "-ErrorLog" & ".txt" For Append As #1
    Print #1, Err.DESCRIPTION
    Close #1
End Sub

Sub DMIS_VERSION()
    On Error GoTo FALSEUSERS
    'UPDATED BY : ACL
    'DATE       : 02022011
    'DESCRIPTION: TO CHECK LATEST VERSION OF APPLICATION
    Dim rsALLPROFILE                                   As ADODB.Recordset
    Dim rsUSERNAME                                     As ADODB.Recordset
    Dim SQL                                            As String
    Dim SQL1                                           As String
    Dim SQL2                                           As String
    Dim SQL3                                           As String
    Dim SQL4                                           As String
    
    If COMPANY_CODE <> "" Then
        If COMPANY_CODE = COMPANY_VERSION Then
            SQL = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_RAMS_USERS' AND COLUMN_NAME='USER_NAME') " & vbCrLf
            SQL = SQL & "SELECT USER_NAME FROM ALL_RAMS_USERS"
            Set rsUSERNAME = gconACCESS.Execute(SQL)
            If Not rsUSERNAME.EOF And Not rsUSERNAME.BOF Then
            Else
FALSEUSERS:
                SQL1 = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_RAMS_USERS' AND COLUMN_NAME='USERNAME') " & vbCrLf
                SQL1 = SQL1 & "EXEC SP_RENAME 'ALL_RAMS_USERS.USERNAME','USER_NAME','COLUMN'"
                gconACCESS.Execute SQL1
    
                SQL2 = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_RAMS_USERS' AND COLUMN_NAME='USER_NAME') " & vbCrLf
                SQL2 = "ALTER VIEW ALL_VW_RAMS_PACCESS " & vbCrLf
                SQL2 = SQL2 & "AS " & vbCrLf
                SQL2 = SQL2 & "SELECT USERID,USER_NAME,PASSWORD AS USERPASS,USERGROUP AS LOGLEVEL, USERCODE, LOCK " & vbCrLf
                SQL2 = SQL2 & "From DBO.ALL_RAMS_USERS"
                gconACCESS.Execute SQL2
    
                SQL3 = "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_PROFILE' AND COLUMN_NAME='VERSION') " & vbCrLf
                SQL3 = SQL3 & "ALTER TABLE ALL_PROFILE " & vbCrLf
                SQL3 = SQL3 & " ADD VERSION INT"
                gconACCESS.Execute SQL3
    
                SQL4 = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_RAMS_USERS' AND COLUMN_NAME='USER_NAME') " & vbCrLf
                SQL4 = "ALTER VIEW ALL_VW_USERACESS " & vbCrLf
                SQL4 = SQL4 & "AS" & vbCrLf
                SQL4 = SQL4 & "SELECT  ARU.USERID,ALL_RAMS_USERS.USER_NAME,ARU.MODULEID,ARM.MAINMODULENAME,ARM.DESCRIPTIONS,ARM.MODULE_TYPE, " & vbCrLf
                SQL4 = SQL4 & "ARU.ACESS_ADD,ARU.ACESS_EDIT,ARU.ACESS_DELETE,ARU.ACESS_VIEW,ARU.ACESS_PRINT,ARU.ACESS_PROCESS,ARU.ACESS_SYSTEM,ARU.ACESS_POST, " & vbCrLf
                SQL4 = SQL4 & "ARU.ACESS_UNPOST , ARU.ACESS_CANCELENTRY " & vbCrLf
                SQL4 = SQL4 & "FROM ALL_RAMS_USERSACESS AS ARU INNER JOIN " & vbCrLf
                SQL4 = SQL4 & "ALL_RAMS_MODULES AS ARM ON ARU.MODULEID = ARM.MODULEID INNER JOIN " & vbCrLf
                SQL4 = SQL4 & "ALL_RAMS_USERS ON ARU.USERID = ALL_RAMS_USERS.USERID"
                gconACCESS.Execute SQL4
            End If
    
            Set rsALLPROFILE = New ADODB.Recordset
            rsALLPROFILE.Open "SELECT ISNULL(VERSION,0) AS VERSION FROM ALL_PROFILE WHERE MODULENAME='" & App.EXEName & "'", gconACCESS, adOpenForwardOnly
            If Not rsALLPROFILE.EOF And Not rsALLPROFILE.BOF Then
                If rsALLPROFILE!Version < App.Revision Then
                    gconACCESS.Execute ("UPDATE ALL_PROFILE SET VERSION = '" & App.Revision & "' WHERE MODULENAME='" & App.EXEName & "'")
                ElseIf rsALLPROFILE!Version > App.Revision Then
                    MsgBox "You are using old " & App.EXEName & " version." & Chr(13) & "Please ask the administrator for the latest update!"
                    End
                End If
            End If
            Set rsALLPROFILE = Nothing
        End If
    End If
End Sub

Sub GET_COMPANYCODE()
    Dim rsALLPROFILE                                   As ADODB.Recordset
    Set rsALLPROFILE = New ADODB.Recordset
    rsALLPROFILE.Open "SELECT DISTINCT ISNULL(COMPANYCODE,'') AS COMPANYCODE FROM ALL_PROFILE", gconACCESS, adOpenForwardOnly
    If Not rsALLPROFILE.EOF And Not rsALLPROFILE.BOF Then
        COMPANY_CODE = rsALLPROFILE!COMPANYCODE
    End If
    Set rsALLPROFILE = Nothing
End Sub

Sub COMPANYCODE_VERSION()
    Dim CTR                                            As Integer
    Dim xCTR                                           As Integer
    CTR = 6
    ReDim COMPANY(CTR) As String
    COMPANY(0) = "HSR"
    COMPANY(1) = "HSP"
    COMPANY(2) = "HLP"
    COMPANY(3) = "HGC"
    COMPANY(4) = "HGH"
    COMPANY(5) = "HAM"
    
    For xCTR = 0 To CTR
        If COMPANY(xCTR) = COMPANY_CODE Then
            COMPANY_VERSION = COMPANY(xCTR)
        End If
    Next
End Sub

Function CHANGE_USER() As Boolean
    On Error GoTo FALSEUSER
    Dim SQL                                            As String
    Dim rsUSERNAME                                     As ADODB.Recordset
    SQL = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_RAMS_USERS' AND COLUMN_NAME='USER_NAME') " & vbCrLf
    SQL = SQL & "SELECT TOP 1 USER_NAME FROM ALL_RAMS_USERS"
    Set rsUSERNAME = gconACCESS.Execute(SQL)
    If Not rsUSERNAME.EOF And Not rsUSERNAME.BOF Then
        CHANGE_USER = True
    Else
FALSEUSER:
        CHANGE_USER = False
    End If
End Function

Function VALID_COMPANY_CODE_FORHAI() As Boolean
    Dim i                                              As Long
    Dim COUNTER                                        As Long

    COUNTER = 2
    ReDim STR(COUNTER) As String

    STR(0) = "HLP": STR(1) = "HAM": STR(2) = "HSP":

    For i = 0 To COUNTER
        If STR(i) = COMPANY_CODE Then
            VALID_COMPANY_CODE_FORHAI = True
        End If
    Next
End Function

Function TERMS(XXX As String) As String
    Dim rsxterms                                       As ADODB.Recordset
    Set rsxterms = New ADODB.Recordset
    rsxterms.Open "Select isnull(creditdays,0) as creditdays from all_customer where cuscde = '" & XXX & "'", gconDMIS
    If Not (rsxterms.EOF And rsxterms.BOF) Then
        TERMS = N2Str2Zero(rsxterms!CREDITDAYS)
    Else
        TERMS = 0
    End If
End Function

Sub ADD_AMIS_FIELD()
'Updated by: ACL
'Additional fields
    Dim SQL                                            As String
    SQL = ""
    SQL = "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='AMIS_JOURNAL_HD' AND COLUMN_NAME='DATEPOSTED')" & vbCrLf
    SQL = SQL & "ALTER TABLE AMIS_JOURNAL_HD" & vbCrLf
    SQL = SQL & "ADD DATEPOSTED      SMALLDATETIME" & vbCrLf
    gconDMIS.Execute (SQL)

    SQL = ""
    SQL = "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='AMIS_JOURNAL_HD' AND COLUMN_NAME='DATECANCELLED')" & vbCrLf
    SQL = SQL & "ALTER TABLE AMIS_JOURNAL_HD" & vbCrLf
    SQL = SQL & "ADD DATECANCELLED   SMALLDATETIME"
    gconDMIS.Execute (SQL)

    SQL = ""
    SQL = "IF NOT EXISTS(SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_VENDOR_TABLE' AND COLUMN_NAME='ATC')" & vbCrLf
    SQL = SQL & "ALTER TABLE ALL_VENDOR_TABLE" & vbCrLf
    SQL = SQL & "ADD ATC     NVARCHAR(10)"
    gconDMIS.Execute (SQL)

    SQL = ""
    SQL = "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_VENDOR_TABLE' AND COLUMN_NAME='ENTRY_DATE')" & vbCrLf
    SQL = SQL & "ALTER TABLE ALL_VENDOR_TABLE" & vbCrLf
    SQL = SQL & "ADD ENTRY_DATE  SMALLDATETIME"
    gconDMIS.Execute (SQL)
    
    SQL = ""
    SQL = "IF NOT EXISTS(SELECT * FROM SYS.DEFAULT_CONSTRAINTS WHERE NAME='DF_VENTRY_DATE')" & vbCrLf
    SQL = SQL & "ALTER TABLE ALL_VENDOR_TABLE" & vbCrLf
    SQL = SQL & "ADD CONSTRAINT DF_VENTRY_DATE DEFAULT (GETDATE()) FOR ENTRY_DATE"
    
    SQL = ""
    SQL = "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_CUSTOMER_TABLE' AND COLUMN_NAME='ENTRY_DATE')" & vbCrLf
    SQL = SQL & "ALTER TABLE ALL_CUSTOMER_TABLE" & vbCrLf
    SQL = SQL & "ADD ENTRY_DATE  SMALLDATETIME"
    gconDMIS.Execute (SQL)
    
    SQL = ""
    SQL = "IF NOT EXISTS(SELECT * FROM SYS.DEFAULT_CONSTRAINTS WHERE NAME='DF_CENTRY_DATE')" & vbCrLf
    SQL = SQL & "ALTER TABLE ALL_CUSTOMER_TABLE" & vbCrLf
    SQL = SQL & "ADD CONSTRAINT DF_CENTRY_DATE DEFAULT (GETDATE()) FOR ENTRY_DATE"
    
    SQL = ""
    SQL = "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_CUSTOMER_TABLE' AND COLUMN_NAME='TAX_AGENT')" & vbCrLf
    SQL = SQL & "ALTER TABLE ALL_CUSTOMER_TABLE" & vbCrLf
    SQL = SQL & "ADD TAX_AGENT   BIT"
    gconDMIS.Execute SQL
    
    SQL = ""
    SQL = "ALTER VIEW ALL_CUSTOMER" & vbCrLf
    SQL = SQL & "AS" & vbCrLf
    SQL = SQL & "SELECT     CUSCDE, APOD, ACCOUNTNO, CUSCOMP, COMPANYADD, ACCTNAME, LASTNAME, FIRSTNAME, MIDDLEINITIAL, SEX, CUSTOMERADD," & vbCrLf
    SQL = SQL & "           PROVINCIALADD, ZIPCODE, HOMEPHONE, TELEPHONENO, CUSCAT, PLATENO, OLDCODE, CUSTYPE, LEADSOURCE, TITLE, DEPARTMENT, EMAIL," & vbCrLf
    SQL = SQL & "           MOBILE, FAX, ASSISTANT, ASSTPHONE, CITY, BIRTHDATE, SPOUSE, DESCRIPTION, CUSTOMERSOURCELEAD, DELIVERYADDRESS, CREDITTERM," & vbCrLf
    SQL = SQL & "           CREDITDAYS , CREDITLIMIT, USERCODE, LASTUPDATE, TIMEUPDATE, USERCODE2, EDITDATE, EDITTIME, TIN, ID, TAX_AGENT, ENTRY_DATE" & vbCrLf
    SQL = SQL & "From dbo.ALL_Customer_Table" & vbCrLf
    SQL = SQL & "WHERE     (CUSCDE <> '999999')"
    gconDMIS.Execute SQL
End Sub

Public Function Null2Bit(XXX As Variant) As Integer
    If Null2Bool(XXX) = True Then
        Null2Bit = 1
    Else
        Null2Bit = 0
    End If
End Function

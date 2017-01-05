Attribute VB_Name = "modDMISGlobal"
Option Explicit
'UPDATED BY: JUN/ARNOLD-------------------------------------------------
'DATE UPDATED: 06-11-2009
'DESCRIPTION: DECLARE GLOBAL DUE TO PRINTING PURPOSE FOR CUSTOMER LEDGER
Public xBALANCE        As Double
Public BEG_BALANCE_DATE As Date
'UPDATED BY: JUN/ARNOLD-------------------------------------------------

Public QC_MODULE_ON                                    As String
Public LOGCODE As String, LOGNAME As String, LOGLEVEL As String, LOGTIME As String, LOGDATE As String
Attribute LOGNAME.VB_VarUserMemId = 1073741825
Attribute LOGLEVEL.VB_VarUserMemId = 1073741825
Attribute LOGTIME.VB_VarUserMemId = 1073741825
Attribute LOGDATE.VB_VarUserMemId = 1073741825
Public FROM_APPOINTMENT                                As String
Public RECEIVED_FROM_PO                                As String
Public BIR_RELIEF_Connection                           As String
Attribute BIR_RELIEF_Connection.VB_VarUserMemId = 1073741829

Public wizVar, CryptVar                                As Object
Attribute wizVar.VB_VarUserMemId = 1073741830
Attribute CryptVar.VB_VarUserMemId = 1073741830
Public AccessCNT                                       As Integer
Attribute AccessCNT.VB_VarUserMemId = 1073741832
Public AcctCodeArray(5000), AcctNameArray(5000)        As String
Attribute AcctCodeArray.VB_VarUserMemId = 1073741833
Public vREFERENCENO                                    As String
Public SelectEntity                                    As String
Public xPAIDFOR                                        As String
Public rEndingBalance                                  As Double

Public BILANG                                          As Long
Attribute BILANG.VB_VarUserMemId = 1073741835
Public SEARCH_TAB                                      As Long
Public JOURNALTYPE                                     As String
Attribute JOURNALTYPE.VB_VarUserMemId = 1073741837
Public REPORT_RANGETYPE                                As String
Attribute REPORT_RANGETYPE.VB_VarUserMemId = 1073741838
Public REPORT_EXPENSETYPE                              As String
Attribute REPORT_EXPENSETYPE.VB_VarUserMemId = 1073741839
Public CUST_LEDGER_TYPE                                As String
Attribute CUST_LEDGER_TYPE.VB_VarUserMemId = 1073741840
'Public REPORT_AR As Strings
Public Report_AR                                       As String
Attribute Report_AR.VB_VarUserMemId = 1073741841
Public REPORT_AP                                       As String
Public REFRESH_ACCOUNT                                 As Boolean
Attribute REFRESH_ACCOUNT.VB_VarUserMemId = 1073741842
Public CURRENT_CUSCODE                                 As String
Attribute CURRENT_CUSCODE.VB_VarUserMemId = 1073741843
Public CURRENT_VENDORCODE                              As String
Attribute CURRENT_VENDORCODE.VB_VarUserMemId = 1073741844

Public INVOICE_Type                                    As String
Attribute INVOICE_Type.VB_VarUserMemId = 1073741845
Public MYOB_JTYPE                                      As String
Attribute MYOB_JTYPE.VB_VarUserMemId = 1073741846
Public Const VAT_RATE = 12
Public EXTRACT_TYPE                                    As String
Attribute EXTRACT_TYPE.VB_VarUserMemId = 1073741847

Public CASH_SALES                                      As String
Attribute CASH_SALES.VB_VarUserMemId = 1073741848
Public CHARGE_SALES                                    As String
Attribute CHARGE_SALES.VB_VarUserMemId = 1073741849
Public CASH_DISCOUNT                                   As String
Attribute CASH_DISCOUNT.VB_VarUserMemId = 1073741850
Public CHARGE_DISCOUNT                                 As String
Attribute CHARGE_DISCOUNT.VB_VarUserMemId = 1073741851
Public CASH_COSTOFSALES                                As String
Attribute CASH_COSTOFSALES.VB_VarUserMemId = 1073741852
Public CHARGE_COSTOFSALES                              As String
Attribute CHARGE_COSTOFSALES.VB_VarUserMemId = 1073741853
Public OPERATIONAL_EXPENSE                             As String
Attribute OPERATIONAL_EXPENSE.VB_VarUserMemId = 1073741854
Public ADMIN_EXPENSE                                   As String
Attribute ADMIN_EXPENSE.VB_VarUserMemId = 1073741855
Public OTHER_INCOME                                    As String
Attribute OTHER_INCOME.VB_VarUserMemId = 1073741856
Public OTHER_EXPENSE                                   As String
Attribute OTHER_EXPENSE.VB_VarUserMemId = 1073741857
Public CURRENT_ASSET                                   As String
Attribute CURRENT_ASSET.VB_VarUserMemId = 1073741858
Public TAX_CREDITS                                     As String
Attribute TAX_CREDITS.VB_VarUserMemId = 1073741859
Public PROPERTY_EQUIPMENT                              As String
Attribute PROPERTY_EQUIPMENT.VB_VarUserMemId = 1073741860
Public ACCUMULATED_DEPRECIATION                        As String
Attribute ACCUMULATED_DEPRECIATION.VB_VarUserMemId = 1073741861
Public OTHER_ASSET                                     As String
Attribute OTHER_ASSET.VB_VarUserMemId = 1073741862

Public CUSCODE                                         As String
Attribute CUSCODE.VB_VarUserMemId = 1073741863
Public MODULENAME
Attribute MODULENAME.VB_VarUserMemId = 1073741864

Public INVOICE_DETAIL_TYPE, INVOICE_DETAIL_TRANNO      As String
Attribute INVOICE_DETAIL_TYPE.VB_VarUserMemId = 1073741865
Attribute INVOICE_DETAIL_TRANNO.VB_VarUserMemId = 1073741865

Public OVERWRAYT                                       As Boolean
Attribute OVERWRAYT.VB_VarUserMemId = 1073741867
Public GVD_DATABASE_PATH                               As String
Attribute GVD_DATABASE_PATH.VB_VarUserMemId = 1073741868
Public SKIN_PATH                                       As String
Attribute SKIN_PATH.VB_VarUserMemId = 1073741869
Public NEYM                                            As String
Attribute NEYM.VB_VarUserMemId = 1073741870
Public ADRES                                           As String
Attribute ADRES.VB_VarUserMemId = 1073741871
Public TELLNO                                          As String
Attribute TELLNO.VB_VarUserMemId = 1073741872
Public PURLASTNEYM                                     As String
Attribute PURLASTNEYM.VB_VarUserMemId = 1073741873
Public PURFIRSTNEYM                                    As String
Attribute PURFIRSTNEYM.VB_VarUserMemId = 1073741874
Public PURMIDDLE                                       As String
Attribute PURMIDDLE.VB_VarUserMemId = 1073741875
Public PRODUCTNO                                       As String
Attribute PRODUCTNO.VB_VarUserMemId = 1073741876
Public LASTNEYM                                        As String
Attribute LASTNEYM.VB_VarUserMemId = 1073741877
Public FIRSTNEYM                                       As String
Attribute FIRSTNEYM.VB_VarUserMemId = 1073741878
Public MIDDLE                                          As String
Attribute MIDDLE.VB_VarUserMemId = 1073741879
Public Add_o_Edit                                      As String
Attribute Add_o_Edit.VB_VarUserMemId = 1073741880
Public EMPINFOSHOW                                     As Boolean
Attribute EMPINFOSHOW.VB_VarUserMemId = 1073741881

Public SAECODE                                         As String
Attribute SAECODE.VB_VarUserMemId = 1073741882
Public LOGSAE                                          As String
Attribute LOGSAE.VB_VarUserMemId = 1073741883
Public SAENAME                                         As String
Attribute SAENAME.VB_VarUserMemId = 1073741884

Public INVOICENO                                       As String
Public InvoiceType                                     As String

Public rKeyPublicension(1000)                          As Integer
Attribute rKeyPublicension.VB_VarUserMemId = 1073741885
Public EncryptoFile(100000)                            As String
Attribute EncryptoFile.VB_VarUserMemId = 1073741886
Public CryptoKey                                       As Variant
Attribute CryptoKey.VB_VarUserMemId = 1073741887
Public Maxwiz                                          As Long
Attribute Maxwiz.VB_VarUserMemId = 1073741888
Public SEARCH_BY                                       As String
Attribute SEARCH_BY.VB_VarUserMemId = 1073741889
Public VInoArray(3000)                                 As String
Attribute VInoArray.VB_VarUserMemId = 1073741890
Public VICusNamArray(3000)                             As String
Attribute VICusNamArray.VB_VarUserMemId = 1073741891
Public CusVInoArray(3000)                              As String
Attribute CusVInoArray.VB_VarUserMemId = 1073741892
Public CusNamArray(3000)                               As String
Attribute CusNamArray.VB_VarUserMemId = 1073741893
Public CusNameArray(3000)                              As String
Attribute CusNameArray.VB_VarUserMemId = 1073741894
Public CusProdNoArray(3000)                            As String
Attribute CusProdNoArray.VB_VarUserMemId = 1073741895
Public CusProdNoArray2(3000)                           As String
Attribute CusProdNoArray2.VB_VarUserMemId = 1073741896
Public CusCodeArray(3000)                              As String
Attribute CusCodeArray.VB_VarUserMemId = 1073741897
Public BILANG2                                         As Long
Attribute BILANG2.VB_VarUserMemId = 1073741898
Public BILANG_CusName                                  As Long
Attribute BILANG_CusName.VB_VarUserMemId = 1073741899
Public BILANG_CusName2                                 As Long
Attribute BILANG_CusName2.VB_VarUserMemId = 1073741900
Public CUST_REPT_TYPE                                  As String
Attribute CUST_REPT_TYPE.VB_VarUserMemId = 1073741901
Public PARTS_ISSUED_TO_CUSTOMER_TYPE                   As String

Public Const WorkTimeStart                             As String = "8:00 AM"
Public Const WorkTimeEnd                               As String = "5:00 PM"
Public FILE_GRAPH                                      As String
Attribute FILE_GRAPH.VB_VarUserMemId = 1073741902

'SIGNATORIES AND ADDRESSES
Public PREPARED_BY, CHECKED_BY, GENERAL_MANAGER, APPROVED_BY, ACCOUNT_NO, BANK_NAME, BANK_LOCATION, BANK_MANAGER, SECRETARY, NOTED_BY As String
Attribute PREPARED_BY.VB_VarUserMemId = 1073741903
Attribute CHECKED_BY.VB_VarUserMemId = 1073741903
Attribute GENERAL_MANAGER.VB_VarUserMemId = 1073741903
Attribute APPROVED_BY.VB_VarUserMemId = 1073741903
Attribute ACCOUNT_NO.VB_VarUserMemId = 1073741903
Attribute BANK_NAME.VB_VarUserMemId = 1073741903
Attribute BANK_LOCATION.VB_VarUserMemId = 1073741903
Attribute BANK_MANAGER.VB_VarUserMemId = 1073741903
Attribute SECRETARY.VB_VarUserMemId = 1073741903
Attribute NOTED_BY.VB_VarUserMemId = 1073741903
Public PreparedBy, ApprovedBy, CheckedBy, SalesDispatcher, GeneralManager, DeliveredBy, FinancingManager As String
Attribute PreparedBy.VB_VarUserMemId = 1073741913
Attribute ApprovedBy.VB_VarUserMemId = 1073741913
Attribute CheckedBy.VB_VarUserMemId = 1073741913
Attribute SalesDispatcher.VB_VarUserMemId = 1073741913
Attribute GeneralManager.VB_VarUserMemId = 1073741913
Attribute DeliveredBy.VB_VarUserMemId = 1073741913
Attribute FinancingManager.VB_VarUserMemId = 1073741913
'SIGNATORIES AND ADDRESSES
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'UPDATED BY: JUN
'DATE UPDATED: 09142008
'DESCRIPTION: SIGNATORIES DESIGNATION
'SIGNATORIES DESIGNATION
Public PreparedByDesig, CheckedByDesig, SalesApprovedDesig, SalesDispatcherDesig, GeneralManagerDesig, DeliveredByDesig, FinancingManagerDesig As String
Attribute CheckedByDesig.VB_VarUserMemId = 1073741831
Attribute SalesApprovedDesig.VB_VarUserMemId = 1073741831
Attribute SalesDispatcherDesig.VB_VarUserMemId = 1073741831
Attribute GeneralManagerDesig.VB_VarUserMemId = 1073741831
Attribute DeliveredByDesig.VB_VarUserMemId = 1073741831
Attribute FinancingManagerDesig.VB_VarUserMemId = 1073741831
'-------------------------------------------------------------------------------------------------------------------------------------------------------


Public gconINVENTORY                                   As ADODB.Connection
Attribute gconINVENTORY.VB_VarUserMemId = 1073741920

Public PROC_TYPE                                       As String
Attribute PROC_TYPE.VB_VarUserMemId = 1073741921
Public WAREHOUSETYPE As String, COUNTERTYPE As String, MAT_COUNTERTYPE As String, ORDERTYPE As String, VPAMCOR As String, BIR_YearEnd As String, ISSREPTYPE As String, ORDER_REPORT As String
Attribute WAREHOUSETYPE.VB_VarUserMemId = 1073741922
Attribute COUNTERTYPE.VB_VarUserMemId = 1073741922
Attribute MAT_COUNTERTYPE.VB_VarUserMemId = 1073741922
Attribute ORDERTYPE.VB_VarUserMemId = 1073741922
Attribute VPAMCOR.VB_VarUserMemId = 1073741922
Attribute BIR_YearEnd.VB_VarUserMemId = 1073741922
Attribute ISSREPTYPE.VB_VarUserMemId = 1073741922
Attribute ORDER_REPORT.VB_VarUserMemId = 1073741922
Public BIRDATA_Connection, BIR_DATABASE_PATH           As String
Attribute BIRDATA_Connection.VB_VarUserMemId = 1073741930
Attribute BIR_DATABASE_PATH.VB_VarUserMemId = 1073741930
Public CSMS_PARTSQUERY                                 As Boolean
Attribute CSMS_PARTSQUERY.VB_VarUserMemId = 1073741932

Public TOTJOBAMT, TOTJOBDISC, TOTJOBTAX                As Double
Attribute TOTJOBAMT.VB_VarUserMemId = 1073741933
Attribute TOTJOBDISC.VB_VarUserMemId = 1073741933
Attribute TOTJOBTAX.VB_VarUserMemId = 1073741933
Public TOTPARTSAMT, TOTPARTSDISC, TOTPARTSTAX          As Double
Attribute TOTPARTSAMT.VB_VarUserMemId = 1073741936
Attribute TOTPARTSDISC.VB_VarUserMemId = 1073741936
Attribute TOTPARTSTAX.VB_VarUserMemId = 1073741936
Public TOTMATAMT, TOTMATDISC, TOTMATTAX                As Double
Attribute TOTMATAMT.VB_VarUserMemId = 1073741939
Attribute TOTMATDISC.VB_VarUserMemId = 1073741939
Attribute TOTMATTAX.VB_VarUserMemId = 1073741939
Public TOTACCAMT, TOTACCDISC, TOTACCTAX                As Double
Attribute TOTACCAMT.VB_VarUserMemId = 1073741835
Attribute TOTACCDISC.VB_VarUserMemId = 1073741835
Attribute TOTACCTAX.VB_VarUserMemId = 1073741835
Public DNPIDFrom, DNPIDTo                              As Long
Attribute DNPIDFrom.VB_VarUserMemId = 1073741942
Attribute DNPIDTo.VB_VarUserMemId = 1073741942


Public gconBIR_RELIEF                                  As ADODB.Connection
Attribute gconBIR_RELIEF.VB_VarUserMemId = 1073741944
Public PARTSQUERY                                      As Integer
Attribute PARTSQUERY.VB_VarUserMemId = 1073741945
Public rKeyDimension(1000)                             As Integer
Attribute rKeyDimension.VB_VarUserMemId = 1073741946

Public VoiceMsg                                        As Boolean
Attribute VoiceMsg.VB_VarUserMemId = 1073741947
Public STOCK_TYPE                                      As String
Attribute STOCK_TYPE.VB_VarUserMemId = 1073741948
Public PRR_REPORT                                      As String
Attribute PRR_REPORT.VB_VarUserMemId = 1073741949
Public FORECASTING_BUTTON_CLICK                        As Integer
Attribute FORECASTING_BUTTON_CLICK.VB_VarUserMemId = 1073741950
Public PRR_BUTTON_CLICK                                As Integer
Attribute PRR_BUTTON_CLICK.VB_VarUserMemId = 1073741838


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
Attribute Y_REGRESSION_INTERVAL.VB_VarUserMemId = 1073741951
Public Const X_MEAN_INTERVAL = 1

Public Const HARI_LEAD_TIME = 1.25
Public Const HARI_ORDER_FREQUENCY = 1.5


Public ROSHOW                                          As Boolean
Attribute ROSHOW.VB_VarUserMemId = 1073741952
Public ESTISHOW                                        As Boolean
Attribute ESTISHOW.VB_VarUserMemId = 1073741953
Public ESTIKCNT                                        As Integer
Attribute ESTIKCNT.VB_VarUserMemId = 1073741954

Public RO_OR_ESTI_OR_PART                              As String
Attribute RO_OR_ESTI_OR_PART.VB_VarUserMemId = 1073741955

Public TOTJOBDISCVAL                                   As Double
Attribute TOTJOBDISCVAL.VB_VarUserMemId = 1073741956
Public TOTPARTSDISCVAL                                 As Double
Attribute TOTPARTSDISCVAL.VB_VarUserMemId = 1073741957
Public TOTMATDISCVAL                                   As Double
Attribute TOTMATDISCVAL.VB_VarUserMemId = 1073741958
Public TOTACCDISCVAL                                   As Double
Attribute TOTACCDISCVAL.VB_VarUserMemId = 1073741839



Public SearchBy, SEARCHCUSTOMERNAME, SEARCHPLATENO     As String
Attribute SearchBy.VB_VarUserMemId = 1073741959
Attribute SEARCHCUSTOMERNAME.VB_VarUserMemId = 1073741959
Attribute SEARCHPLATENO.VB_VarUserMemId = 1073741959

Public QUESTION_TEST                                   As String
Attribute QUESTION_TEST.VB_VarUserMemId = 1073741962

Public EDIT_RO                                         As String
Attribute EDIT_RO.VB_VarUserMemId = 1073741963



Public SUBTOTAL_ADDON                                  As Double
Attribute SUBTOTAL_ADDON.VB_VarUserMemId = 1073741964
Public BOOKTYPE                                        As String
Attribute BOOKTYPE.VB_VarUserMemId = 1073741965
Public VAT_OR                                          As Integer
Attribute VAT_OR.VB_VarUserMemId = 1073741966
Public OR_VAT_NONVAT                                   As String
Attribute OR_VAT_NONVAT.VB_VarUserMemId = 1073741967
Public CANCEL_OR_VAT_NONVAT                            As String
Attribute CANCEL_OR_VAT_NONVAT.VB_VarUserMemId = 1073741968
Public BRANCH_CODE, OR_NUMBER_GLOBAL                   As String
Attribute BRANCH_CODE.VB_VarUserMemId = 1073741969
Attribute OR_NUMBER_GLOBAL.VB_VarUserMemId = 1073741969

Public RECEIPTS_AMOUNT, AMOUNT_TENDERED, CHANGE_DUE    As Double
Attribute RECEIPTS_AMOUNT.VB_VarUserMemId = 1073741971
Attribute AMOUNT_TENDERED.VB_VarUserMemId = 1073741971
Attribute CHANGE_DUE.VB_VarUserMemId = 1073741971
Public MODE_OF_PAYMENT                                 As String
Attribute MODE_OF_PAYMENT.VB_VarUserMemId = 1073741974

Public Const MAX_PETTYFUND = 0
Public Const MAX_LTOFUND = 0
Public Const CHANGE_FUND = 3000
Public CURRENT_CUTOFF_DATE                             As String
Attribute CURRENT_CUTOFF_DATE.VB_VarUserMemId = 1073741975
Public CASHPOSITION_CUTOFF_DATE                        As String
Attribute CASHPOSITION_CUTOFF_DATE.VB_VarUserMemId = 1073741976
Public IsLTOIsPettyCash                                As String
Attribute IsLTOIsPettyCash.VB_VarUserMemId = 1073741977
Public CMIS_Report_Range                               As String
Attribute CMIS_Report_Range.VB_VarUserMemId = 1073741978
Public CMIS_Type_Of_Report                             As String
Attribute CMIS_Type_Of_Report.VB_VarUserMemId = 1073741979
Public CURRENT_CUST_CODE                               As String
Attribute CURRENT_CUST_CODE.VB_VarUserMemId = 1073741980
Public CASH_OPTIONS                                    As String
Attribute CASH_OPTIONS.VB_VarUserMemId = 1073741981
Public TYPE_ON_HAND                                    As String
Attribute TYPE_ON_HAND.VB_VarUserMemId = 1073741982

Public PERIODMONTH, PERIODYEAR                         As Integer
Attribute PERIODMONTH.VB_VarUserMemId = 1073741983
Attribute PERIODYEAR.VB_VarUserMemId = 1073741983

Public EMPLOYEE_NO                                     As String
Attribute EMPLOYEE_NO.VB_VarUserMemId = 1073741985
Public PREPARED_BY_DESIGNATION, APPROVED_BY_DESIGNATION As String
Attribute PREPARED_BY_DESIGNATION.VB_VarUserMemId = 1073741986
Attribute APPROVED_BY_DESIGNATION.VB_VarUserMemId = 1073741986

Public IssuedBy, NotedBy
Attribute IssuedBy.VB_VarUserMemId = 1073741988
Attribute NotedBy.VB_VarUserMemId = 1073741988

Public LEDGERSHOW, HEADEMPINFOSHOW, CASHOW             As Boolean
Attribute LEDGERSHOW.VB_VarUserMemId = 1073741990
Attribute HEADEMPINFOSHOW.VB_VarUserMemId = 1073741990
Attribute CASHOW.VB_VarUserMemId = 1073741990
Public GENFROM, GENTO, HEADOREMP, EMP_TYPE             As String
Attribute GENFROM.VB_VarUserMemId = 1073741993
Attribute GENTO.VB_VarUserMemId = 1073741993
Attribute HEADOREMP.VB_VarUserMemId = 1073741993
Attribute EMP_TYPE.VB_VarUserMemId = 1073741993

Public EmpInfoEmpno                                    As Object
Attribute EmpInfoEmpno.VB_VarUserMemId = 1073741997
Public FormYearlyRequest                               As String
Attribute FormYearlyRequest.VB_VarUserMemId = 1073741998

Public SQL_STATEMENT                                   As String
Attribute SQL_STATEMENT.VB_VarUserMemId = 1073741840
Public PROCESS_OPTION                                  As String
Attribute PROCESS_OPTION.VB_VarUserMemId = 1073741999
Public IMPNO                                           As String
Attribute IMPNO.VB_VarUserMemId = 1073742000

Public OVERTIME_CODES                                  As String
Attribute OVERTIME_CODES.VB_VarUserMemId = 1073742001
Public OVERTIME_RATE                                   As Double
Attribute OVERTIME_RATE.VB_VarUserMemId = 1073742002

Public PAYROLLCODE_FROM1                               As Integer
Attribute PAYROLLCODE_FROM1.VB_VarUserMemId = 1073742003
Public PAYROLLCODE_FROM2                               As Integer
Attribute PAYROLLCODE_FROM2.VB_VarUserMemId = 1073742004
Public PAYROLLCODE_TO1                                 As Integer
Attribute PAYROLLCODE_TO1.VB_VarUserMemId = 1073742005
Public PAYROLLCODE_TO2                                 As Integer
Attribute PAYROLLCODE_TO2.VB_VarUserMemId = 1073742006
Public PAYROLL_BASE                                    As String
Attribute PAYROLL_BASE.VB_VarUserMemId = 1073742007
Public PAYROLL_NO_OF_DAYS                              As Integer
Attribute PAYROLL_NO_OF_DAYS.VB_VarUserMemId = 1073742008
Public PAYROLL_CODE                                    As Integer
Attribute PAYROLL_CODE.VB_VarUserMemId = 1073742009

Public CUTTOFF_CODE                                    As String
Attribute CUTTOFF_CODE.VB_VarUserMemId = 1073741841
Public PAY_MONTH                                       As Integer
Attribute PAY_MONTH.VB_VarUserMemId = 1073741842
Public PAY_YEAR                                        As Integer
Attribute PAY_YEAR.VB_VarUserMemId = 1073741843
Public DEDUCTION_OPTION                                As String
Attribute DEDUCTION_OPTION.VB_VarUserMemId = 1073741844


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
Function getJTYPE(XINVOICE As String, xinvoiceType As String) As String
    'Update by BTT:11/21/2008
    Dim rs                                             As New ADODB.Recordset
    Set rs = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_HD WHERE INVOICENO ='" & XINVOICE & "' and INVOICETYPE ='" & xinvoiceType & "'")
    If Not (rs.EOF And rs.BOF) Then
        getJTYPE = Null2String(rs!jtype)
    Else
        getJTYPE = "NULL"
    End If
    Set rs = Nothing
End Function

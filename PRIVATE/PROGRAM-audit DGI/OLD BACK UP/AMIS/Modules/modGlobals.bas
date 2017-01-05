Attribute VB_Name = "modAMISGlobals"
Option Explicit
Public LOGCODE As String, LOGNAME As String, LOGLEVEL As String, LOGTIME As String, LOGDATE As String
Attribute LOGNAME.VB_VarUserMemId = 1073741824
Attribute LOGLEVEL.VB_VarUserMemId = 1073741824
Attribute LOGTIME.VB_VarUserMemId = 1073741824
Attribute LOGDATE.VB_VarUserMemId = 1073741824

Public BIR_RELIEF_Connection                           As String
Attribute BIR_RELIEF_Connection.VB_VarUserMemId = 1073741830

Public wizVar, CryptVar                                As Object
Attribute wizVar.VB_VarUserMemId = 1073741831
Attribute CryptVar.VB_VarUserMemId = 1073741831
Public AccessCNT                                       As Integer
Attribute AccessCNT.VB_VarUserMemId = 1073741833
Public AcctCodeArray(5000), AcctNameArray(5000)        As String
Attribute AcctCodeArray.VB_VarUserMemId = 1073741834

Public BILANG, SEARCH_TAB                              As Long
Attribute BILANG.VB_VarUserMemId = 1073741836
Attribute SEARCH_TAB.VB_VarUserMemId = 1073741836
Public JOURNALTYPE                                     As String
Attribute JOURNALTYPE.VB_VarUserMemId = 1073741838
Public REPORT_RANGETYPE                                As String
Attribute REPORT_RANGETYPE.VB_VarUserMemId = 1073741839
Public REPORT_EXPENSETYPE                              As String
Attribute REPORT_EXPENSETYPE.VB_VarUserMemId = 1073741840
Public CUST_LEDGER_TYPE                                As String
Attribute CUST_LEDGER_TYPE.VB_VarUserMemId = 1073741841
'Public REPORT_AR As Strings
Public Report_Ar                                       As String
Attribute Report_Ar.VB_VarUserMemId = 1073741842
Public REFRESH_ACCOUNT                                 As Boolean
Attribute REFRESH_ACCOUNT.VB_VarUserMemId = 1073741843
Public CURRENT_CUSCODE                                 As String
Attribute CURRENT_CUSCODE.VB_VarUserMemId = 1073741844
Public CURRENT_VENDORCODE                              As String
Attribute CURRENT_VENDORCODE.VB_VarUserMemId = 1073741845

Public INVOICE_Type                                    As String
Attribute INVOICE_Type.VB_VarUserMemId = 1073741846
Public MYOB_JTYPE                                      As String
Attribute MYOB_JTYPE.VB_VarUserMemId = 1073741847
Public Const VAT_RATE = 12
Public EXTRACT_TYPE                                    As String
Attribute EXTRACT_TYPE.VB_VarUserMemId = 1073741848

Public CASH_SALES                                      As String
Attribute CASH_SALES.VB_VarUserMemId = 1073741849
Public CHARGE_SALES                                    As String
Attribute CHARGE_SALES.VB_VarUserMemId = 1073741850
Public CASH_DISCOUNT                                   As String
Attribute CASH_DISCOUNT.VB_VarUserMemId = 1073741851
Public CHARGE_DISCOUNT                                 As String
Attribute CHARGE_DISCOUNT.VB_VarUserMemId = 1073741852
Public CASH_COSTOFSALES                                As String
Attribute CASH_COSTOFSALES.VB_VarUserMemId = 1073741853
Public CHARGE_COSTOFSALES                              As String
Attribute CHARGE_COSTOFSALES.VB_VarUserMemId = 1073741854
Public OPERATIONAL_EXPENSE                             As String
Attribute OPERATIONAL_EXPENSE.VB_VarUserMemId = 1073741855
Public ADMIN_EXPENSE                                   As String
Attribute ADMIN_EXPENSE.VB_VarUserMemId = 1073741856
Public OTHER_INCOME                                    As String
Attribute OTHER_INCOME.VB_VarUserMemId = 1073741857
Public OTHER_EXPENSE                                   As String
Attribute OTHER_EXPENSE.VB_VarUserMemId = 1073741858
Public CURRENT_ASSET                                   As String
Attribute CURRENT_ASSET.VB_VarUserMemId = 1073741859
Public TAX_CREDITS                                     As String
Attribute TAX_CREDITS.VB_VarUserMemId = 1073741860
Public PROPERTY_EQUIPMENT                              As String
Attribute PROPERTY_EQUIPMENT.VB_VarUserMemId = 1073741861
Public ACCUMULATED_DEPRECIATION                        As String
Attribute ACCUMULATED_DEPRECIATION.VB_VarUserMemId = 1073741862
Public OTHER_ASSET                                     As String
Attribute OTHER_ASSET.VB_VarUserMemId = 1073741863

Public CUSCODE                                         As String
Attribute CUSCODE.VB_VarUserMemId = 1073741864
Public Const MODULENAME = "AMIS"

Public INVOICE_DETAIL_TYPE, INVOICE_DETAIL_TRANNO As String


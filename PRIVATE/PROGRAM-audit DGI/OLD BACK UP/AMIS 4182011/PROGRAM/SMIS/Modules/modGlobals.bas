Attribute VB_Name = "modSMISGlobals"
Option Explicit
Public OVERWRAYT                                                      As Boolean
Public GVD_DATABASE_PATH                                              As String
Public SKIN_PATH                                                      As String
Public NEYM                                                           As String
Attribute NEYM.VB_VarUserMemId = 1073741832
Public ADRES                                                          As String
Public TELLNO                                                         As String
Public PURLASTNEYM                                                    As String
Public PURFIRSTNEYM                                                   As String
Attribute PURFIRSTNEYM.VB_VarUserMemId = 1073741835
Public PURMIDDLE                                                      As String
Public CusCode                                                        As String
Attribute CusCode.VB_VarUserMemId = 1073741838
Public PRODUCTNO                                                      As String
Public LASTNEYM                                                       As String
Attribute LASTNEYM.VB_VarUserMemId = 1073741840
Public FIRSTNEYM                                                      As String
Public MIDDLE                                                         As String
Public Add_o_Edit                                                     As String
Public EMPINFOSHOW                                                    As Boolean
Attribute EMPINFOSHOW.VB_VarUserMemId = 1073741844

Public LOGCODE                                                        As String
Attribute LOGCODE.VB_VarUserMemId = 1073741845
Public LOGLEVEL                                                       As String
Public LOGNAME                                                        As String
Public LOGTIME                                                        As String
Public LOGDATE                                                        As String
Public SAECODE                                                        As String
Public LOGSAE                                                         As String
Public SAENAME                                                        As String

Public rKeyPublicension(1000)                                         As Integer
Attribute rKeyPublicension.VB_VarUserMemId = 1073741850
Public EncryptoFile(100000)                                           As String
Attribute EncryptoFile.VB_VarUserMemId = 1073741851
Public CryptoKey                                                      As Variant
Attribute CryptoKey.VB_VarUserMemId = 1073741852
Public Maxwiz, AccessCNT                                              As Long
Attribute Maxwiz.VB_VarUserMemId = 1073741853
Attribute AccessCNT.VB_VarUserMemId = 1073741853
Public wizVar, CryptVar                                               As Object
Attribute wizVar.VB_VarUserMemId = 1073741855
Attribute CryptVar.VB_VarUserMemId = 1073741855
Public SEARCH_BY                                                      As String
Attribute SEARCH_BY.VB_VarUserMemId = 1073741857
Public VInoArray(3000)                                                As String
Attribute VInoArray.VB_VarUserMemId = 1073741858
Public VICusNamArray(3000)                                            As String
Public CusVInoArray(3000)                                             As String
Attribute CusVInoArray.VB_VarUserMemId = 1073741860
Public CusNamArray(3000)                                              As String
Public CusNameArray(3000)                                             As String
Attribute CusNameArray.VB_VarUserMemId = 1073741862
Public CusProdNoArray(3000)                                           As String
Public CusProdNoArray2(3000)                                          As String
Public CusCodeArray(3000)                                             As String
Public BILANG                                                         As Long
Attribute BILANG.VB_VarUserMemId = 1073741866
Public BILANG2                                                        As Long
Public BILANG_CusName                                                 As Long
Attribute BILANG_CusName.VB_VarUserMemId = 1073741868
Public BILANG_CusName2                                                As Long
Public SEARCH_TAB                                                     As String
Attribute SEARCH_TAB.VB_VarUserMemId = 1073741870
Public CUST_REPT_TYPE                                                 As String
Attribute CUST_REPT_TYPE.VB_VarUserMemId = 1073741871


Public Const MODULENAME                                               As String = "SMIS"
Public Const WorkTimeStart                                            As String = "8:00 AM"
Public Const WorkTimeEnd                                              As String = "5:00 PM"
Public FILE_GRAPH                                                     As String

'SIGNATORIES AND ADDRESSES
Public PREPARED_BY, CHECKED_BY, GENERAL_MANAGER, APPROVED_BY, ACCOUNT_NO, BANK_NAME, BANK_LOCATION, BANK_MANAGER, SECRETARY, NOTED_BY As String
Attribute CHECKED_BY.VB_VarUserMemId = 1073741853
Attribute GENERAL_MANAGER.VB_VarUserMemId = 1073741853
Attribute APPROVED_BY.VB_VarUserMemId = 1073741853
Attribute ACCOUNT_NO.VB_VarUserMemId = 1073741853
Attribute BANK_NAME.VB_VarUserMemId = 1073741853
Attribute BANK_LOCATION.VB_VarUserMemId = 1073741853
Attribute BANK_MANAGER.VB_VarUserMemId = 1073741853
Attribute SECRETARY.VB_VarUserMemId = 1073741853
Attribute NOTED_BY.VB_VarUserMemId = 1073741853
Public PreparedBy, ApprovedBy, CheckedBy, SalesDispatcher, GeneralManager, DeliveredBy, FinancingManager As String
Attribute PreparedBy.VB_VarUserMemId = 1073741862
Attribute ApprovedBy.VB_VarUserMemId = 1073741862
Attribute CheckedBy.VB_VarUserMemId = 1073741862
Attribute SalesDispatcher.VB_VarUserMemId = 1073741862
Attribute GeneralManager.VB_VarUserMemId = 1073741862
Attribute DeliveredBy.VB_VarUserMemId = 1073741862
Attribute FinancingManager.VB_VarUserMemId = 1073741862
'SIGNATORIES AND ADDRESSES


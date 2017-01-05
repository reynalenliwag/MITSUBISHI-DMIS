Attribute VB_Name = "modPMIOSGlobals"
Option Explicit
Public OVERWRAYT                                       As Boolean
Public gconINVENTORY                                   As ADODB.Connection

Public CUSCODE, LASTNEYM, FIRSTNEYM, MIDDLE, PROC_TYPE As String
Attribute CUSCODE.VB_VarUserMemId = 1073741831
Attribute LASTNEYM.VB_VarUserMemId = 1073741831
Attribute FIRSTNEYM.VB_VarUserMemId = 1073741831
Attribute MIDDLE.VB_VarUserMemId = 1073741831
Attribute PROC_TYPE.VB_VarUserMemId = 1073741831
Public NEYM, ADRES, WAREHOUSETYPE, COUNTERTYPE, MAT_COUNTERTYPE, ORDERTYPE, vPAMCOR, BIR_YearEnd, ISSREPTYPE As String
Attribute NEYM.VB_VarUserMemId = 1073741836
Attribute ADRES.VB_VarUserMemId = 1073741836
Attribute WAREHOUSETYPE.VB_VarUserMemId = 1073741836
Attribute COUNTERTYPE.VB_VarUserMemId = 1073741836
Attribute MAT_COUNTERTYPE.VB_VarUserMemId = 1073741836
Attribute ORDERTYPE.VB_VarUserMemId = 1073741836
Attribute vPAMCOR.VB_VarUserMemId = 1073741836
Attribute BIR_YearEnd.VB_VarUserMemId = 1073741836
Attribute ISSREPTYPE.VB_VarUserMemId = 1073741836
Public BIRDATA_Connection, BIR_DATABASE_PATH           As String
Attribute BIRDATA_Connection.VB_VarUserMemId = 1073741844
Attribute BIR_DATABASE_PATH.VB_VarUserMemId = 1073741844
Public CSMS_PARTSQUERY                                 As Boolean
Attribute CSMS_PARTSQUERY.VB_VarUserMemId = 1073741846

Public TOTJOBAMT, TOTJOBDISC, TOTJOBTAX                As Double
Attribute TOTJOBAMT.VB_VarUserMemId = 1073741847
Attribute TOTJOBDISC.VB_VarUserMemId = 1073741847
Attribute TOTJOBTAX.VB_VarUserMemId = 1073741847
Public TOTPARTSAMT, TOTPARTSDISC, TOTPARTSTAX          As Double
Attribute TOTPARTSAMT.VB_VarUserMemId = 1073741850
Attribute TOTPARTSDISC.VB_VarUserMemId = 1073741850
Attribute TOTPARTSTAX.VB_VarUserMemId = 1073741850
Public TOTMATAMT, TOTMATDISC, TOTMATTAX                As Double
Attribute TOTMATAMT.VB_VarUserMemId = 1073741853
Attribute TOTMATDISC.VB_VarUserMemId = 1073741853
Attribute TOTMATTAX.VB_VarUserMemId = 1073741853
Public DNPIDFrom, DNPIDTo                              As Long
Attribute DNPIDFrom.VB_VarUserMemId = 1073741856
Attribute DNPIDTo.VB_VarUserMemId = 1073741856

Public PARTSQUERY                                      As Integer
Attribute PARTSQUERY.VB_VarUserMemId = 1073741865
Public rKeyDimension(1000)                             As Integer
Attribute rKeyDimension.VB_VarUserMemId = 1073741866
Public EncryptoFile(100000)                            As String
Attribute EncryptoFile.VB_VarUserMemId = 1073741867
Public CryptoKey                                       As Variant
Attribute CryptoKey.VB_VarUserMemId = 1073741868

Public VoiceMsg                                        As Boolean
Attribute VoiceMsg.VB_VarUserMemId = 1073741875
Public STOCK_TYPE                                      As String
Attribute STOCK_TYPE.VB_VarUserMemId = 1073741876
Public PRR_REPORT                                      As String
Attribute PRR_REPORT.VB_VarUserMemId = 1073741877

Public Const MAX_ISS_LINE = 35

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
Attribute Y_REGRESSION_INTERVAL.VB_VarUserMemId = 1073741841
Public Const X_MEAN_INTERVAL = 1

Public Const HARI_LEAD_TIME = 1.25
Public Const HARI_ORDER_FREQUENCY = 1.5

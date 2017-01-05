Attribute VB_Name = "modCSMIOSGlobals"
Option Explicit
Public OVERWRAYT                                                      As Boolean

Public LOGCODE                                                        As String
Public LOGNAME                                                        As String
Public LOGLEVEL                                                       As String
Public LOGTIME                                                        As String
Public LOGDATE                                                        As String
Public ROSHOW                                                         As Boolean
Public ESTISHOW                                                       As Boolean
Public ESTIKCNT                                                       As Integer

Public CUSCODE                                                        As String
Public LASTNEYM                                                       As String
Public FIRSTNEYM                                                      As String
Public MIDDLE                                                         As String
Public NEYM                                                           As String
Public ADRES                                                          As String
Public RO_OR_ESTI_OR_PART                                             As String
Public MAT_COUNTERTYPE                                                As String
Public VPAMCOR                                                        As String

Public TOTJOBAMT                                                      As Double
Public TOTJOBDISC                                                     As Double
Public TOTJOBDISCVAL                                                  As Double
Public TOTJOBTAX                                                      As Double
Public TOTPARTSAMT                                                    As Double
Public TOTPARTSDISC                                                   As Double
Public TOTPARTSDISCVAL                                                As Double
Public TOTPARTSTAX                                                    As Double
Public TOTMATAMT                                                      As Double
Public TOTMATDISC                                                     As Double
Public TOTMATDISCVAL                                                  As Double
Public TOTMATTAX                                                      As Double
Public Const VAT_RATE = 12

Public PMIS_Connection                                                As String
Public PMISBackUp_Connection                                          As String
Public CSMS_Connection                                                As String
Public SMIS_Connection                                                As String
Public CMIS_Connection                                                As String
Public AMIS_Connection                                                As String
Public LOGIN_Connection                                               As String

Public SEARCH_BY, SEARCH_TAB, COUNTERTYPE                             As String
Attribute SEARCH_BY.VB_VarUserMemId = 1073741871
Attribute SEARCH_TAB.VB_VarUserMemId = 1073741871
Attribute COUNTERTYPE.VB_VarUserMemId = 1073741871

Public PARTSQUERY, BILANG, BILANG2                                    As Long
Attribute PARTSQUERY.VB_VarUserMemId = 1073741874
Attribute BILANG.VB_VarUserMemId = 1073741874
Attribute BILANG2.VB_VarUserMemId = 1073741874
Public CSMS_PARTSQUERY                                                As Boolean
Attribute CSMS_PARTSQUERY.VB_VarUserMemId = 1073741877
Public DNPIDFrom, DNPIDTo                                             As Long
Attribute DNPIDFrom.VB_VarUserMemId = 1073741878
Attribute DNPIDTo.VB_VarUserMemId = 1073741878

Public wizVar, CryptVar                                               As Object
Attribute wizVar.VB_VarUserMemId = 1073741882
Attribute CryptVar.VB_VarUserMemId = 1073741882
Public AccessCNT                                                      As Integer
Attribute AccessCNT.VB_VarUserMemId = 1073741884
Public SearchBy, SEARCHCUSTOMERNAME, SEARCHPLATENO                    As String
Attribute SearchBy.VB_VarUserMemId = 1073741885
Attribute SEARCHCUSTOMERNAME.VB_VarUserMemId = 1073741885
Attribute SEARCHPLATENO.VB_VarUserMemId = 1073741885
Public Const MODULENAME = "CSMS"

Public QUESTION_TEST                                                  As String

Public EDIT_RO As String

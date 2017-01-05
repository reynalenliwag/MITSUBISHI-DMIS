Attribute VB_Name = "modCRISGlobal"
'UPDATED BY: JUN---------------------------------
'DATE UPDATED: 06-02-2009
Public QC_MODULE_ON                     As String
'UPDATED BY: JUN---------------------------------
Public SQL_STATEMENT                    As String
Public FROM_APPOINTMENT                 As String
Public ESTIKCNT                         As Integer
Public RO_OR_ESTI_OR_PART               As String
Public ROSHOW                           As Boolean
Public JOURNALTYPE                      As String
Public SEARCHCUSTOMERNAME               As String
Public SEARCHPLATENO                    As String
Public ESTISHOW                         As Boolean
Public QUESTION_TEST                    As String

Public Maxwiz                           As Long
Public AccessCNT                        As Long
Public wizVar                           As Object
Public CryptVar                         As Object
Public LOGCODE                          As String
Public LOGNAME                          As String
Public LOGLEVEL                         As String
Public LOGDATE                          As String
Public LOGTIME                          As String
Public PROC_TYPE                        As String
Public SKIN_PATH                        As String
Public SEARCH_TAB                       As String
Public MODULENAME
Public TOTJOBAMT                        As Double
Public TOTJOBDISC                       As Double
Public TOTJOBDISCVAL                    As Double
Public TOTJOBTAX                        As Double
Public TOTPARTSAMT                      As Double
Public TOTPARTSDISC                     As Double
Public TOTPARTSDISCVAL                  As Double
Public TOTPARTSTAX                      As Double
Public TOTMATAMT                        As Double
Public TOTMATDISC                       As Double
Public TOTMATDISCVAL                    As Double
Public TOTMATTAX                        As Double

Public TOTACCAMT, TOTACCDISC, TOTACCTAX                               As Double
Public TOTACCDISCVAL                                                  As Double

Public Const VAT_RATE = 12
Public EDIT_RO
Public PREPARED_BY
Public CHECKED_BY
Public APPROVED_BY
Public ACCOUNT_NO
Public BANK_MANAGER
Public SECRETARY
Public NOTED_BY
Public GENERAL_MANAGER
Public CUST_REPT_TYPE                   As String

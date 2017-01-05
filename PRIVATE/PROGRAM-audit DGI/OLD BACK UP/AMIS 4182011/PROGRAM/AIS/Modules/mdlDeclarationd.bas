Attribute VB_Name = "mdlAISDeclaration"
Option Explicit
Public SAVE_OR_EDIT                                                   As String

Public SAVE_OR_EDIT_TRAINING                                          As String
Public SAVE_OR_EDIT_PAPERS As String, SAVE_OR_EDIT_EMP                As String
Attribute SAVE_OR_EDIT_PAPERS.VB_VarUserMemId = 1073741831
Attribute SAVE_OR_EDIT_EMP.VB_VarUserMemId = 1073741831
Public SAVE_OR_EDIT_REF                                               As String
Attribute SAVE_OR_EDIT_REF.VB_VarUserMemId = 1073741833

Public PAPERS_ENTRY_ID As Integer, TRAINING_ENTRY_ID                  As Integer
Attribute PAPERS_ENTRY_ID.VB_VarUserMemId = 1073741834
Attribute TRAINING_ENTRY_ID.VB_VarUserMemId = 1073741834
Public EMP_ENTRY_ID                                                   As Integer
Public REFERENCE_ENTRY_ID                                             As Integer
Attribute REFERENCE_ENTRY_ID.VB_VarUserMemId = 1073741840

Public AIS_REPORT_PATH                                                As String
Attribute AIS_REPORT_PATH.VB_VarUserMemId = 1073741842
Public POSITION_ID                                                    As Long
Attribute POSITION_ID.VB_VarUserMemId = 1073741843
Public APPLICANT_ID                                                   As Long
Attribute APPLICANT_ID.VB_VarUserMemId = 1073741844
Public APPLICANT_TYPE                                                 As String
Attribute APPLICANT_TYPE.VB_VarUserMemId = 1073741845
Public Const INC_MSG = "Incomplete Entry Fields"
Public Const INC_TITLE = "Incomplete"

Public POSITION_DOC_ID                                                As Integer
Public POSITION_EXP_ID                                                As Integer
Public POSITION_EDU_ID                                                As Integer

Public POSITION_SAVE_OR_EDIT_EDU                                      As String
Attribute POSITION_SAVE_OR_EDIT_EDU.VB_VarUserMemId = 1073741849
Public POSITION_SAVE_OR_EDIT_EXP                                      As String
Public POSITION_SAVE_OR_EDIT_DOC                                      As String
Attribute POSITION_SAVE_OR_EDIT_DOC.VB_VarUserMemId = 1073741851

Public POSITION_EDU_ENTRY_ID                                          As Integer
Attribute POSITION_EDU_ENTRY_ID.VB_VarUserMemId = 1073741852
Public POSITION_DOC_ENTRY_ID                                          As Integer
Attribute POSITION_DOC_ENTRY_ID.VB_VarUserMemId = 1073741853

Public FROM_FORM_APPLY                                                As String
Attribute FROM_FORM_APPLY.VB_VarUserMemId = 1073741854

Public Const AIS_REPORT_Connection = "DSN=DMIS;DSQ=DMIS"

Public AIS_PICTURES_PATH                                              As String

Public Const AIS_MASTER_APPLICATION = 1030
Public Const AIS_PROCESS_SCHEDULEINTERVIEW = 1025
Public Const AIS_PROCESS_SCHEDULEEXAM = 1024
Public Const AIS_MASTER_TYPESOFEXAM = 1031
Public Const AIS_MASTER_OPENPOSITION = 1032
Public Const AIS_INQUIRY_SEARCH = 1029
Public Const AIS_PROCESS_UPLOAD = 1026
Public Const AIS_SYSTEM_EXIT = 1033

Attribute VB_Name = "modAudit"
'==========================================================================================
'FUNCTION / FEATURE :Opens SQL Connection to DMIS_AUDIT Database
'DATE STARTED       :5/16/200710:19
'LAST UPDATED       :5/16/200710:19
'DATABASE UPDATES   : Created Database DMIS_AUDIT Refer
'WHO UPDATED        :AXP5/16/200710:19
'UDPATING COCODE    :AXP516200710:19
'==========================================================================================
'A  =   ADD
'E  =   EDIT
'X  =   DELETED
'P  =   POSTED
'U  =   UNPOSTED
'C  =   CANCELLED
'V  =   VIEW
'R  =   PROCESS
'G  =   GENRATING
'I  =   INQUIRY
'==========================================================================================
'FUNCTION / FEATURE :LogAudit: Appends The Log Activity of Users
'DATE STARTED       :5/16/200712:57
'LAST UPDATED       :5/16/200712:57
'DATABASE UPDATES   : Created Table DMIS_AUDIT WHich Stores all the Log Activity Of Users
'WHO UPDATED        :AXP5/16/200712:57
'UDPATING COCODE    :AXP516200712:57
'==========================================================================================
' HOW TO USE IT
' SAMPLE FOR ADD  ON CUSTOMER MASTER FILE
' IT TAKES 2 REQ PARAMETER
'1. ACTION
'2. DESCRIPTION
'3. TRACKING ID
' SINCE IN DMIS WE USE MOSTLY GENERATED NUMBER IT IS VERY MUCH ADVISED TO USE
' SUCH CODE RATHER THAN USING ID (FIELD)
' CHECK PMIS FOR REFERENCE OR CUSTOMER MASTER FILE

'LogAudit "A", "CUSTOMER MASTER FILE", txtCusCode
' SAMPLE FOR EDIT ON CUSTOMER MASTER FILE
' LogAudit "E", "CUSTOMER MASTER FILE", txtCusCode
' SAMPLE FOR DELETE ON CUSTOMER MASTER FILE
' LogAudit "X", "CUSTOMER MASTER FILE", txtCusCode

Public gconAudit                        As ADODB.Connection
Public cmdAudit                         As ADODB.Command
'UPDATING CODE: AXP10/24/200720:53
'ADDED PUBLIC VARIABLE FOR TO KNOW WHERE AUDIT IS OK OR NOT IF ITS OK IT WILL KEEP INSERTING IT
'IF NOT IT WILL OMIT
Public AUDIT_OK                         As Boolean
Function OpenSQLAudit() As Boolean
    'UDPATING COCODE    :AXP516200710:19
    On Error GoTo errcde
    Set gconAudit = New ADODB.Connection
    gconAudit.Mode = adModeReadWrite
    gconAudit.CursorLocation = adUseClient
    gconAudit.ConnectionString = DMIS_Audit_Connection
    gconAudit.Open
    OpenSQLAudit = True
    AUDIT_OK = True
    Exit Function
errcde:
    OpenSQLAudit = False
    AUDIT_OK = False
End Function

Sub LogAudit(USER_ACTION As String, MODULE_NAME As String, Optional ByVal TrackingMemo As String = vbNullString)
    'UDPATING COCODE    :AXP516200712:57
    If AUDIT_OK = False Then Exit Sub
    If TrackingMemo <> vbNullString Then
        gconAudit.Execute ("INSERT INTO DMIS_AUDIT ( USER_ID, USER_ACTION,MODULE_NAME, ACTION_DATE,TRACKING_MEMO ) VALUES( " & LOGID & "," & N2Str2Null(USER_ACTION) & "," & N2Str2Null(MODULE_NAME) & " , getdate()," & N2Str2Null(TrackingMemo) & ")")
    Else
        gconAudit.Execute ("INSERT INTO DMIS_AUDIT ( USER_ID, USER_ACTION,MODULE_NAME, ACTION_DATE ) VALUES( " & LOGID & "," & N2Str2Null(USER_ACTION) & "," & N2Str2Null(MODULE_NAME) & " , getdate())")
    End If
End Sub



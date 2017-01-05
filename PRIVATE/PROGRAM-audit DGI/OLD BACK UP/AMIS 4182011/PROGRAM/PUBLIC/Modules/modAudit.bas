Attribute VB_Name = "modAudit"

'FUNCTION / FEATURE :Opens SQL Connection to DMIS_AUDIT Database
'DATE STARTED       :5/16/200710:19
'LAST UPDATED       :5/16/200710:19
'DATABASE UPDATES   : Created Database DMIS_AUDIT Refer
'WHO UPDATED        :AXP5/16/200710:19
'UDPATING COCODE    :AXP516200710:19

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

'FUNCTION / FEATURE :LogAudit: Appends The Log Activity of Users
'DATE STARTED       :5/16/200712:57
'LAST UPDATED       :5/16/200712:57
'DATABASE UPDATES   : Created Table DMIS_AUDIT WHich Stores all the Log Activity Of Users
'WHO UPDATED        :AXP5/16/200712:57
'UDPATING COCODE    :AXP516200712:57

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

Public gconAudit                                       As ADODB.Connection
Public cmdAudit                                        As ADODB.Command

'ADDED PUBLIC VARIABLE FOR TO KNOW WHERE AUDIT IS OK OR NOT IF ITS OK IT WILL KEEP INSERTING IT
'IF NOT IT WILL OMIT
Public AUDIT_OK                                        As Boolean

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
    On Error Resume Next
    If AUDIT_OK = False Then Exit Sub
    If TrackingMemo <> vbNullString Then
        gconAudit.Execute ("INSERT INTO DMIS_AUDIT ( USER_ID, USER_ACTION,MODULE_NAME, ACTION_DATE,TRACKING_MEMO ) VALUES( " & LOGID & "," & N2Str2Null(USER_ACTION) & "," & N2Str2Null(MODULE_NAME) & " , getdate()," & N2Str2Null(TrackingMemo) & ")")
    Else
        gconAudit.Execute ("INSERT INTO DMIS_AUDIT ( USER_ID, USER_ACTION,MODULE_NAME, ACTION_DATE ) VALUES( " & LOGID & "," & N2Str2Null(USER_ACTION) & "," & N2Str2Null(MODULE_NAME) & " , getdate())")
    End If
End Sub

Function GetBeforeValue(vID As Long, vTABLE As String) As String
    Dim rsKUTO                                         As New ADODB.Recordset
    Dim vResult                                        As String
    Set rsKUTO = gconDMIS.Execute("SELECT * FROM " & vTABLE & " WHERE ID = " & vID & "")
    'MsgBox rsKUTO.GetString(adClipString, , "-")
    If Not (rsKUTO.BOF And rsKUTO.EOF) Then
        '        Do While Not rsKUTO.EOF
        '        RO-1000 ADDED EDITED R0-1001
    Else
        '
    End If

    Set rsKUTO = Nothing
End Function

Sub NEW_LogAudit(xUSER_ACTION As String, xMODULENAME As String, xTRACKINGMEMO As String, xTRANID As String, XTYPE As String, XTRANNO As String, xTRANTYPE As String, xDETID As String)
'UDPATING COCODE    :MJP 06032008 04:54 PM
    On Error GoTo ErrorCode

    If AUDIT_OK = False Then Exit Sub
    gconAudit.Execute ("INSERT INTO DMIS_AUDIT (USER_ID, USER_ACTION, MODULE_NAME, ACTION_DATE, TRACKING_MEMO, TRANSACTION_ID, TYPE, TRANNO, TRANSTYPE, APPNAME, DETID) " & _
                       "VALUES(" & LOGID & _
                       "," & N2Str2Null(xUSER_ACTION) & _
                       "," & N2Str2Null(xMODULENAME) & _
                       ", getdate() " & _
                       ",'" & Replace(xTRACKINGMEMO, "'", "''") & "'" & _
                       "," & N2Str2Null(xTRANID) & _
                       "," & N2Str2Null(XTYPE) & _
                       "," & N2Str2Null(XTRANNO) & _
                       "," & N2Str2Null(xTRANTYPE) & _
                       "," & N2Str2Null(MODULENAME) & _
                       "," & N2Str2Null(xDETID) & ")")
    Exit Sub
ErrorCode:
    AUDIT_OK = False
End Sub

Function FindTransactionID(XTRANNO As Variant, xKEYFIELD As Variant, xTABLE As Variant, Optional XDETAILS As String, Optional XDETTYPE As String, Optional XDETFIELD As String) As String
    Dim rstmp                                          As New ADODB.Recordset
    If XDETAILS = "" Then
        Set rstmp = gconDMIS.Execute("SELECT ID FROM " & xTABLE & " WHERE " & xKEYFIELD & " = " & N2Str2Null(XTRANNO) & "")
        If Not (rstmp.BOF And rstmp.EOF) Then
            FindTransactionID = rstmp!ID
        Else
            FindTransactionID = ""
        End If
    Else
        Set rstmp = gconDMIS.Execute("SELECT ID FROM " & xTABLE & " WHERE " & xKEYFIELD & " = " & N2Str2Null(XTRANNO) & " AND " & XDETFIELD & " = " & N2Str2Null(XDETTYPE) & "")
        If Not (rstmp.EOF And rstmp.EOF) Then
            FindTransactionID = rstmp!ID
        Else
            FindTransactionID = ""
        End If
    End If
    Set rstmp = Nothing
End Function

Function FindUniqueKey(xUNIQUEKEY As String, xVALUETOFIND As String, xTABLE As String, xKEYFIELD As String) As String
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT " & xUNIQUEKEY & " AS UNIQUEKEY FROM " & xTABLE & " WHERE " & xKEYFIELD & " = " & xVALUETOFIND & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        FindUniqueKey = Null2String(rstmp!UNIQUEKEY)
    Else
        FindUniqueKey = ""
    End If
    Set rstmp = Nothing
End Function

Function FindNewID(xVOUCHERNO As Variant, xKEYFIELD1 As Variant, xTABLE As Variant, Optional xJType As String, Optional xKEYFIELD2 As String) As String
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT ID FROM " & xTABLE & " WHERE " & xKEYFIELD1 & " = " & xVOUCHERNO & " AND " & xKEYFIELD2 & " = " & xJType & "")
    If Not (rstmp.EOF And rstmp.EOF) Then
        FindNewID = rstmp!ID
    Else
        FindNewID = ""
    End If
    Set rstmp = Nothing
End Function

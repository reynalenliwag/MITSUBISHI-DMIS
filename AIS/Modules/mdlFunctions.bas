Attribute VB_Name = "mdlAISFunctions"
Option Explicit

Function GetTime_TMP(Index As Integer) As String
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select SET_TIME From HRMS_TIME3 Where TIME_ID = " & Index & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetTime_TMP = RSTMP!Set_Time
    End If
End Function

Public Sub FillCBOTime(cbo As ComboBox)
    cbo.AddItem "8:00 - 9:00 AM"
    cbo.AddItem "9:00 - 10:00 AM"
    cbo.AddItem "10:00 - 11:00 AM"
    cbo.AddItem "11:00 - 12:00 AM"
    cbo.AddItem "1:00 - 2:00 PM"
    cbo.AddItem "2:00 - 3:00 PM"
    cbo.AddItem "3:00 - 4:00 PM"
    cbo.AddItem "4:00 - 5:00 PM"

    cbo.ListIndex = 0
End Sub

Public Sub fillcombo_up(cbx As ComboBox)
    
'    cbx.AddItem "2007"
'    cbx.AddItem "2008"
'    cbx.AddItem "2009"
'    cbx.AddItem "2010"
'    cbx.AddItem "2011"
'    cbx.AddItem "2012"
'    cbx.AddItem "2013"
'    cbx.AddItem "2014"
'    cbx.AddItem "2015"
    
        
        Dim I As Integer
        
        For I = 2005 To 2015
               cbx.AddItem I
        Next I
       
    cbx.Text = Format(LOGDATE, "yyyy")
    
End Sub



Function ChangeColor(MYCOLOR As Long) As Long
    If MYCOLOR = 16711680 Then MYCOLOR = 0: ChangeColor = MYCOLOR: Exit Function        'Blue   - Black
    If MYCOLOR = 0 Then MYCOLOR = 225: ChangeColor = MYCOLOR: Exit Function             'Black  - Red
    If MYCOLOR = 225 Then MYCOLOR = &H8000&: ChangeColor = MYCOLOR: Exit Function       'Red    - Green
    If MYCOLOR = &H8000& Then MYCOLOR = &H40C0&: ChangeColor = MYCOLOR: Exit Function   'Green  - Orange
    If MYCOLOR = &H40C0& Then MYCOLOR = 16711680: ChangeColor = MYCOLOR: Exit Function  'Orange - Blue
End Function

Function GetExamType(EXAMTYPE As Integer) As String
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select ExamDescription From HRMS_ExamType Where ExamID = " & EXAMTYPE & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetExamType = RSTMP!ExamDescription
    End If
End Function

Public Sub DisplayAMessage(Msg As String, TITLE As String)
    MsgBox "" & Msg & "", vbExclamation, "" & TITLE & ""
End Sub

Public Sub GenerateNewID(TABLE As String, ID As Integer)
    Dim SQL                                                           As String
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select Entry_ID From " & TABLE & " Where Applicant_ID = " & APPLICANT_ID & _
                               " Order By Entry_ID ASC")

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            ID = RSTMP!Entry_ID
            RSTMP.MoveNext
        Loop
    End If
    ID = ID + 1
    Set RSTMP = Nothing
End Sub

Public Function GetRS(strSQL As String) As Recordset
'    Dim oRs                                                           As Recordset
'    Set oRs = New ADODB.Recordset
'    oRs.CursorLocation = adUseClient
'    oRs.Open strSQL, gconDMIS, adOpenForwardOnly, adLockReadOnly
'    Set oRs.ActiveConnection = Nothing
'    Set gconDMIS.Execute = oRs
'    Set oRs = Nothing
End Function

Function FindApplicantName(APP_ID As Integer) As String
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select FirstName,LastName From HRMS_APPLICANT_PERSONAL Where Applicant_ID = " & APP_ID & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindApplicantName = RSTMP!lastname & "," & RSTMP!FIRSTNAME
    End If
End Function

Function ReturnExamType(ID As Integer) As String
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select ExamDescription From HRMS_EXAMTYPE Where ExamID = " & ID & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        ReturnExamType = RSTMP!ExamDescription
    End If
End Function

'Sub DMIS_VERSION()
'On Error GoTo FALSEUSERS
'    'UPDATED BY : ACL
'    'DATE       : 02022011
'    'DESCRIPTION: TO CHECK LATEST VERSION OF APPLICATION
'    Dim rsALLPROFILE                                   As ADODB.Recordset
'    Dim rsUSERNAME As ADODB.Recordset
'    Dim SQL                                            As String
'    Dim SQL1                                           As String
'    Dim SQL2                                           As String
'    Dim SQL3                                           As String
'    Dim SQL4                                            As String
'
'    If COMPANY_CODE = COMPANY_VERSION Then
'        SQL = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_RAMS_USERS' AND COLUMN_NAME='USER_NAME') " & vbCrLf
'        SQL = SQL & "SELECT USER_NAME FROM ALL_RAMS_USERS"
'        Set rsUSERNAME = gconACCESS.Execute(SQL)
'        If Not rsUSERNAME.EOF And Not rsUSERNAME.BOF Then
'        Else
'FALSEUSERS:
'            SQL1 = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_RAMS_USERS' AND COLUMN_NAME='USERNAME') " & vbCrLf
'            SQL1 = SQL1 & "EXEC SP_RENAME 'ALL_RAMS_USERS.USERNAME','USER_NAME','COLUMN'"
'            gconACCESS.Execute SQL1
'
'            SQL2 = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_RAMS_USERS' AND COLUMN_NAME='USER_NAME') " & vbCrLf
'            SQL2 = "ALTER VIEW ALL_VW_RAMS_PACCESS " & vbCrLf
'            SQL2 = SQL2 & "AS " & vbCrLf
'            SQL2 = SQL2 & "SELECT USERID,USER_NAME,PASSWORD AS USERPASS,USERGROUP AS LOGLEVEL, USERCODE, LOCK " & vbCrLf
'            SQL2 = SQL2 & "From DBO.ALL_RAMS_USERS"
'            gconACCESS.Execute SQL2
'
'            SQL3 = "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_PROFILE' AND COLUMN_NAME='VERSION') " & vbCrLf
'            SQL3 = SQL3 & "ALTER TABLE ALL_PROFILE " & vbCrLf
'            SQL3 = SQL3 & " ADD VERSION INT"
'            gconACCESS.Execute SQL3
'
'            SQL4 = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_RAMS_USERS' AND COLUMN_NAME='USER_NAME') " & vbCrLf
'            SQL4 = "ALTER VIEW ALL_VW_USERACESS " & vbCrLf
'            SQL4 = SQL4 & "AS" & vbCrLf
'            SQL4 = SQL4 & "SELECT  ARU.USERID,ALL_RAMS_USERS.USER_NAME,ARU.MODULEID,ARM.MAINMODULENAME,ARM.DESCRIPTIONS,ARM.MODULE_TYPE, " & vbCrLf
'            SQL4 = SQL4 & "ARU.ACESS_ADD,ARU.ACESS_EDIT,ARU.ACESS_DELETE,ARU.ACESS_VIEW,ARU.ACESS_PRINT,ARU.ACESS_PROCESS,ARU.ACESS_SYSTEM,ARU.ACESS_POST, " & vbCrLf
'            SQL4 = SQL4 & "ARU.ACESS_UNPOST , ARU.ACESS_CANCELENTRY " & vbCrLf
'            SQL4 = SQL4 & "FROM ALL_RAMS_USERSACESS AS ARU INNER JOIN " & vbCrLf
'            SQL4 = SQL4 & "ALL_RAMS_MODULES AS ARM ON ARU.MODULEID = ARM.MODULEID INNER JOIN " & vbCrLf
'            SQL4 = SQL4 & "ALL_RAMS_USERS ON ARU.USERID = ALL_RAMS_USERS.USERID"
'            gconACCESS.Execute SQL4
'        End If
'
'        Set rsALLPROFILE = New ADODB.Recordset
'        rsALLPROFILE.Open "SELECT ISNULL(VERSION,0) AS VERSION FROM ALL_PROFILE WHERE MODULENAME='" & App.EXEName & "'", gconACCESS, adOpenForwardOnly
'        If Not rsALLPROFILE.EOF And Not rsALLPROFILE.BOF Then
'            If rsALLPROFILE!Version < App.Revision Then
'                gconACCESS.Execute ("UPDATE ALL_PROFILE SET VERSION = '" & App.Revision & "' WHERE MODULENAME='" & App.EXEName & "'")
'            ElseIf rsALLPROFILE!Version > App.Revision Then
'                MsgBox "You are using old " & App.EXEName & " version." & Chr(13) & "Please ask the administrator for the latest update!"
'                End
'            End If
'        End If
'        Set rsALLPROFILE = Nothing
'    End If
'End Sub
'
'Sub GET_COMPANYCODE()
'    Dim rsALLPROFILE                                   As ADODB.Recordset
'    Set rsALLPROFILE = New ADODB.Recordset
'    rsALLPROFILE.Open "SELECT ISNULL(COMPANYCODE,'') AS COMPANYCODE FROM ALL_PROFILE WHERE MODULENAME='" & App.EXEName & "'", gconACCESS, adOpenForwardOnly
'    If Not rsALLPROFILE.EOF And Not rsALLPROFILE.BOF Then
'        COMPANY_CODE = rsALLPROFILE!COMPANYCODE
'    End If
'    Set rsALLPROFILE = Nothing
'End Sub
'
'Sub COMPANYCODE_VERSION()
'    Dim CTR                                            As Integer
'    Dim xCTR As Integer
'    CTR = 1
'    ReDim company(CTR) As String
'    company(0) = "ACL"
'    company(1) = "HHH"
'
'    For xCTR = 0 To CTR
'        If company(xCTR) = COMPANY_CODE Then
'            COMPANY_VERSION = company(xCTR)
'        End If
'    Next
'End Sub
'
'Function CHANGE_USER() As Boolean
'On Error GoTo FALSEUSER
'    Dim SQL As String
'    Dim rsUSERNAME As ADODB.Recordset
'    SQL = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_RAMS_USERS' AND COLUMN_NAME='USER_NAME') " & vbCrLf
'    SQL = SQL & "SELECT TOP 1 USER_NAME FROM ALL_RAMS_USERS"
'    Set rsUSERNAME = gconACCESS.Execute(SQL)
'    If Not rsUSERNAME.EOF And Not rsUSERNAME.BOF Then
'        CHANGE_USER = True
'    Else
'FALSEUSER:
'        CHANGE_USER = False
'    End If
'End Function
'

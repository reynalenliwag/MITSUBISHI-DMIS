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
    
        
        Dim i As Integer
        
        For i = 2005 To 2015
               cbx.AddItem i
        Next i
       
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

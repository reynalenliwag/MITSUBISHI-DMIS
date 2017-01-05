Attribute VB_Name = "modDbHelper"
Option Explicit
Function GetCommand(nType As ADODB.CommandTypeEnum, nText As String) As ADODB.Command
    Connect
    Dim tempCmd                              As ADODB.Command
    Set tempCmd = New ADODB.Command
    tempCmd.ActiveConnection = gconDMIS
    tempCmd.CommandType = nType
    tempCmd.CommandText = nText
    Set GetCommand = tempCmd
    Set tempCmd = Nothing
End Function
Public Function GetRS(strSQL As String) As Recordset
    On Error GoTo adder:
    Dim oRS                                  As Recordset
    Set oRS = New ADODB.Recordset
    oRS.CursorLocation = adUseClient
    oRS.Open strSQL, gconDMIS, adOpenForwardOnly, adLockReadOnly
    '    Set oRS.ActiveConnection = Nothing
    Set GetRS = oRS
    Set oRS = Nothing
    Exit Function
adder:
    Call Err.Raise(1, 1, "Check your SQL Syntax" & vbCrLf & Err.Description)
End Function

Function GetString(ByVal strx As Variant) As String
    If IsNull(strx) = True Then
        GetString = Empty
    Else
        GetString = Trim(strx)
    End If
End Function
Function GetDouble(ByVal strx As Variant) As Double
    If IsNumeric(strx) = False Then
        GetDouble = 0
    Else
        GetDouble = CDbl(strx)
    End If
End Function
Function GetLong(ByVal strx As Variant) As Long
    If Len(Trim$(strx)) = 0 Or IsNumeric(strx) = False Then
        GetLong = 0
    Else
        GetLong = CLng(strx)
    End If
End Function
Public Function BinToBoolean(nVal As Long) As Boolean
    If nVal = 0 Then
        BinToBoolean = False
    Else
        BinToBoolean = True
    End If
End Function
Public Function BooleanToBin(bVal As Boolean) As Long
    If bVal Then
        BooleanToBin = 1
    Else
        BooleanToBin = 0
    End If
End Function
Public Function ColorToStr(ByVal clr As OLE_COLOR) As String
    Dim strColor                             As String
    strColor = clr Mod 256
    strColor = strColor & ", " & (clr \ 256 Mod 256)
    strColor = strColor & ", " & (clr \ 256 \ 256 Mod 256)
    ColorToStr = strColor
End Function


Public Function GetRecord(nSQL As String) As String
    Dim TempRs                               As ADODB.Recordset

    Set TempRs = gconDMIS.Execute(nSQL)
    If Not TempRs.EOF Or TempRs.BOF Then
        GetRecord = TempRs.Collect(0)
    End If
    Set TempRs = Nothing

End Function

Public Sub Connect()
    On Error GoTo ConError:
    Dim Error                                As ADODB.Error
    If gconDMIS.State = adStateOpen Then: Exit Sub
    gconDMIS.CursorLocation = adUseServer
    gconDMIS.ConnectionString = SQLConnectionString
    If gconDMIS.State = adStateClosed Then
        gconDMIS.Open
    End If
    Exit Sub
ConError:
    For Each Error In gconDMIS.Errors
        Select Case Error.NativeError
            Case "2003", 17
                Call MsgBox("Could Not Connect To ""SQL SERVER"". " _
                          & vbCrLf & "Please Try to Start Your Service and Try Again. " _
                          & vbCrLf & "Application Will Now Exit." _
                   , vbCritical Or vbSystemModal Or vbMsgBoxHelpButton, "Connection Error", App.HelpFile, 100001)
            Case "1045"
                Call MsgBox("Cannot Connect To My SQL Database. " _
                          & vbCrLf & "Invalid Credential for Database. Please Check your Users name and Password. " _
                          & vbCrLf & "Application Will Now Exit." _
                   , vbCritical Or vbSystemModal Or vbMsgBoxHelpButton, "Invalid Credential", App.HelpFile, 10003)
            Case "1049"
                Call MsgBox("Invalid Database for the Application." _
                          & vbCrLf & "Please Contact Support. Application Will Now Exit" _
                   , vbCritical, "Fatal Error")
            Case Else

                Call MsgBox("Fatal Error. Could Not Connect To ""SQL SERVER""" _
                          & vbCrLf & "Please Contact Support. Application Will Now Exit" _
                   , vbCritical, "Fatal Error")

        End Select
        Set gconDMIS = Nothing: End
    Next
End Sub



Function GetCustomerCode(lastname As String) As String
Dim TempRs As ADODB.Recordset
If Len(lastname) = 0 Then
    Exit Function
End If
    Dim lAlpha As String
    lAlpha = Left(Trim(lastname), 1)
        Set TempRs = gconDMIS.Execute("Select CTLCDE From ALL_CUSCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (TempRs.EOF Or TempRs.BOF) Then
        GetCustomerCode = Left(lastname, 1) & Format(Mid(TempRs.Collect(0), 2, 5), "00000")
    End If
End Function
Sub SetCustomerCode(lastname As String)
Dim SQLX As String
SQLX = "Update ALL_CUSCTL SET CTLCDE='" & lastname & "'" _
            & " Where LEFT(CTLCDE,1)='@AX'"
        SQLX = Replace(SQLX, "@AX", Left(lastname, 1))
        gconDMIS.Execute SQLX
End Sub



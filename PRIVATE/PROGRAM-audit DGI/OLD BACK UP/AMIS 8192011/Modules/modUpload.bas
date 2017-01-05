Attribute VB_Name = "modUpload"
Public Conn         As ADODB.Connection
Public Conn2         As ADODB.Connection
Public xExcelFile   As String

Sub Connect()
    Set Conn = New ADODB.Connection
    Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xExcelFile & ";Extended Properties=Excel 8.0;Persist Security Info=False"
    Conn.ConnectionTimeout = 40
    Conn.Open
    
'    Set Conn2 = New ADODB.Connection
'    Conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DPCDATA.MDB" & ";Persist Security Info=False"
'    Conn2.Open
End Sub

Public Function NumericVal(NumericText As Variant) As Double
    Dim counter                                   As Integer
    Dim NumericValue                              As String
    NumericValue = ""
    If Trim(NumericText) <> "" Then
        If IsNumeric(NumericText) = True Then
            If Val(NumericText) >= 0 Then
                For counter = 1 To Len(NumericText)
                    If Mid(NumericText, counter, 1) <> "," Then
                        NumericValue = NumericValue & Mid(NumericText, counter, 1)
                    End If
                Next
                NumericVal = NumericValue
            Else
                NumericVal = Val(NumericText)
            End If
        Else
            NumericVal = 0
        End If
    Else
        NumericVal = 0
    End If
End Function


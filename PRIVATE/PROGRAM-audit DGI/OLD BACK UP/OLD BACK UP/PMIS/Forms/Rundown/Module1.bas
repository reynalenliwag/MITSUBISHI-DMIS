Attribute VB_Name = "Module1"
Public Function NumericVal(NumericText As Variant) As Double
Dim Counter As Integer
Dim NumericValue As String
NumericValue = ""
If Trim(NumericText) <> "" Then
   If IsNumeric(NumericText) = True Then
      For Counter = 1 To Len(NumericText)
          If Mid(NumericText, Counter, 1) <> "," Then
             NumericValue = NumericValue & Mid(NumericText, Counter, 1)
          End If
      Next
      NumericVal = NumericValue
   Else
      NumericVal = 0
   End If
Else
   NumericVal = 0
End If
End Function


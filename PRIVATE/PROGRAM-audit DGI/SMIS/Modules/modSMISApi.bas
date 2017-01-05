Attribute VB_Name = "modSMISApi"
Option Explicit
Private Const CB_LIMITTEXT = &H141
Private Const CB_FINDSTRING = &H14C
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_ERR = (-1)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000

Public Sub SetComboMaxLength(ComboBox As ComboBox, ByVal lMaxLength As Long)
    SendMessage ComboBox.hwnd, CB_LIMITTEXT, lMaxLength, ByVal 0&
End Sub
Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
                                                                                   As ColumnHeader = Nothing)

    Dim C                                                             As ColumnHeader
    If LV.ListItems.Count = 0 Then Exit Sub
    If Column Is Nothing Then
        For Each C In LV.ColumnHeaders
            SendMessage LV.hwnd, LVM_FIRST + 30, C.Index - 1, -1
        Next
    Else
        SendMessage LV.hwnd, LVM_FIRST + 30, Column.Index - 1, -1
    End If
    LV.Refresh
End Sub

Public Function SetComboIndex(C As ComboBox) As Long
    Dim i                                                             As Long
    Dim STR                                                           As String
    STR = C.Text
    i = SendMessage(C.hwnd, CB_FINDSTRING, -1, ByVal STR)
    If i <> CB_ERR Then
        SetComboIndex = i
    Else
        SetComboIndex = -1
    End If
End Function
Public Sub SetComboWidth(C As ComboBox, xWidth As Long)
    Call SendMessage(C.hwnd, CB_SETDROPPEDWIDTH, xWidth, 0)
End Sub

Public Sub SetComboSelect(C As ComboBox)
    Dim i                                                             As Long
    Dim j                                                             As Long
    Dim strPartial                                                    As String
    Dim strTotal                                                      As String
    strPartial = C.Text
    i = SendMessage(C.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
    If i <> CB_ERR Then
        strTotal = C.List(i)
        j = Len(strTotal) - Len(strPartial)
        If j <> 0 Then
            C.SelText = Right$(strTotal, j)
            C.SelStart = Len(strPartial)
            C.SelLength = j
        End If
    End If
End Sub

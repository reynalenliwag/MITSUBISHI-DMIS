Attribute VB_Name = "modPopUps"

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Enum PopIcon
    Delete = 900
    InfoFriend = 901
    InfoHelp = 902
    InfoOk = 903
    InfoStop = 904
    InfoVoid = 905
    InfoWait = 906
    InfoWarning = 907
    NaviEnd = 908
    NaviBegin = 909
    NoEntry = 910
    RecLocekd = 911
    RecSave = 912
    RecSaveError = 913
    RecSaveInfo = 914
    RecSaveOk = 915
    RecSaveWarning = 916
    RecNotFound = 919
    Refresh = 917
    Star = 918
    NONE = 0
End Enum
Public Type RECT
    Left                                     As Long
    Top                                      As Long
    Right                                    As Long
    Bottom                                   As Long
    Center                                   As Long
End Type
Public Type POINTAPI
    x                                        As Long
    y                                        As Long
End Type

Public Sub MessagePop(PopIcon As PopIcon, Title As String, Message As String, Optional ByVal Interval As Integer = 2000, Optional ByVal Position As Integer = 0, Optional ByVal heightx As Integer = 95)

    Dim strMsg()                             As String
    Dim REC                                  As RECT
    strMsg = Split(Message, "^")
    frmMain.PopCntrl.Close
    With frmMain.PopCntrl
        If heightx > 0 Then
            .SetSize 270, heightx
        End If
        .ShowDelay = Interval
        .Item(1).caption = Title
        If PopIcon <> NONE Then
            .Item(2).IconIndex = PopIcon
        Else
            .Item(2).IconIndex = 0
            .Item(3).Left = 10
            .Item(3).Width = 260
        End If
        .Item(3).caption = strMsg(0)

        If UBound(strMsg) = 1 Then
            .Item(4).caption = strMsg(1)
        Else
            .Item(3).Height = heightx - 40
        End If

        If IsModal Then
            Dim frm                          As Form
            For Each frm In Forms
                frm.Enabled = False
            Next
        End If

        If Position = 1 Then
            GetWindowRect frmMain.hwnd, REC
            .Right = REC.Right - ((frmMain.ScaleWidth \ Screen.TwipsPerPixelX) \ 2) + (frmMain.PopCntrl.Width \ 2)
            .Bottom = (REC.Bottom) - ((frmMain.ScaleHeight \ Screen.TwipsPerPixelY) \ 2) + (frmMain.PopCntrl.Height \ 2)
        ElseIf Position = 2 Then
            GetWindowRect frmMain.hwnd, REC
            .Right = REC.Right - 5
            .Bottom = (REC.Bottom) - 5
        End If
        .Show
    End With

End Sub


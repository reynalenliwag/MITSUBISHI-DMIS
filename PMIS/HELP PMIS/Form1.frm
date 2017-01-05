VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Res File"
      Height          =   585
      Left            =   360
      TabIndex        =   3
      Top             =   1950
      Width           =   3345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate FILES"
      Height          =   585
      Left            =   360
      TabIndex        =   2
      Top             =   1380
      Width           =   3345
   End
   Begin VB.CommandButton CMD2 
      Caption         =   "Command1"
      Height          =   585
      Left            =   360
      TabIndex        =   1
      Top             =   810
      Width           =   3345
   End
   Begin VB.CommandButton CMD1 
      Caption         =   "Command1"
      Height          =   585
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
'                                  (ByVal hwndCaller As Long, ByVal pszFile As String, _
'                                   ByVal uCommand As Long, ByVal dwData As Long) As Long
'Private Declare Function HTMLHelpTopic Lib "hhctrl.ocx" Alias "HtmlHelpA" _
'                                       (ByVal hwndCaller As Long, ByVal pszFile As String, _
'                                        ByVal uCommand As Long, ByVal dwData As String) As Long
'Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)
'Private Const HH_DISPLAY_TOPIC = &H0
'Private Const HH_CLOSE_ALL = &H12
'Const HH_HELP_CONTEXT                                  As Long = &HF
'Private Sub cmdExit_Click()
'    Unload Me
'End Sub
'
'Private Sub CMD1_Click()
'    hwRes = HtmlHelp(Me.hWnd, "pmis.chm", HH_DISPLAY_TOC, 0)
'End Sub

Private Sub CMD2_Click()
    MsgBox "HI HELP", vbMsgBoxHelpButton, "", "pmis.chm", 102
End Sub

Private Sub CMD3_Click()
    MsgBox "HI HELP", vbMsgBoxHelpButton, "", "pmis.chm", 102
End Sub


Private Sub Command1_Click()
    Dim XLAPP                                          As New Excel.Application
    Dim XLBook                                         As New Excel.Workbook
    Dim XLsheet                                        As New Excel.Worksheet
    Set XLBook = XLAPP.Workbooks.Open(App.Path & "\help.xls")
    Set XLsheet = XLBook.Worksheets(1)


    Dim HTML_1                                         As String
    Dim HTML_2                                         As String
    Dim HTML_3                                         As String
    Dim HTML_4                                         As String
    Dim INX                                            As Integer
    HTML_1 = XLsheet.Cells(2, "A")
    INX = 2


    Do While HTML_1 <> ""
        HTML_1 = XLsheet.Cells(INX, "A")
        HTML_2 = XLsheet.Cells(INX, "C")
        HTML_3 = XLsheet.Cells(INX, "D")
        Open App.Path & "\DOCS\HELP" & Format(INX, "000000") & ".htm" For Output As #1
        Print #1, "<html xmlns=""http://www.w3.org/1999/xhtml"">"
        Print #1, "<head>"
        Print #1, "<meta http-equiv=""Content-Type"" content=""text/html"" charset=UTF-8"">"
        Print #1, "<STYLE>"
        Print #1, ".yellow"
        Print #1, "{"
        Print #1, "border: solid 1px #DEDEDE;"
        Print #1, "background: #FFFFCC;"
        Print #1, "color: #222222;"
        Print #1, "padding: 4px;"
        Print #1, "text-align: left;"
        Print #1, "}"

        Print #1, "body , td"
        Print #1, "{"
        Print #1, "font-family: ""Lucida Grande"" , ""Lucida Sans Unicode"" , Verdana, Arial, Helvetica, sans-serif;"
        Print #1, "font-size: 12px;"
        Print #1, "margin: 0;"
        Print #1, "border: 0;"
        Print #1, "padding: 0;"
        Print #1, "}"

        Print #1, ".Gray"
        Print #1, "{"
        Print #1, "border: solid 1px #DEDEDE;"
        Print #1, "background: #EFEFEF;"
        Print #1, "color: #222222;"
        Print #1, "padding: 4px;"
        Print #1, "font-weight: bolder;"
        Print #1, "text-align: center;"
        Print #1, "}"
        Print #1, "</STYLE>"
        Print #1, "</head>"
        Print #1, "<body>"
        Print #1, "<table>"
        Print #1, "<tbody>"
        Print #1, "<tr>"
        Print #1, "<td>"
        Print #1, "<img src=""Images/logo.gif"">"
        Print #1, "</td>"
        Print #1, "</tr>"
        Print #1, "<tr>"
        Print #1, "<td class=""Gray"">"
        Print #1, HTML_1
        Print #1, "</td>"
        Print #1, "</tr>"
        Print #1, "<tr>"
        Print #1, "<td class=""yellow"">"
        Print #1, "<p>"
        Print #1, HTML_2
        Print #1, "</p>"
        Print #1, "<p>"
        Print #1, HTML_3
        Print #1, "</p>"
        Print #1, "</td>"
        Print #1, "</tr>"
        Print #1, "</tbody>"
        Print #1, "</table>"
        Print #1, "</body>"
        Print #1, "</html>"
        Close #1
        INX = INX + 1
    Loop

    HTML_1 = XLsheet.Cells(2, "A")
    INX = 2

    Open App.Path & "\MAP.h" For Output As #1

    Do While HTML_1 <> ""
        HTML_1 = XLsheet.Cells(INX, "A")
        Print #1, "#define DOCS\HELP" & Format(INX, "000000") & " " & (INX + 100)
        INX = INX + 1
    Loop
    Close #1

    HTML_1 = XLsheet.Cells(2, "A")
    INX = 2

    Open App.Path & "\PMIS.hhp" For Output As #1

    Print #1, "[Options]"
    Print #1, "Compatibility = 1.1 Or later"
    Print #1, "Compiled file = PMIS.chm"
    Print #1, "Default topic = docs\Default.htm"
    Print #1, "Display compile progress=No"
    Print #1, "Language=0x409 English (United States)"
    Print #1, ""
    Print #1, "[Files]"


    Do While HTML_1 <> ""
        HTML_1 = XLsheet.Cells(INX, "A")
        Print #1, "DOCS\HELP" & Format(INX, "000000") & ".htm"
        INX = INX + 1
    Loop
    Print #1, ""
    Print #1, "[Map]"
    Print #1, "#include Map.h"
    Print #1, "[INFOTYPES]"
    Close #1

















    Set XLsheet = Nothing
    XLBook.Close
    Set XLAPP = Nothing

End Sub

Private Sub Command2_Click()
 
    Dim B()                                            As Byte
    Dim s                                              As String
    Dim i                                              As Long

    Dim Temp                                           As String
    Dim StartPosition                                  As Long
    Dim mHandle                                        As Integer

    s = ""
    B = LoadResData(101, "CUSTOM")
    For i = 0 To UBound(B())
        s = s & Chr(B(i))
    Next i
    Erase B
    
    mHandle = FreeFile
    Open "C:\PMIS.CHM" For Binary As #mHandle
    StartPosition = LOF(mHandle)
    Temp = s
    Put #mHandle, , Temp
    Put #mHandle, , StartPosition
    Close #mHandle

    'webbrowser1.Navigate "C: emp.htm"
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Call HtmlHelp(Me.hWnd, "", HH_CLOSE_ALL, 0)
'    Sleep 1000
End Sub



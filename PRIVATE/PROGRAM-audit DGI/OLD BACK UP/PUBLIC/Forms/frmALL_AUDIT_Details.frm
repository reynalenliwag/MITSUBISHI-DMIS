VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmALL_AUDIT_Details 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Audit Record Details"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   ClipControls    =   0   'False
   Icon            =   "frmALL_AUDIT_Details.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   9225
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3555
      Left            =   990
      ScaleHeight     =   3525
      ScaleWidth      =   7485
      TabIndex        =   2
      Top             =   690
      Visible         =   0   'False
      Width           =   7515
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3105
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmALL_AUDIT_Details.frx":030A
         Top             =   360
         Width           =   7395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7110
         TabIndex        =   3
         Top             =   0
         Width           =   345
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   -30
         TabIndex        =   5
         Top             =   0
         Width           =   7515
         _Version        =   655364
         _ExtentX        =   13256
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "SQL Statment"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin SHDocVwCtl.WebBrowser wbName 
      Height          =   5145
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   9075
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   " | Press F3 To Print The Information | Press F1 To View Technical Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   5280
      Width           =   9105
   End
End
Attribute VB_Name = "frmALL_AUDIT_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function SHOW_AUDITDETAILS(labid As Single) As Boolean
    On Error GoTo ERROR_CODE:
    Dim RSMEMO                                         As ADODB.Recordset
    Set RSMEMO = gconAudit.Execute("SELECT ACTION_DATE, DESCRIPTION,USERACTION,TRACKING_MEMO, USERNAME,TRANNO FROM ALL_VW_AUDIT WHERE ID = " & labid)
    If (RSMEMO.EOF Or RSMEMO.BOF) Then
        Exit Function
    End If


    Dim valid_datalines                                As String
    Dim STR                                            As String
    Dim C_CHAR()                                       As String
    Dim D_CHAR()                                       As String
    Dim i                                              As Integer
    Dim FIRSTPOSITION
    Dim SECONDPOSITON                                  As Integer
    Dim THIRDPOSITION                                  As Integer
    Dim FOURTHPOSITION                                 As Integer
    Dim COLUMNLINES                                    As String
    Dim DATALINES                                      As String
    Dim T_COL                                          As Integer
    Dim D_COL                                          As Integer
    Dim COMMAND_TYPE                                   As String
    Dim zzz                                            As Integer
    Dim is_inside                                      As Boolean
    Dim var_ACTION                                     As String
    Dim Date_Value
    Dim time_value
    Dim STRHTML                                        As String
    STR = Null2String(RSMEMO!TRACKING_MEMO)
    Text1 = STR

    If InStr(1, UCase(STR), "WHERE") > 0 And InStr(1, UCase(STR), "UPDATE") > 0 Then
        COMMAND_TYPE = "UPDATE"
    ElseIf InStr(1, UCase(STR), "INSERT") > 0 Then
        COMMAND_TYPE = "INSERT"
    Else
        If Null2String(RSMEMO!USERACTION) = "V" Then
        
        ElseIf Null2String(RSMEMO!USERACTION) = "CL" Then
        Else
            SHOW_AUDITDETAILS = False
            Exit Function
        End If
    End If

    var_ACTION = GetUserAction(UCase(LTrim(RTrim(Null2String(RSMEMO!USERACTION)))))

    If IsDate((RSMEMO!ACTION_DATE)) = True Then
        Date_Value = DateValue(RSMEMO!ACTION_DATE)
        time_value = TimeValue(RSMEMO!ACTION_DATE)
    End If

    If COMMAND_TYPE = "INSERT" Then
        FIRSTPOSITION = InStr(1, STR, "(")
        SECONDPOSITON = InStr(FIRSTPOSITION, STR, ")")
        THIRDPOSITION = InStr(SECONDPOSITON, STR, "(")
        FOURTHPOSITION = InStr(THIRDPOSITION, STR, ")")
        COLUMNLINES = Mid(STR, FIRSTPOSITION + 1, SECONDPOSITON - FIRSTPOSITION - 1)
        DATALINES = Mid(STR, THIRDPOSITION + 1, FOURTHPOSITION - THIRDPOSITION - 1)
        valid_datalines = ""
        is_inside = False
        For zzz = 1 To Len(DATALINES)
            If Mid(DATALINES, zzz, 1) = "'" And is_inside = False Then
                is_inside = True
                zzz = zzz + 1
                If Mid(DATALINES, zzz, 1) = "'" And is_inside = True Then
                    is_inside = False
                Else
                    If is_inside = True And Mid(DATALINES, zzz, 1) = "," Then
                    Else
                        valid_datalines = valid_datalines & Mid(DATALINES, zzz, 1)
                    End If
                End If
            Else
                If Mid(DATALINES, zzz, 1) = "'" And is_inside = True Then
                    is_inside = False
                Else
                    If is_inside = True And Mid(DATALINES, zzz, 1) = "," Then
                    Else
                        valid_datalines = valid_datalines & Mid(DATALINES, zzz, 1)
                    End If
                End If
            End If
        Next
        C_CHAR = Split(COLUMNLINES, ",")
        D_CHAR = Split(valid_datalines, ",")
        T_COL = UBound(C_CHAR)
        D_COL = UBound(D_CHAR)


        For i = 0 To UBound(C_CHAR)
            STRHTML = STRHTML & "<TR><TD class='T1'> " & C_CHAR(i) & "&nbsp; </TD>"
            STRHTML = STRHTML & "<TD class='T2'> " & Replace(D_CHAR(i), "'", "") & "&nbsp;</TD></TR>"
        Next
    Else
        STR = Mid(STR, 1, InStr(1, UCase(STR), "WHERE"))

        'Dim TRY As String
        'TRY = GetNewSTR(STR)
        STR = Replace(STR, N2Str2Null(""), "''")
        C_CHAR = Split(STR, "'")

        For i = 0 To UBound(C_CHAR) - 1
            C_CHAR(i) = Replace(UCase(C_CHAR(i)), "=", "")
            If i = 0 Then                             ' THIS IS ALWAYS FIRST LINE IN UPDATE STATEMENTS
                C_CHAR(i) = Replace(UCase(C_CHAR(i)), "UPDATE", "")
                C_CHAR(i) = Replace(UCase(C_CHAR(i)), "SET", "")

                FIRSTPOSITION = InStr(1, LTrim(C_CHAR(i)), " ")
                C_CHAR(i) = Replace(UCase(C_CHAR(i)), Mid(C_CHAR(i), 1, FIRSTPOSITION), "")
                STRHTML = STRHTML & "<TR><td CLASS='T1'>" & Replace(C_CHAR(i), ",", "</TD>")
            Else
                If i Mod 2 = 0 Then
                    STRHTML = STRHTML & "<TR>" & Replace(C_CHAR(i), ",", "</TD><TD CLASS='T1'>")
                Else
                    STRHTML = STRHTML & "<TD CLASS='T2'>" & Replace(C_CHAR(i), "'", "") & "</td></TR>"
                End If
            End If
        Next
    End If

    Open App.Path & "\a.HTML" For Output As #1
    Print #1, "<html>"
    Print #1, "<style>"
    'Print #1, "Div {BORDER:1PX SOLID ORANGE;color:navy;text-align:CENTER;font:700 11px Arial;BACKGROUND-COLOR:#ffffa6;BORDER-BOTTOM:3PX DOUBLE ORANGE;}"
    Print #1, "Body {font:10pt Arial;background-color:buttonface;border:0px; margin:0;}"
    Print #1, ".THD1{color:blue;font:9pt Arial;background:buttonface;text-transform:uppercase; width:20%;font-weight:700;}"
    Print #1, ".THD2{font:9pt Arial;background:white;text-transform:uppercase; width:100PX;BORDER:1PX SOLID BLACK;}"
    Print #1, ".T1{color:blue;font:9pt Arial;background:buttonface;text-transform:uppercase; width:20%;font-weight:700;}"
    Print #1, ".T2{font:9pt Arial;background:white;text-transform:uppercase; width:80%;BORDER:1PX SOLID BLACK;}"
    Print #1, "td{font:9pt Arial;background:white;text-transform:uppercase;}"
    Print #1, "</style>"
    Print #1, "<body onkeydown='keyascii=0;' scroll=auto>"
    Print #1, "<Table WIDTH=100%>"
    Print #1, "<tr><Td CLASS='THD1'>Application Name:</Td><Td CLASS='THD2'>" & MODULENAME & "</Td><Td CLASS='THD1'>DESCRIPTION:</Td><Td CLASS='THD2'>" & Null2String(RSMEMO!Description) & "</Td><tr>"
    Print #1, "<tr><Td CLASS='THD1'>Date:</Td><Td CLASS='THD2'>" & Date_Value & "</Td><Td CLASS='THD1'>Time:</Td><Td CLASS='THD2'>" & time_value & "</Td><tr>"
    Print #1, "<tr><Td CLASS='THD1'>USERS NAME:</Td><Td CLASS='THD2'>" & Null2String(RSMEMO!UserName) & "</Td><Td CLASS='THD1'>USERS ACTION:</Td><Td CLASS='THD2'>" & Null2String(RSMEMO!USERACTION) & "</Td><tr>"

    Print #1, "</Table>"
    Print #1, "<hr/>"
    Print #1, "<table WIDTH=100%>"; Replace(STRHTML, "NULL", "") & "</table>"
    Print #1, "</body>"
    Close #1

    wbName.Navigate2 App.Path & "\a.html"

    DoEvents
    Exit Function
    
ERROR_CODE:
    SHOW_AUDITDETAILS = False

End Function

Private Sub cmdPrint_Click()
    If wbName.ReadyState = READYSTATE_COMPLETE Then
        wbName.ExecWB OLECMDID_PRINTPREVIEW2, OLECMDEXECOPT_DODEFAULT
    End If
End Sub

Private Sub Command1_Click()
    Picture1.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Picture1.Visible = True Then
            Picture1.Visible = False
            Exit Sub
        Else
            Unload Me
            Exit Sub
        End If
    End If
    If KeyCode = vbKeyF1 And Len(Text1) > 0 Then
        Picture1.Visible = True
        Exit Sub
    End If

    If KeyCode = vbKeyF3 Then
        cmdPrint_Click
        Exit Sub
    End If
End Sub

Private Sub Form_Load()

    CenterMe frmMain, Me, 1
End Sub


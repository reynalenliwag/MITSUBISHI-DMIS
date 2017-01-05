VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMS_PrintDailyLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Daily Log"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4185
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1200
      TabIndex        =   5
      Top             =   2220
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SET PATH AND FILENAME"
      Height          =   555
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   1665
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   1830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   39780353
      CurrentDate     =   39665
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   2280
      TabIndex        =   1
      Top             =   420
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   39780353
      CurrentDate     =   39665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PATH \ FILENAME"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3885
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   4155
      _Version        =   655364
      _ExtentX        =   7329
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "DATE RANGE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Alignment       =   1
      GradientColorLight=   8421504
      GradientColorDark=   4210752
      ForeColor       =   16777215
   End
End
Attribute VB_Name = "frmHRMS_PrintDailyLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Command2.Enabled = False
    CommonDialog1.Filename = "DTR.CSV"

    CommonDialog1.InitDir = "C:\"
    CommonDialog1.ShowSave
    Label1.Caption = CommonDialog1.Filename

    Command2.Enabled = True
End Sub

Function GetName(XEMPNO As String) As String
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT LASTNAME, FIRSTNAME FROM HRMS_EMPINFO WHERE EMPNO = '" & XEMPNO & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetName = Null2String(RSTMP!lastname & ", " & RSTMP!FIRSTNAME)
    Else
        GetName = Null2String("")
    End If
    Set RSTMP = Nothing
End Function

Private Sub Command2_Click()
    On Error GoTo adder:
    Screen.MousePointer = 11
    Dim INAM                                                          As String
    Dim OUTAM                                                         As String
    Dim INPM                                                          As String
    Dim OUTPM                                                         As String
    

    Open Label1.Caption For Output As #1
    Print #1, "EMPNO" & ",  " & "FULL NAME" & ", " & "DATE LOG" & ", " & "LOG1" & ", " & "LOG2" & ", " & "LOG3" & ", " & "LOG4"
    Dim rsAttend                                                      As ADODB.Recordset
    Dim RSHRMS As New ADODB.Recordset
    Set RSHRMS = gconDMIS.Execute("SELECT EMPNO, LASTNAME, FIRSTNAME FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A' ORDER BY LASTNAME")
    If Not (RSHRMS.BOF And RSHRMS.EOF) Then
        Do While Not RSHRMS.EOF
            Print #1, Null2String(RSHRMS!EMPNO) & ", " & Null2String(RSHRMS!lastname & " " & RSHRMS!FIRSTNAME)
            Set rsAttend = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE DATETODAY BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND EMPNO = '" & Null2String(RSHRMS!EMPNO) & "' ORDER BY DATETODAY")
            If Not rsAttend.EOF And Not rsAttend.BOF Then
                rsAttend.MoveFirst
                While Not rsAttend.EOF
                    'If Null2String(rsAttend!EMPNO) = "55062" Then Stop
                    If Null2String(rsAttend!INAM) <> "" Then INAM = Format(rsAttend!INAM, "short time")
                    If Null2String(rsAttend!INAM) = "" Then INAM = ""
                    
                    If Null2String(rsAttend!OUTAM) <> "" Then OUTAM = Format(rsAttend!OUTAM, "short time")
                    If Null2String(rsAttend!OUTAM) = "" Then OUTAM = ""
                    
                    If Null2String(rsAttend!INPM) <> "" Then INPM = Format(rsAttend!INPM, "short time")
                    If Null2String(rsAttend!INPM) = "" Then INPM = ""
                    
                    If Null2String(rsAttend!OUTPM) <> "" Then OUTPM = Format(rsAttend!OUTPM, "short time")
                    If Null2String(rsAttend!OUTPM) = "" Then OUTPM = ""
        
                    Print #1, "" & ", " & ", " & DateValue(rsAttend!datetoday) & ", " & Null2String(INAM) & ", " & Null2String(OUTAM) & ", " & Null2String(INPM) & ", " & Null2String(OUTPM); ""
                    rsAttend.MoveNext
                Wend
            End If
            Set rsAttend = Nothing
            RSHRMS.MoveNext
        Loop
    End If
    Close #1
    MsgBox "File has succesfully created"
    Screen.MousePointer = 0
    
    Exit Sub
adder:
    If Err.NUMBER = 70 Then
        MsgBox "Please Close Your File", vbExclamation
        Err.Clear
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    
    
    DTPicker1.Value = firstDay(Date)
    DTPicker2.Value = lastDay(Date)
    Screen.MousePointer = 0
End Sub


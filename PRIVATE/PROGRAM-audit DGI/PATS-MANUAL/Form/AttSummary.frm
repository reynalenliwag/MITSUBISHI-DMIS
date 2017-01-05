VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAttSummary 
   BackColor       =   &H00D7C6B5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance Summary"
   ClientHeight    =   3495
   ClientLeft      =   3915
   ClientTop       =   3330
   ClientWidth     =   3750
   ForeColor       =   &H00D7C6B5&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3495
   ScaleWidth      =   3750
   Begin VB.TextBox TxtHolidays 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2010
      TabIndex        =   7
      Top             =   1950
      Width           =   1635
   End
   Begin VB.Frame FrmProgBar 
      BackColor       =   &H00D7C6B5&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   3615
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   210
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Lbl1ProgBar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   225
         Width           =   2535
      End
      Begin VB.Label Lbl2ProgBar 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3060
         TabIndex        =   14
         Top             =   225
         Width           =   420
      End
   End
   Begin Crystal.CrystalReport rptAttSummary 
      Left            =   -30
      Top             =   3630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Attendance Summary"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.ComboBox cboMonth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   2835
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   1920
      TabIndex        =   11
      Top             =   2970
      Width           =   1710
   End
   Begin VB.CommandButton CmdShow 
      Caption         =   "Show Summary"
      Height          =   405
      Left            =   135
      TabIndex        =   10
      Top             =   2970
      Width           =   1710
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D7C6B5&
      Caption         =   "Print Selection"
      Height          =   1410
      Left            =   60
      TabIndex        =   9
      Top             =   465
      Width           =   3615
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   1440
         ScaleHeight     =   1035
         ScaleWidth      =   2025
         TabIndex        =   18
         Top             =   180
         Width           =   2085
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00D7C6B5&
         Caption         =   "16 - 31"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   5
         Top             =   660
         Width           =   870
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00D7C6B5&
         Caption         =   "01 - 31"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   6
         Top             =   1005
         Width           =   870
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00D7C6B5&
         Caption         =   "01 - 15"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   4
         Top             =   315
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee  Selection"
      Height          =   240
      Left            =   600
      TabIndex        =   8
      Top             =   3660
      Width           =   1755
      Begin VB.OptionButton Option1 
         Caption         =   "Allowance Base"
         Height          =   345
         Index           =   2
         Left            =   315
         TabIndex        =   3
         Top             =   930
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Contractual"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   2
         Top             =   660
         Width           =   1320
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Employed"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   1
         Top             =   315
         Width           =   1095
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Month of"
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Official Holidays:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   1995
      Width           =   1680
   End
End
Attribute VB_Name = "frmAttSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DaysSelect                                      As Integer
Public EmploymentStatus                                As String
Public Filter                                          As String

Private Sub Parse()
    Dim H                                              As String
    Dim H1, H2, H3, I, v                               As Integer

    H1 = 0: H2 = 0: H3 = 0
    H = Trim(TxtHolidays)
    If H <> "" Then
        v = InStr(H, ",")
        If v = 0 Then
            H1 = Val(H)
        Else
            H1 = Val(Left(H, v - 1))
            I = InStr(v + 1, H, ",")
            If I = 0 Then
                H2 = Val(Mid(H, v + 1))
            Else
                H2 = Val(Mid(H, v + 1, I - v - 1))
                H3 = Val(Mid(H, I + 1))
            End If
        End If
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
    frmLOGIN.TxtEmpNumber.SetFocus
End Sub

Private Sub CmdShow_Click()
    Dim C                                              As Integer
    Dim D                                              As Integer
    Dim K                                              As Integer
    Dim W                                              As Integer
    Dim U1                                             As Integer
    Dim U2                                             As Integer
    Dim U3                                             As Integer
    Dim U4                                             As Integer
    Dim TU                                             As Integer
    Dim ST                                             As Integer
    Dim EN                                             As Integer
    Dim I                                              As Integer
    Dim D1                                             As Integer
    Dim D2                                             As Integer
    Dim DAYSABSENT                                     As Single
    Dim ACTUALWORKDAYS                                 As Single
    Dim UNDERTIME                                      As Integer
    Dim UTHRS                                          As Integer
    Dim UTMINS                                         As Integer
    Dim LATE                                           As Integer
    Dim LATEHRS                                        As Integer
    Dim LATEMINS                                       As Integer
    Dim OVERTIME                                       As Integer
    Dim OTHRS                                          As Integer
    Dim OTMINS                                         As Integer

    Dim DAYLIST                                        As String
    Dim PE                                             As String
    ReDim da(31, 5) As String

    Dim TDATE                                          As Date
    Dim THEMONTH                                       As Integer
    Dim THEYEAR                                        As Integer
    TDATE = Date

    If cboMonth.Text = Date2Month(OneMonth(Date, -2)) Then
        TDATE = OneMonth(Date, -2)
    ElseIf cboMonth.Text = Date2Month(OneMonth(Date, -1)) Then
        TDATE = OneMonth(Date, -1)
    ElseIf cboMonth.Text = Date2Month(Date) Then
        TDATE = Date
    ElseIf cboMonth.Text = Date2Month(OneMonth(Date, 1)) Then
        TDATE = OneMonth(Date, 1)
    End If

    THEMONTH = Month(TDATE)
    THEYEAR = Year(TDATE)
    'Parse the holidays
    Dim H1(31)                                         As String
    Dim H                                              As String
    Dim Dy                                             As String
    Dim C1(10)                                         As Integer

    'H = Trim(TxtHolidays): k = 1
    'If H <> "" Then
    '   'find location of commas
    '   For i = 1 To Len(H)
    '       If Mid(H, i, 1) = "," Then
    '          k = k + 1: C1(k) = i
    '       End If
    '   Next i
    '   C1(k + 1) = Len(H) + 1
    '   'Place date markers ("0" not a holiday, "1" holiday, "A" holiday AM only, "P" holiday PM only)
    '   For i = 1 To k
    '       Dy = Mid(H, C1(i) + 1, C1(i + 1) - (C1(i) + 1))
    '       If Val(Dy) <= 31 Then
    '          If UCase(Right(Dy, 1)) = "A" Then
    '             H1(Val(Dy)) = "A"
    '          ElseIf UCase(Right(Dy, 1)) = "P" Then
    '             H1(Val(Dy)) = "P"
    '          Else
    '             H1(Val(Dy)) = "1"
    '          End If
    '       End If
    '   Next i
    'End If

    '**********************************************
    GCONDMIS.Execute "Delete from HRMS_TempSummary"
    FrmProgBar.Visible = True
    Me.Refresh
    Lbl1ProgBar.Caption = "Creating Summary,  Please wait..."
    'rsEmpInfo.MoveFirst
    I = 0
    Set rsEMPINFO = New ADODB.Recordset
    Set rsEMPINFO = GCONDMIS.Execute("Select * from HRMS_EmpInfo order by EmpNo asc")
    Do Until rsEMPINFO.EOF
        I = I + 1
        ProgressBar1.Value = (I / rsEMPINFO.RecordCount) * 100
        Lbl2ProgBar.Caption = Int(ProgressBar1.Value) & "%"
        Me.Refresh
        'If Null2String(rsEmpInfo!divcode) = thedivcode Then
        If Null2String(rsEMPINFO!ACTIVEINACTIVE) = "A" Then
            'If Null2String(rsEmpInfo!EMPLEVEL) = EmploymentStatus Then
            Set rsCard = New ADODB.Recordset
            rsCard.Open "Select * from HRMS_Attend Where right(EmpNo,4) = " & N2Str2Null(rsEMPINFO!empno) & " AND Month(DateToday) = " & Month(TDATE) & " AND Year(DateToday) = " & Year(TDATE), GCONDMIS
            If Not rsCard.BOF And Not rsCard.EOF Then
                rsCard.MoveFirst
                Do Until rsCard.EOF
                    D = Day(rsCard!DateToday)
                    da(D, 1) = rsCard!DateToday
                    da(D, 2) = Format(Null2String(rsCard!InAm), "HH:MM:SS AM/PM")
                    da(D, 3) = Format(Null2String(rsCard!OutAm), "HH:MM:SS AM/PM")
                    da(D, 4) = Format(Null2String(rsCard!InPm), "HH:MM:SS AM/PM")
                    da(D, 5) = Format(Null2String(rsCard!OutPM), "HH:MM:SS AM/PM")

                    If D > C Then C = D
                    rsCard.MoveNext
                Loop
                For K = 1 To C
                    If da(K, 1) = "" Then
                        da(K, 1) = CDate(Month(TDATE) & "/" & Str(K) & "/" & Right(Year(TDATE), 2))
                        da(K, 2) = ""
                        da(K, 3) = ""
                        da(K, 4) = ""
                        da(K, 5) = ""
                    End If
                Next K
                ACTUALWORKDAYS = 0: DAYSABSENT = 0: UNDERTIME = 0: LATE = 0: OVERTIME = 0: DAYLIST = ""
                If DaysSelect = 0 Then
                    'st = 1: en = 15: Pe = "01-15"
                    If C > 15 Then
                        C = 15
                    End If
                    ST = 1: EN = C: PE = "01-15"
                ElseIf DaysSelect = 1 Then
                    ST = 16: EN = C: PE = "16-31"
                Else
                    ST = 1: EN = C: PE = "01-31"
                End If
                For K = ST To EN
                    W = Weekday(da(K, 1))
                    'Dow = Mid(DaysOfWeek, (w - 1) * 3 + 1, 3)
                    'If w = 2 Or w = 3 Or w = 4 Or w = 5 Or w = 6 Then (Until Friday Only)
                    If W = 2 Or W = 3 Or W = 4 Or W = 5 Or W = 6 Or W = 7 Then
                        If H1(K) <> "1" Then

                            U1 = 0: U2 = 0: U3 = 0: U4 = 0: TU = 0
                            If da(K, 2) = "" And da(K, 3) = "" And da(K, 4) = "" And da(K, 5) = "" Then
                                If H1(K) = "" Then
                                    DAYSABSENT = DAYSABSENT + 1
                                    DAYLIST = DAYLIST + Str(K)
                                ElseIf H1(K) = "A" Then
                                    DAYSABSENT = DAYSABSENT + 0.5
                                    DAYLIST = DAYLIST + Str(K) + "P"
                                ElseIf H1(K) = "P" Then
                                    DAYSABSENT = DAYSABSENT + 0.5
                                    DAYLIST = DAYLIST + Str(K) + "A"
                                End If

                            ElseIf da(K, 2) = "" Or da(K, 3) = "" Or da(K, 4) = "" Or da(K, 5) = "" Then
                                If H1(K) = "" Or H1(K) = "P" Then
                                    D1 = 0
                                    If da(K, 2) = "" Or da(K, 3) = "" Then
                                        DAYSABSENT = DAYSABSENT + 0.5
                                        D1 = K
                                    End If
                                    If da(K, 2) <> "" And da(K, 3) <> "" Then
                                        If da(K, 2) > #8:00:00 AM# Then
                                            U1 = DateDiff("n", #8:00:00 AM#, da(K, 2))
                                        End If
                                        If da(K, 3) < #11:59:59 AM# Then
                                            U2 = DateDiff("n", da(K, 3), #11:59:00 AM#) + 1
                                        End If
                                    End If
                                End If
                                If H1(K) = "" Or H1(K) = "A" Then
                                    D2 = 0
                                    If da(K, 4) = "" Or da(K, 5) = "" Then
                                        DAYSABSENT = DAYSABSENT + 0.5
                                        D2 = K
                                    End If
                                    If da(K, 4) <> "" And da(K, 5) <> "" Then
                                        If da(K, 4) > #1:00:00 PM# Then
                                            U3 = DateDiff("n", #1:00:00 PM#, da(K, 4))
                                        End If
                                        If da(K, 5) < #5:00:00 PM# Then
                                            U4 = DateDiff("n", da(K, 5), #5:00:00 PM#)
                                        End If
                                    End If
                                End If
                                If D1 = D2 Then
                                    DAYLIST = DAYLIST + Str(K)
                                ElseIf D1 > 0 Then
                                    DAYLIST = DAYLIST + Str(K) + "A"
                                ElseIf D2 > 0 Then
                                    DAYLIST = DAYLIST + Str(K) + "P"
                                End If
                            Else
                                If H1(K) = "" Or H1(K) = "P" Then
                                    If da(K, 2) > #8:00:00 AM# Then
                                        U1 = DateDiff("n", #8:00:00 AM#, da(K, 2))
                                    End If
                                    If da(K, 3) < #11:59:59 AM# Then
                                        U2 = DateDiff("n", da(K, 3), #11:59:00 AM#) + 1
                                    End If
                                End If
                                If H1(K) = "" Or H1(K) = "A" Then
                                    If da(K, 4) > #1:00:00 PM# Then
                                        U3 = DateDiff("n", #1:00:00 PM#, da(K, 4))
                                    End If
                                    If da(K, 5) < #5:00:00 PM# Then
                                        U4 = DateDiff("n", da(K, 5), #5:00:00 PM#)
                                    End If
                                End If
                            End If
                            TU = U1 + U2 + U3 + U4
                            UNDERTIME = UNDERTIME + TU
                            LATE = LATE + TU
                            OVERTIME = OVERTIME + TU
                        End If
                    End If
                Next K
                'If DaysAbsent > 0 Or UnderTime > 0 Then
                UTHRS = Int(UNDERTIME / 60)
                UTMINS = UNDERTIME Mod 60
                LATEHRS = Int(LATE / 60)
                LATEMINS = LATE Mod 60
                OTHRS = Int(OVERTIME / 60)
                OTMINS = OVERTIME Mod 60
                GCONDMIS.Execute "Insert into HRMS_TempSummary (EmpNo, ActualWorkDays, DaysAbsent, UThrs, UTmins, LateHrs, LateMins, OTHrs, OTMins, Days, Period, EMPLEVEL) " & _
                                 "Values (" & N2Str2Null(rsEMPINFO!empno) & "," & ACTUALWORKDAYS & ", " & DAYSABSENT & "," & UTHRS & "," & UTMINS & "," & LATEHRS & "," & LATEMINS & "," & OTHRS & "," & OTMINS & ",'" & DAYLIST & "','" & PE & "','" & EmploymentStatus & "')"
                'End If
            End If
            'End If
        End If
        'End If
        rsEMPINFO.MoveNext
        Erase da
        ReDim da(31, 5) As String
    Loop

    FrmProgBar.Visible = False
    'frmLOGIN.TxtEmpNumber.SetFocus
    frmLOGIN.Refresh
    'Filter = "{Master.divcode} = '" & thedivcode & "' AND {Master.EMPLEVEL} = '" & EmploymentStatus & "'"

    'rptAttSummary.WindowTitle = "PATS"
    'rptAttSummary.ReportFileName = HRMS_REPORT_PATH & "AttSum.Rpt"
    'rptAttSummary.SelectionFormula = Filter
    'rptAttSummary.Action = 1
    'rptAttSummary.Password = "bioskeyw7ux2"
    'SendKeys ("bioskeyw7ux2") & Chr$(13)
    PrintSQLReport rptAttSummary, HRMS_REPORT_PATH & "ATTSum.Rpt", Filter, DMIS_REPORT_Connection, 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    fillcbomonth
    cboMonth = Date2Month(Date)
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        EmploymentStatus = "E"
    ElseIf Index = 1 Then
        EmploymentStatus = "C"
    Else
        EmploymentStatus = "A"
    End If
End Sub

Private Sub Option2_Click(Index As Integer)
    DaysSelect = Index
    CmdShow.SetFocus
End Sub

Sub fillcbomonth()
    Dim X                                              As Integer
    Dim Thedate                                        As Date
    Thedate = OneMonth(Date, -3)
    For X = 1 To 4
        Thedate = OneMonth(Thedate, 1)
        cboMonth.AddItem Date2Month(CDate(Thedate))
    Next
End Sub

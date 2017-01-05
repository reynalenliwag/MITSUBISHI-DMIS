VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmHRMSUpDateAttendance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Attendance"
   ClientHeight    =   8250
   ClientLeft      =   1545
   ClientTop       =   3180
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "UpdateAttendance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8250
   ScaleWidth      =   7575
   Begin FlexCell.Grid Grid1 
      Height          =   4935
      Left            =   0
      TabIndex        =   19
      Top             =   600
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8705
      Appearance      =   0
      BackColor2      =   12907725
      BackColorBkg    =   -2147483645
      Cols            =   5
      DefaultFontName =   "Courier New"
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      GridColor       =   12632256
      Rows            =   30
   End
   Begin VB.CheckBox chkOTRetain 
      Caption         =   "Retain OT Data"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6060
      TabIndex        =   17
      Top             =   5670
      Value           =   1  'Checked
      Width           =   1905
   End
   Begin VB.CheckBox chkConfidential 
      Caption         =   "Process Confidential Employees"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   150
      TabIndex        =   16
      Top             =   5670
      Width           =   3675
   End
   Begin VB.CheckBox chkProbReg 
      Caption         =   "Process for Probationary/Regular Employees"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   150
      TabIndex        =   15
      Top             =   5910
      Width           =   3675
   End
   Begin VB.CheckBox chkAllowanceBase 
      Caption         =   "Process for Allowance Base Employees"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   150
      TabIndex        =   14
      Top             =   6450
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.CheckBox chkContractual 
      Caption         =   "Process for Contractual Employees"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   150
      TabIndex        =   13
      Top             =   6180
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   6300
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   45
      Width           =   1155
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   3630
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   45
      Width           =   2115
   End
   Begin VB.ComboBox cboQuensina 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   45
      Width           =   3045
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   795
      Left            =   6540
      MouseIcon       =   "UpdateAttendance.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "UpdateAttendance.frx":0594
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancel"
      Top             =   7410
      Width           =   945
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6540
      Picture         =   "UpdateAttendance.frx":1616
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7410
      Width           =   945
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "Go"
      Height          =   795
      Left            =   5670
      MouseIcon       =   "UpdateAttendance.frx":1A58
      MousePointer    =   99  'Custom
      Picture         =   "UpdateAttendance.frx":1BAA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Update Attendance Now"
      Top             =   7410
      Width           =   885
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   90
      ScaleHeight     =   1245
      ScaleWidth      =   7425
      TabIndex        =   3
      Top             =   6720
      Width           =   7425
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   90
         ScaleHeight     =   345
         ScaleWidth      =   5055
         TabIndex        =   4
         Top             =   720
         Width           =   5055
         Begin VB.Label labName 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   4935
         End
      End
      Begin wizProgBar.Prg gauProgress 
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   270
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   556
         Picture         =   "UpdateAttendance.frx":2C2C
         BackColor       =   14215660
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "UpdateAttendance.frx":2C48
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   30
         ScaleHeight     =   495
         ScaleWidth      =   5235
         TabIndex        =   7
         Top             =   660
         Width           =   5235
         Begin wizButton.cmd cmd1 
            Height          =   465
            Left            =   30
            TabIndex        =   8
            Top             =   0
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   820
            TX              =   "cmd1"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "UpdateAttendance.frx":2C64
         End
      End
      Begin VB.Label lblPercent 
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
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
         Height          =   225
         Left            =   45
         TabIndex        =   9
         Top             =   0
         Width           =   5595
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   705
      Left            =   -60
      TabIndex        =   18
      Top             =   -30
      Width           =   8025
      _Version        =   655364
      _ExtentX        =   14155
      _ExtentY        =   1244
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmHRMSUpDateAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EMP_RATE_MIN                                       As Double
Dim EMP_RATE_HRS                                       As Double
Dim EMP_RATE_DAY                                       As Double
Dim PROCESS_OPTION                                     As String
Dim GENFROM                                            As String
Dim GENTO                                              As String
Dim SHIFTFROM1                                         As String
Dim SHIFTFROM2                                         As String
Dim SHIFTTO1                                           As String
Dim SHIFTTO2                                           As String
Dim GRACE_PERIOD                                       As Integer
Dim LOG_OPTION                                         As Integer
Dim LOGINAM                                            As String
Dim LOGOUTLUNCH                                        As String
Dim LOGINLUNCH                                         As String
Dim LOGOUTPM                                           As String
Dim START_OF_OT                                        As String
Dim START_OF_ND                                        As String
Dim OT_COMPUTE                                         As Integer

Function CheckWorkDay(DAYNUMBER As Integer) As Boolean
    CheckWorkDay = False
    If DAYNUMBER = 1 Then
        CheckWorkDay = False
    ElseIf DAYNUMBER = 2 Then
        CheckWorkDay = True
    ElseIf DAYNUMBER = 3 Then
        CheckWorkDay = True
    ElseIf DAYNUMBER = 4 Then
        CheckWorkDay = True
    ElseIf DAYNUMBER = 5 Then
        CheckWorkDay = True
    ElseIf DAYNUMBER = 6 Then
        CheckWorkDay = True
    ElseIf DAYNUMBER = 7 Then
        If COMPANY_CODE = "HARI" Then
            CheckWorkDay = False
        Else
            CheckWorkDay = True
        End If
    End If
End Function

Function ComputeLateAndUndertime(TIME1 As String, TIME2 As String, TIME3 As String, TIME4 As String, SHIFTINAM As String, SHIFTLUNCHOUT As String, SHIFTLUNCHIN As String, SHIFTOUTPM As String, GRACE_PERIOD As Integer, LOG_OPTION As Integer)
    Dim UNDERTIMEVAR                                   As Double
    UNDERTIMEVAR = 0
    ComputeLateAndUndertime = 0
    If LOG_OPTION = 0 Then
        If TimeValue(TIME1) > TimeValue(DateAdd("n", GRACE_PERIOD, SHIFTINAM)) Then
            UNDERTIMEVAR = DateDiff("n", TimeValue(SHIFTINAM), TimeValue(TIME1))
            ComputeLateAndUndertime = ComputeLateAndUndertime + Round(UNDERTIMEVAR, 2)
        End If
        If TimeValue(TIME1) > TimeValue(DateAdd("n", GRACE_PERIOD, SHIFTLUNCHIN)) Then
            UNDERTIMEVAR = DateDiff("n", TimeValue(SHIFTLUNCHIN), TimeValue(TIME1))
            ComputeLateAndUndertime = ComputeLateAndUndertime + Round(UNDERTIMEVAR, 2)
        End If
        If TimeValue(TIME2) < TimeValue(SHIFTLUNCHOUT) Then
            UNDERTIMEVAR = DateDiff("n", TimeValue(TIME2), TimeValue(SHIFTLUNCHOUT))
            ComputeLateAndUndertime = ComputeLateAndUndertime + Round(UNDERTIMEVAR, 2)
        End If
        If TimeValue(TIME2) < TimeValue(SHIFTOUTPM) Then
            UNDERTIMEVAR = DateDiff("n", TimeValue(TIME2), TimeValue(SHIFTOUTPM))
            ComputeLateAndUndertime = ComputeLateAndUndertime + Round(UNDERTIMEVAR, 2)
        End If
    ElseIf LOG_OPTION = 1 Then

    End If
    ComputeLateAndUndertime = Round(ComputeLateAndUndertime, 2)
End Function

Function ComputeAbsence(TIME1 As String, TIME2 As String, TIME3 As String, TIME4 As String, SHIFTINAM As String, SHIFTLUNCHOUT As String, SHIFTLUNCHIN As String, SHIFTOUTPM As String, GRACE_PERIOD As Integer, LOG_OPTION As Integer, Optional XDATE_TODAY As String, Optional NOWDATE As String)
    Dim ABSENTVAR                                      As Double
    Dim I                                              As Integer
    ABSENTVAR = 0
    ComputeAbsence = 0
    I = 1
    'ADD BY MJP : FOR CHECKING PURPOSE
        'If Not XDATE_TODAY = "" Then Debug.Print DateValue(XDATE_TODAY)
        
        'FIND OUT SOME DATA IN HRMS_ATTEND HAVE AN DATE IN INAM, OUTPM BUT SOME DONT HAVE
        'SO ITS HARD TO CHECK IF THE EMPLOYEE OUT AT THE NEXT DAY
    '-----------------------
    
    If LOG_OPTION = 0 Then
        If TIME1 = "" Or TIME2 = "" Then
            
            Dim leave As New ADODB.Recordset
            Dim reqdesc As String

            Set leave = gconDMIS.Execute("SELECT * FROM HRMS_REQUESTLEAVE_OT where '" & NOWDATE & "' between dte_from  and  dte_to  and status = 'A' and empno = '" & Grid1.Cell(I, 1).Text & "'")

            If Not (leave.BOF And leave.EOF) Then
                 reqdesc = Null2String(leave!reqdesc)
            Else
                ABSENTVAR = 1
                GoTo ABSENTLABEL
            End If
            
        ElseIf TIME1 <> "" And TIME2 <> "" Then
            If (TimeValue(TIME1) > TimeValue(SHIFTLUNCHOUT)) And (TimeValue(TIME2) < TimeValue(SHIFTLUNCHIN)) Then
                ABSENTVAR = 1
                GoTo ABSENTLABEL
            ElseIf TimeValue(TIME1) > TimeValue(SHIFTLUNCHOUT) Then
                ABSENTVAR = 0.5
            ElseIf TimeValue(TIME2) < TimeValue(SHIFTLUNCHIN) Then
                ABSENTVAR = 0.5
            End If
        End If
    ElseIf LOG_OPTION = 1 Then
        If TIME1 = "" And TIME2 = "" And TIME3 = "" And TIME4 = "" Then
            ABSENTVAR = 1
            GoTo ABSENTLABEL
        ElseIf TIME1 <> "" And TIME2 = "" And TIME3 = "" And TIME4 = "" Then
            ABSENTVAR = 1
            GoTo ABSENTLABEL
        Else
            If TIME1 <> "" And TIME2 = "" And TIME3 <> "" And TIME4 = "" Then
                If TIME3 < SHIFTLUNCHIN Then
                    ABSENTVAR = 0.5
                End If
            ElseIf TIME1 <> "" And TIME2 = "" And TIME3 <> "" And TIME4 <> "" Then
                If TIME4 < SHIFTLUNCHIN Then
                    ABSENTVAR = 0.5
                End If
            ElseIf TIME1 <> "" And TIME2 <> "" And TIME3 <> "" And TIME4 <> "" Then
                If TIME4 < SHIFTLUNCHIN Then
                    ABSENTVAR = 0.5
                End If
            End If
        End If
    End If
ABSENTLABEL:
    ComputeAbsence = Round(ABSENTVAR, 2)
End Function

Function ComputeOvertime(TIME1 As String, TIME2 As String, TIME3 As String, TIME4 As String, SHIFTINAM As String, SHIFTLUNCHOUT As String, SHIFTLUNCHIN As String, SHIFTOUTPM As String, GRACE_PERIOD As Integer, LOG_OPTION As Integer, CODE As String, OTSTART As String, NDSTART As String, LEVEL As String, EMPNO As String, CUTOFF As String, OTDATE As String)
    Dim OVERTIMEVAR                                    As Double
    Dim OTHOUR                                         As Double
    Dim AMOUNT                                         As Double
    AMOUNT = 0
    OVERTIMEVAR = 0
    ComputeOvertime = 0
    OTHOUR = 0
    If (CODE = "003" Or CODE = "004" Or CODE = "005" Or CODE = "002") Then
        OVERTIMEVAR = DateDiff("n", TimeValue(SHIFTINAM), TimeValue(TIME2))
        If (((TimeValue(TIME2) > TimeValue(SHIFTLUNCHOUT)) And (TimeValue(TIME2) < TimeValue(SHIFTOUTPM))) Or (TimeValue(TIME2) > TimeValue(SHIFTOUTPM))) Then
            OVERTIMEVAR = OVERTIMEVAR - 60
            If TimeValue(TIME2) > TimeValue(SHIFTOUTPM) Then
                OVERTIMEVAR = OVERTIMEVAR - DateDiff("n", TimeValue(SHIFTOUTPM), TimeValue(TIME2))
            End If
        End If
        OTHOUR = Convert_To_Hour(OVERTIMEVAR)
        AMOUNT = OTHOUR * GetRate(CODE) * EMP_RATE_HRS
        Call SaveOT(CODE, OTHOUR, Round(AMOUNT, 2), LEVEL, EMPNO, CUTOFF, OTDATE)
        AMOUNT = 0
        OVERTIMEVAR = 0
        OTHOUR = 0
    End If
    
    If (DateDiff("h", TIME1, TIME2) > 8) And (TimeValue(TIME2) > TimeValue(OTSTART)) Then
        If (TimeValue(TIME2) <= TimeValue(NDSTART)) And (TimeValue(TIME2) > TimeValue(OTSTART)) Then
            OVERTIMEVAR = DateDiff("n", TimeValue(OTSTART), TimeValue(TIME2))
            If OVERTIMEVAR > 0 Then
                OTHOUR = Convert_To_Hour(OVERTIMEVAR)
                AMOUNT = OTHOUR * GetRate(GetCodeOT(CODE)) * EMP_RATE_HRS
                Call SaveOT(GetCodeOT(CODE), OTHOUR, Round(AMOUNT, 2), LEVEL, EMPNO, CUTOFF, OTDATE)
                AMOUNT = 0
                OVERTIMEVAR = 0
                OTHOUR = 0
            End If
        End If
    End If

    ComputeOvertime = Round(OVERTIMEVAR, 2)
End Function

Function IsOnLeave(vDATE As Date) As Boolean
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_REQUESTLEAVE_OT WHERE STATUS = 'A' AND (DTE_FROM BETWEEN '" & vDATE & "' AND '" & vDATE & "' OR DTE_TO BETWEEN '" & vDATE & "' AND '" & vDATE & "')")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        IsOnLeave = True
    Else
        IsOnLeave = False
    End If
    Set RSTMP = Nothing
End Function

Function CheckSpecialType(MONTHNUMBER As Integer, DAYNUMBER As Integer) As Boolean
    CheckSpecialType = False
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_HOLIDAY_LIST WHERE MANTH = " & MONTHNUMBER & " AND DEYT = " & DAYNUMBER & " AND TYPE = '1'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckSpecialType = True
    Else
        CheckSpecialType = False
    End If
    Set RSTMP = Nothing
End Function

Function CheckHoliday(MONTHNUMBER As Integer, DAYNUMBER As Integer) As Boolean
    CheckHoliday = False
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_HOLIDAY_LIST WHERE MANTH = " & MONTHNUMBER & " AND DEYT = " & DAYNUMBER & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckHoliday = True
    Else
        CheckHoliday = False
    End If
    Set RSTMP = Nothing
End Function

Function GetCodeNightDifferential(CODE As String) As String
    GetCodeNightDifferential = ""
    If CODE = "001" Then
        GetCodeNightDifferential = "012"
    ElseIf CODE = "002" Then
        GetCodeNightDifferential = "013"
    ElseIf CODE = "003" Then
        GetCodeNightDifferential = "014"
    ElseIf CODE = "004" Then
        GetCodeNightDifferential = "015"
    ElseIf CODE = "005" Then
        GetCodeNightDifferential = "016"
    End If
End Function

Function GetCodeOT(CODE As String) As String
    GetCodeOT = ""
    If CODE = "001" Then
        GetCodeOT = "001"
    ElseIf CODE = "002" Then
        GetCodeOT = "020"
    ElseIf CODE = "003" Then
        GetCodeOT = "008"
    ElseIf CODE = "004" Then
        GetCodeOT = "009"
    ElseIf CODE = "005" Then
        GetCodeOT = "010"
    End If
End Function

Function GetRate(CODE As String) As Double
    GetRate = 0
    Dim rsTemp                                         As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT PAY_RATE FROM HRMS_OTCODES WHERE PAY_CODE = '" & CODE & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GetRate = N2Str2Zero(rsTemp!pay_rate)
    End If
    Set rsTemp = Nothing
End Function

Function Convert_To_Hour(NO_OF_MINUTES)
    Convert_To_Hour = 0
    Dim OTHOUR                                         As Double
    Dim OTMINUTE                                       As Double

    OTHOUR = NO_OF_MINUTES \ 60
    OTMINUTE = NO_OF_MINUTES Mod 60
    Convert_To_Hour = OTHOUR + (OTMINUTE / 60)
    Convert_To_Hour = Round(Convert_To_Hour, 2)
End Function

Sub InitGrid()
    With Grid1
        .Cols = 4
        .Column(0).Width = 50
        .Column(1).Width = 80
        .Column(2).Width = 250
        .Column(3).Width = 80
        .Cell(0, 0).Text = "L/N"
        .Cell(0, 1).Text = "EMPNO"
        .Cell(0, 2).Text = "EMPLOYEE NAME"
        .Cell(0, 3).Text = "OPTION"
        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(1).Alignment = cellCenterCenter
        .Column(3).Alignment = cellCenterCenter
    End With
End Sub

Sub FillGrid()
    PROCESS_OPTION = ""
    Dim matt                                           As Integer
    If cboQuensina.Text = "1st Cut-Off" Then
        matt = 1
    Else
        matt = 2
    End If

    If (cboQuensina.Text) = "1st Cut-Off" Then
        GENFROM = DateSerial(cboyear.Text, What_month(cboMOnth.Text), PAYROLLCODE_FROM1)
        GENTO = DateSerial(cboyear.Text, What_month(cboMOnth.Text), PAYROLLCODE_TO1)
        If PAYROLLCODE_FROM1 > PAYROLLCODE_TO1 Then
            GENFROM = DateSerial(NumericVal(cboyear.Text), What_month(cboMOnth.Text) - 1, PAYROLLCODE_FROM1)
            If What_month(cboMOnth.Text) = 1 Then
                GENFROM = DateSerial(NumericVal(cboyear.Text) - 1, 12, PAYROLLCODE_FROM1)
            End If
        End If
    Else
        GENFROM = DateSerial(cboyear.Text, What_month(cboMOnth.Text), PAYROLLCODE_FROM2)
        GENTO = DateSerial(cboyear.Text, What_month(cboMOnth.Text), PAYROLLCODE_TO2)
    End If

    If chkConfidential.Value = 1 Then
        PROCESS_OPTION = " EMPLEVEL = 'M'"
    End If
    If PROCESS_OPTION <> "" Then
        If chkProbReg.Value = 1 Then
            PROCESS_OPTION = PROCESS_OPTION & " OR EMPLEVEL = 'E'"
        End If
    Else
        PROCESS_OPTION = " EMPLEVEL = 'E'"
    End If
    If PROCESS_OPTION <> "" Then
        If chkContractual.Value = 1 Then
            PROCESS_OPTION = PROCESS_OPTION & " OR EMPLEVEL = 'C'"
        End If
    Else
        PROCESS_OPTION = " EMPLEVEL = 'C'"
    End If
    If PROCESS_OPTION <> "" Then
        If chkAllowanceBase.Value = 1 Then
            PROCESS_OPTION = PROCESS_OPTION & " OR EMPLEVEL = 'A'"
        End If
    Else
        PROCESS_OPTION = " EMPLEVEL = 'A'"
    End If
    If chkAllowanceBase.Value = 0 And chkConfidential.Value = 0 And chkContractual.Value = 0 And chkProbReg.Value = 0 Then
        Grid1.Rows = 1
        MsgBox "Please select an option to process...", vbInformation, "No Option to Process"
        cmdGO.Enabled = False
        Exit Sub
    End If
    PROCESS_OPTION = "(" & PROCESS_OPTION & ")"


'    Dim rsCheckPayroll                                 As ADODB.Recordset
'    Set rsCheckPayroll = New ADODB.Recordset
'    Set rsCheckPayroll = gconDMIS.Execute("SELECT * FROM HRMS_PAYROLL WHERE " & PROCESS_OPTION & " AND CUT_OFF = '" & matt & "' AND PAY_MONTH = " & What_month(cboMOnth.Text) & " AND PAY_YEAR = " & cboyear.Text & "")
'    If Not rsCheckPayroll.EOF And Not rsCheckPayroll.BOF Then
'        MsgBox "Payroll Already Generated! Please Clear Generated Payroll First.", vbInformation, "Not Allowed"
'        cmdGO.Enabled = False
'        Exit Sub
'    End If


    Grid1.Rows = 1
    Dim rsEMPINFO2                                     As ADODB.Recordset
    Set rsEMPINFO2 = New ADODB.Recordset
    rsEMPINFO2.Open "SELECT * FROM HRMS_EMPINFO WHERE (DATEHIRED <= '" & Format(GENTO, "SHORT DATE") & "') AND " & PROCESS_OPTION & " ORDER BY LASTNAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Grid1.Rows = 1
    If Not rsEMPINFO2.EOF And Not rsEMPINFO2.BOF Then
        rsEMPINFO2.MoveFirst
        While Not rsEMPINFO2.EOF
            Grid1.AddItem Null2String(rsEMPINFO2!EMPNO) & Chr(9) & Null2String(rsEMPINFO2!lastname) & ", " & Null2String(rsEMPINFO2!FIRSTNAME) & Chr(9) & "DELETE"
            rsEMPINFO2.MoveNext
        Wend
    End If
    Set rsEMPINFO2 = Nothing
    cmdGO.Enabled = True
End Sub

Sub GetEmployeeShift(CODE As String)
    SHIFTFROM1 = Format("08:00:00 AM", "HH:MM:SS AM/PM")
    SHIFTTO1 = Format("12:00:00 PM", "HH:MM:SS AM/PM")
    SHIFTFROM2 = Format("1:00:00 PM", "HH:MM:SS AM/PM")
    SHIFTTO2 = Format("5:00:00 PM", "HH:MM:SS AM/PM")

    Dim rsTemp                                         As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_TIME_SHIFT_CODE WHERE SHIFTCODE = '" & CODE & "'")
    If Not (rsTemp.BOF And rsTemp.EOF) Then
        LOG_OPTION = N2Str2Zero(rsTemp!PATS)
        GRACE_PERIOD = N2Str2Zero(rsTemp!GRACE_PERIOD)
        SHIFTFROM1 = Format(Null2String(rsTemp!FROM1), "HH:MM:SS AM/PM")
        SHIFTTO1 = Format(Null2String(rsTemp!LUNCHOUT), "HH:MM:SS AM/PM")
        SHIFTFROM2 = Format(Null2String(rsTemp!LUNCHIN), "HH:MM:SS AM/PM")
        SHIFTTO2 = Format(Null2String(rsTemp!TO1), "HH:MM:SS AM/PM")
        START_OF_OT = Format(Null2String(rsTemp!OTSTART), "HH:MM:SS AM/PM")
        
        If START_OF_OT = "" Then
            START_OF_OT = "06:00:00 PM"
        End If
        
        START_OF_ND = Format(Null2String(rsTemp!NDSTART), "HH:MM:SS AM/PM")
        
        OT_COMPUTE = N2Str2Zero(rsTemp!OTCOMPUTE)
        
        
        
        If START_OF_ND = "" Then
            START_OF_ND = "12:00:00 PM"
        End If
    End If
    Set rsTemp = Nothing
End Sub

Sub GetEmployeeLogs(INAM As String, OUTAM As String, INPM As String, OUTPM As String, SHIFTLOG_OPTION As Integer)
    LOGINAM = ""
    LOGOUTLUNCH = ""
    LOGINLUNCH = ""
    LOGOUTPM = ""
    If LOG_OPTION = 1 Then
        LOGINAM = INAM
        LOGOUTLUNCH = OUTAM
        LOGINLUNCH = INPM
        LOGOUTPM = OUTPM
    ElseIf LOG_OPTION = 0 Then
        LOGINAM = INAM
        LOGOUTPM = OUTPM
    End If
End Sub

Sub GetEmployeeRate(vBASICSALARY As Double, vempstatus As String)
    Dim rsSETUPDEDUCTION                               As ADODB.Recordset
    Set rsSETUPDEDUCTION = New ADODB.Recordset
    Set rsSETUPDEDUCTION = gconDMIS.Execute("SELECT * FROM HRMS_SETUPDEDUCTION")

    If vempstatus = "M" Then
        EMP_RATE_DAY = Round(((vBASICSALARY * 12) / N2Str2Zero(rsSETUPDEDUCTION!WORKING_DAY)), 2)
        EMP_RATE_HRS = Round(EMP_RATE_DAY / 8, 2)
        EMP_RATE_MIN = Round(EMP_RATE_HRS / 60, 2)
    ElseIf vempstatus = "D" Then
        EMP_RATE_DAY = Round(vBASICSALARY, 2)
        EMP_RATE_HRS = Round(EMP_RATE_DAY / 8, 2)
        EMP_RATE_MIN = Round(EMP_RATE_HRS / 60, 2)
    End If
End Sub

Sub SaveOT(OTCODE As String, OTHOUR As Double, OTAMOUNT As Double, LEVEL As String, EMPNO As String, CUTOFF As String, OTDATE As String)
    If OT_COMPUTE = 1 Then
        If chkOTRetain.Value = 1 Then Exit Sub
        
        gconDMIS.Execute ("INSERT INTO HRMS_OVERTIME " & _
                          "(EMPLEVEL, EMPNO, OCODE, DEYT, DEYT2, TOTALHR, AMOUNT, HOLIDAY, JUSTIFICATION, CUT_OFF, PAY_MONTH, PAY_YEAR) " & _
                        " VALUES ('" & LEVEL & _
                          "'," & EMPNO & _
                          ", '" & OTCODE & _
                          "', " & N2Date2Null(OTDATE) & _
                          ", " & N2Date2Null(OTDATE) & _
                          ", " & OTHOUR & _
                          ", " & OTAMOUNT & _
                          "," & N2Str2Null("") & _
                          ",'" & "SYSTEM COMPUTED" & _
                          "'," & CUTOFF & _
                          "," & What_month(cboMOnth.Text) & _
                          "," & cboyear.Text & ")")
    End If
End Sub

Private Sub chkAllowanceBase_Click()
    FillGrid
End Sub

Private Sub chkConfidential_Click()
    FillGrid
End Sub

Private Sub chkContractual_Click()
    FillGrid
End Sub

Private Sub chkProbReg_Click()
    FillGrid
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
    If Function_Access(LOGID, "Acess_Process", "PROCESS UPDATE ATTENDANCE") = False Then Exit Sub

    If MsgBox("Process Update Attendance", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub
    Dim matt                                           As String
    Dim ABSENCE                                        As Double
    Dim UNDERTIME                                      As Double
    Dim OVERTIME                                       As Double
    Dim DAYS_ENTERED                                   As Double
    Dim DAYS_ACTUAL                                    As Integer
    Dim I                                              As Integer
    Dim checkleave                                     As Boolean
    
    cmdGO.Visible = False
    cmdCancel.Visible = False
    cmdDone.Enabled = False


    If cboQuensina.Text = "1st Cut-Off" Then
        matt = "1"
    Else
        matt = "2"
    End If

    '*****************************************************************
    'DESCRIPTION : NOT DELETE THE MANUALY IMPUT BY USER
        gconDMIS.Execute ("DELETE FROM HRMS_DEDUCTIONS WHERE " & _
            " (MANUAL <> 'Y' OR MANUAL IS NULL) " & _
            " AND (PARTICULAR = 'LT' OR PARTICULAR = 'WD' OR PARTICULAR = 'HD' OR PARTICULAR = 'UT') " & _
            " AND CUT_OFF  = " & matt & _
            " AND PAY_MONTH = " & What_month(cboMOnth.Text) & _
            " AND PAY_YEAR = " & cboyear.Text & _
            " AND " & PROCESS_OPTION)
    '*****************************************************************
    If chkOTRetain.Value = 0 Then
        gconDMIS.Execute ("DELETE FROM HRMS_OVERTIME " & _
            " WHERE MANUAL <> 'Y' " & _
            " AND CUT_OFF  = " & matt & _
            " AND PAY_MONTH = " & What_month(cboMOnth.Text) & _
            " AND PAY_YEAR = " & cboyear.Text & _
            " AND " & PROCESS_OPTION)
    End If
    I = 1
    
    While I < Grid1.Rows
        Screen.MousePointer = 11
        Dim rsEmpInfo                                  As ADODB.Recordset
        Set rsEmpInfo = New ADODB.Recordset
        Set rsEmpInfo = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE EMPNO = '" & Grid1.Cell(I, 1).Text & "'")
        'Set rsEmpInfo = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE EMPNO = '200720'")
        If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
            'VARIABLES
            '==================================================
            ABSENCE = 0
            UNDERTIME = 0
            OVERTIME = 0
            DAYS_ENTERED = 0
            DAYS_ACTUAL = 0
            '==================================================

            labName.Caption = Null2String(rsEmpInfo!lastname) & ", " & Null2String(rsEmpInfo!FIRSTNAME)
            Call GetEmployeeRate(NumericVal(rsEmpInfo!BASICSALARY), Null2String(rsEmpInfo!EMPSTATUS))
            Call GetEmployeeShift(Null2String(rsEmpInfo!Shift))

            Dim rsAttend                               As ADODB.Recordset
            Set rsAttend = New ADODB.Recordset
            Set rsAttend = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND " & _
                "(DATETODAY BETWEEN '" & CDate(GENFROM) & "' AND '" & CDate(GENTO) & "') ORDER BY DATETODAY ASC")
                
            'Set rsAttend = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND " & _
                "(DATETODAY BETWEEN '11/3/2009' AND '11/3/2009') ORDER BY DATETODAY ASC")
            If Not rsAttend.EOF And Not rsAttend.BOF Then
                rsAttend.MoveFirst
                While Not rsAttend.EOF
                    Call GetEmployeeLogs(Null2String(rsAttend!INAM), Null2String(rsAttend!OUTAM), Null2String(rsAttend!INPM), Null2String(rsAttend!OUTPM), LOG_OPTION)

                    Dim ABSENCE_DET                    As Double
                    Dim UNDERTIME_DET                  As Double
                    Dim OVERTIME_DET                   As Double
                    Dim OTCODE                         As String
                    ABSENCE_DET = 0
                    UNDERTIME_DET = 0
                    OVERTIME_DET = 0
                    OTCODE = ""

                    If CheckHoliday(MONTH(rsAttend!DATETODAY), Day(rsAttend!DATETODAY)) = True And CheckWorkDay(Weekday(rsAttend!DATETODAY)) = False And CheckSpecialType(MONTH(rsAttend!DATETODAY), Day(rsAttend!DATETODAY)) = True Then
                        OTCODE = "003"
                        If LOGINAM <> "" And LOGOUTPM <> "" Then
                            OVERTIME_DET = ComputeOvertime(LOGINAM, LOGOUTPM, LOGOUTLUNCH, LOGINLUNCH, SHIFTFROM1, SHIFTTO1, SHIFTFROM2, SHIFTTO2, GRACE_PERIOD, LOG_OPTION, OTCODE, START_OF_OT, START_OF_ND, Null2String(rsEmpInfo!EMPLEVEL), Null2String(rsEmpInfo!EMPNO), matt, DateValue(Null2String(rsAttend!DATETODAY)))
                        End If
                    ElseIf CheckHoliday(MONTH(rsAttend!DATETODAY), Day(rsAttend!DATETODAY)) = True And CheckWorkDay(Weekday(rsAttend!DATETODAY)) = False And CheckSpecialType(MONTH(rsAttend!DATETODAY), Day(rsAttend!DATETODAY)) = False Then
                        OTCODE = "005"
                        If LOGINAM <> "" And LOGOUTPM <> "" Then
                            OVERTIME_DET = ComputeOvertime(LOGINAM, LOGOUTPM, LOGOUTLUNCH, LOGINLUNCH, SHIFTFROM1, SHIFTTO1, SHIFTFROM2, SHIFTTO2, GRACE_PERIOD, LOG_OPTION, OTCODE, START_OF_OT, START_OF_ND, Null2String(rsEmpInfo!EMPLEVEL), Null2String(rsEmpInfo!EMPNO), matt, DateValue(Null2String(rsAttend!DATETODAY)))
                        End If
                    ElseIf CheckHoliday(MONTH(rsAttend!DATETODAY), Day(rsAttend!DATETODAY)) = True And CheckSpecialType(MONTH(rsAttend!DATETODAY), Day(rsAttend!DATETODAY)) = True Then
                        OTCODE = "004"
                        If LOGINAM <> "" And LOGOUTPM <> "" Then
                            OVERTIME_DET = ComputeOvertime(LOGINAM, LOGOUTPM, LOGOUTLUNCH, LOGINLUNCH, SHIFTFROM1, SHIFTTO1, SHIFTFROM2, SHIFTTO2, GRACE_PERIOD, LOG_OPTION, OTCODE, START_OF_OT, START_OF_ND, Null2String(rsEmpInfo!EMPLEVEL), Null2String(rsEmpInfo!EMPNO), matt, DateValue(Null2String(rsAttend!DATETODAY)))
                        End If
                    ElseIf CheckHoliday(MONTH(rsAttend!DATETODAY), Day(rsAttend!DATETODAY)) = True And CheckSpecialType(MONTH(rsAttend!DATETODAY), Day(rsAttend!DATETODAY)) = False Then
                        OTCODE = "002"
                        If LOGINAM <> "" And LOGOUTPM <> "" Then
                            OVERTIME_DET = ComputeOvertime(LOGINAM, LOGOUTPM, LOGOUTLUNCH, LOGINLUNCH, SHIFTFROM1, SHIFTTO1, SHIFTFROM2, SHIFTTO2, GRACE_PERIOD, LOG_OPTION, OTCODE, START_OF_OT, START_OF_ND, Null2String(rsEmpInfo!EMPLEVEL), Null2String(rsEmpInfo!EMPNO), matt, DateValue(Null2String(rsAttend!DATETODAY)))
                        End If
                    ElseIf CheckWorkDay(Weekday(rsAttend!DATETODAY)) = False Then
                        OTCODE = "004"
                        If LOGINAM <> "" And LOGOUTPM <> "" Then
                            OVERTIME_DET = ComputeOvertime(LOGINAM, LOGOUTPM, LOGOUTLUNCH, LOGINLUNCH, SHIFTFROM1, SHIFTTO1, SHIFTFROM2, SHIFTTO2, GRACE_PERIOD, LOG_OPTION, OTCODE, START_OF_OT, START_OF_ND, Null2String(rsEmpInfo!EMPLEVEL), Null2String(rsEmpInfo!EMPNO), matt, DateValue(Null2String(rsAttend!DATETODAY)))
                        End If
                    ElseIf CheckWorkDay(Weekday(rsAttend!DATETODAY)) = True Then
                        ABSENCE_DET = ComputeAbsence(LOGINAM, LOGOUTPM, LOGOUTLUNCH, LOGINLUNCH, SHIFTFROM1, SHIFTTO1, SHIFTFROM2, SHIFTTO2, GRACE_PERIOD, LOG_OPTION, LOGOUTPM, DateValue(Null2String(rsAttend!DATETODAY)))
                        ABSENCE = ABSENCE + ABSENCE_DET
                        DAYS_ACTUAL = DAYS_ACTUAL + 1
                        DAYS_ENTERED = DAYS_ENTERED + (1 - ABSENCE_DET)
                        
                        '** capture leave sa update attendance
                        checkleave = IsLeaves(DateValue(Null2String(rsAttend!DATETODAY)))
                        
                        If checkleave = True Then
                            'do nothing
                        ElseIf ABSENCE_DET = 0 Then
                        
                            UNDERTIME_DET = ComputeLateAndUndertime(LOGINAM, LOGOUTPM, LOGOUTLUNCH, LOGINLUNCH, SHIFTFROM1, SHIFTTO1, SHIFTFROM2, SHIFTTO2, GRACE_PERIOD, LOG_OPTION)
                            UNDERTIME = UNDERTIME + UNDERTIME_DET
                            
                            OTCODE = "001"
                            OVERTIME_DET = ComputeOvertime(LOGINAM, LOGOUTPM, LOGOUTLUNCH, LOGINLUNCH, SHIFTFROM1, SHIFTTO1, SHIFTFROM2, SHIFTTO2, GRACE_PERIOD, LOG_OPTION, OTCODE, START_OF_OT, START_OF_ND, Null2String(rsEmpInfo!EMPLEVEL), Null2String(rsEmpInfo!EMPNO), matt, DateValue(Null2String(rsAttend!DATETODAY)))
                      
                        End If
                        
'* para ma capture leave
'                        If ABSENCE_DET = 0 Then
'                            UNDERTIME_DET = ComputeLateAndUndertime(LOGINAM, LOGOUTPM, LOGOUTLUNCH, LOGINLUNCH, SHIFTFROM1, SHIFTTO1, SHIFTFROM2, SHIFTTO2, GRACE_PERIOD, LOG_OPTION)
'                            UNDERTIME = UNDERTIME + UNDERTIME_DET
'                        End If
'                        If ABSENCE_DET = 0 Then
'                            'compute overtime
'                            OTCODE = "001"
'                            OVERTIME_DET = ComputeOvertime(LOGINAM, LOGOUTPM, LOGOUTLUNCH, LOGINLUNCH, SHIFTFROM1, SHIFTTO1, SHIFTFROM2, SHIFTTO2, GRACE_PERIOD, LOG_OPTION, OTCODE, START_OF_OT, START_OF_ND, Null2String(rsEmpInfo!EMPLEVEL), Null2String(rsEmpInfo!EMPNO), matt, DateValue(Null2String(rsAttend!DATETODAY)))
'                        End If
                    
                    
                    End If
                    
    
                    If ABSENCE_DET > 0 And Null2String(rsEmpInfo!EMPSTATUS) = "M" Then
                        If CheckIfApproveLeave(Null2String(rsAttend!DATETODAY), Null2String(rsEmpInfo!EMPNO)) = 0 Then
                            'WHOLE DAY ABSENT
                            gconDMIS.Execute ("INSERT INTO HRMS_DEDUCTIONS " & _
                                "(EMPLEVEL, EMPNO, DEYT, PARTICULAR, AMOUNT, NOMIN, CUT_OFF, PAY_MONTH, PAY_YEAR) " & _
                                " VALUES ('" & rsEmpInfo!EMPLEVEL & _
                                "', " & N2Str2Null(rsEmpInfo!EMPNO) & _
                                ", " & N2Date2Null(rsAttend!DATETODAY) & _
                                ", '" & "WD" & _
                                "', " & (ABSENCE_DET * EMP_RATE_DAY) & _
                                ", " & (ABSENCE_DET * 480) & _
                                ", " & matt & _
                                ", " & What_month(cboMOnth) & _
                                ", " & cboyear.Text & ")")
                        ElseIf CheckIfApproveLeave(Null2String(rsAttend!DATETODAY), Null2String(rsEmpInfo!EMPNO)) = 2 Then
                            'HALF DAY LEAVE
                            gconDMIS.Execute ("INSERT INTO HRMS_DEDUCTIONS " & _
                                "(EMPLEVEL, EMPNO, DEYT, PARTICULAR, AMOUNT, NOMIN, CUT_OFF, PAY_MONTH, PAY_YEAR) " & _
                                " VALUES ('" & rsEmpInfo!EMPLEVEL & _
                                "', " & N2Str2Null(rsEmpInfo!EMPNO) & _
                                ", " & N2Date2Null(rsAttend!DATETODAY) & _
                                ", '" & "WD" & _
                                "', " & ((ABSENCE_DET / 2) * EMP_RATE_DAY) & _
                                ", " & (ABSENCE_DET * 240) & _
                                ", " & matt & _
                                ", " & What_month(cboMOnth) & _
                                ", " & cboyear.Text & ")")
                        Else
                            'WHOLE DAY LEAVE
                        End If
                    End If
                    If UNDERTIME_DET > 0 Then
                        gconDMIS.Execute ("INSERT INTO HRMS_DEDUCTIONS " & _
                            "(EMPLEVEL, EMPNO, DEYT, PARTICULAR, AMOUNT, NOMIN, CUT_OFF, PAY_MONTH, PAY_YEAR) " & _
                            "VALUES ('" & rsEmpInfo!EMPLEVEL & _
                            "', " & N2Str2Null(rsEmpInfo!EMPNO) & _
                            ", " & N2Date2Null(rsAttend!DATETODAY) & _
                            ", '" & "LT" & _
                            "', " & (UNDERTIME_DET * EMP_RATE_MIN) & _
                            ", " & UNDERTIME_DET & _
                            ", " & matt & _
                            ", " & What_month(cboMOnth) & _
                            ", " & cboyear.Text & ")")
                    End If
                    
'    If cboQuensina.Text = "1st Cut-Off" Then
'        matt = "1"
'    Else
'        matt = "2"
'    End If
    
CONT_NOV:
                    rsAttend.MoveNext
                Wend
            End If
        End If
        gconDMIS.Execute "DELETE FROM HRMS_DAILYMONITORING " & _
            " WHERE EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & _
            " And CUT_OFF = " & matt & _
            " AND PAY_MONTH = " & What_month(cboMOnth) & _
            " AND PAY_YEAR = " & cboyear.Text & ""

        gconDMIS.Execute "INSERT INTO HRMS_DAILYMONITORING " & _
            "(EMPLEVEL, EMPNO, DEYT, ACTUAL, ENTERED, CUT_OFF, PAY_MONTH, PAY_YEAR) " & _
            " Values ('" & rsEmpInfo!EMPLEVEL & _
            "', " & N2Str2Null(rsEmpInfo!EMPNO) & _
            ", " & N2Date2Null(GENTO) & _
            ", " & DAYS_ACTUAL & _
            ", " & DAYS_ENTERED & _
            ", " & matt & _
            ", " & What_month(cboMOnth.Text) & _
            ", " & cboyear.Text & ")"
            
        I = I + 1
        gauProgress.Value = (I / Grid1.Rows) * 100
        lblPercent.Caption = Int(gauProgress.Value) & "%"
        DoEvents
    Wend
    Set rsEmpInfo = Nothing
    Set rsAttend = Nothing
    cmdDone.Enabled = True
    
    'MsgBox "Update Attendance complete!", vbInformation, "HRMS"
    MessagePop InfoFriend, "Process Complete", "Update Attendance complete. Kindly double check the Output in Attendace Deduction."
    Screen.MousePointer = 0
End Sub

Function CheckIfApproveLeave(xDateToday As Date, XEMPNO As String) As Integer
    'IF FUNCTION RETURN
    '0 - NOT VALID LEAVE
    '1 - 1 DAY LEAVE
    '2 - 1/2 DAY LEAVE
    
    Dim RSTMP As New ADODB.Recordset
    Dim matt As String
    If cboQuensina.Text = "1st Cut-Off" Then
        matt = "1"
    Else
        matt = "2"
    End If
'    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_REQUESTLEAVE_OT WHERE " & _
'        " CUT_OFF = " & matt & _
'        " AND PAY_MONTH = " & What_month(cboMonth) & _
'        " AND PAY_YEAR = " & cboYear & _
'        " AND EMPNO = " & XEMPNO & _
'        " AND STATUS = 'A'")
    
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_REQUESTLEAVE_OT WHERE " & _
        " CUT_OFF = " & matt & _
        " AND PAY_MONTH = " & MONTH(CStr(xDateToday)) & _
        " AND PAY_YEAR = " & cboyear & _
        " AND EMPNO = " & XEMPNO & _
        " AND STATUS = 'A' AND " & N2Str2Null(xDateToday) & " BETWEEN DTE_FROM AND DTE_TO")
        
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        'COMMENT BY  : MJP 01142010 0600PM
        'DESCRIPTION : UPDATE TO CAPTURE ALL KIND OF LEAVE SCHEDULE
            '        If Hour(Null2String(RSTMP!OT_FROM)) = 8 Then
            '            If Hour(Null2String(RSTMP!OT_TO)) = 17 Then
            '                '1 DAY LEAVE = 1
            '                CheckIfApproveLeave = 1
            '            Else
            '                '1/2 DAY LEAVE = 2
            '                CheckIfApproveLeave = 2
            '            End If
            '        ElseIf Hour(Null2String(RSTMP!OT_FROM)) = 12 Then
            '            If Hour(Null2String(RSTMP!OT_TO)) = 17 Then
            '                '1/2 DAY LEAVE = 2
            '                CheckIfApproveLeave = 2
            '            Else
            '                'NOT VALID = 0
            '                CheckIfApproveLeave = 0
            '            End If
            '        Else
            '            'NOT VALID = 0
            '            CheckIfApproveLeave = 0
            '        End If
        'COMMENT BY  : MJP 01142010 0600PM
        
        'UPDATE BY   : MJP 01142010 0600PM
        'DESCRIPTION : UPDATE TO CAPTURE ALL KIND OF LEAVE SCHEDULE
            If Hour(Null2String(RSTMP!OT_FROM)) = 8 Then
                If Hour(Null2String(RSTMP!OT_TO)) = 17 Then
                    CheckIfApproveLeave = 1
                ElseIf Hour(Null2String(RSTMP!OT_TO)) = 18 Then
                    CheckIfApproveLeave = 1
                ElseIf Hour(Null2String(RSTMP!OT_TO)) = 12 Then
                    CheckIfApproveLeave = 2
                ElseIf Hour(Null2String(RSTMP!OT_TO)) = 13 Then
                    CheckIfApproveLeave = 2
                Else
                    CheckIfApproveLeave = 0
                End If
            ElseIf Hour(Null2String(RSTMP!OT_FROM)) = 9 Then
                If Hour(Null2String(RSTMP!OT_TO)) = 17 Then
                    CheckIfApproveLeave = 1
                ElseIf Hour(Null2String(RSTMP!OT_TO)) = 18 Then
                    CheckIfApproveLeave = 1
                ElseIf Hour(Null2String(RSTMP!OT_TO)) = 12 Then
                    CheckIfApproveLeave = 2
                ElseIf Hour(Null2String(RSTMP!OT_TO)) = 13 Then
                    CheckIfApproveLeave = 2
                Else
                    CheckIfApproveLeave = 0
                End If
            ElseIf Hour(Null2String(RSTMP!OT_FROM)) = 12 Then
                If Hour(Null2String(RSTMP!OT_TO)) = 17 Then
                    CheckIfApproveLeave = 2
                ElseIf Hour(Null2String(RSTMP!OT_TO)) = 18 Then
                    CheckIfApproveLeave = 2
                Else
                    CheckIfApproveLeave = 0
                End If
            ElseIf Hour(Null2String(RSTMP!OT_FROM)) = 1 Then
                If Hour(Null2String(RSTMP!OT_TO)) = 17 Then
                    CheckIfApproveLeave = 2
                ElseIf Hour(Null2String(RSTMP!OT_TO)) = 18 Then
                    CheckIfApproveLeave = 2
                Else
                    CheckIfApproveLeave = 0
                End If
            Else
                CheckIfApproveLeave = 0
            End If
        'UPDATE BY   : MJP 01142010 0600PM
    End If
    Set RSTMP = Nothing
End Function

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'DrawXPCtl Me
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim rsCutoff                                       As ADODB.Recordset
    Set rsCutoff = New ADODB.Recordset
    Set rsCutoff = gconDMIS.Execute("SELECT PERIODMONTH,PERIODYEAR,NOTEDBY2 FROM HRMS_PAYROLLSETUP")
    If Not (rsCutoff.EOF And rsCutoff.BOF) Then
        If NumericVal(rsCutoff!NOTEDBY2) = 1 Then
            cboQuensina.Clear
            cboQuensina.AddItem "1st Cut-Off"
            cboQuensina.Text = "1st Cut-Off"
        ElseIf NumericVal(rsCutoff!NOTEDBY2) = 2 Then
            cboQuensina.Clear
            cboQuensina.AddItem "2nd Cut-Off"
            cboQuensina.Text = "2nd Cut-Off"
        Else
            MsgBox "Cut-off not set"
        End If
        cboMOnth.Clear
        cboMOnth.AddItem MonthName(Null2String(rsCutoff!PERIODMONTH))
        cboMOnth.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
        cboyear.Clear
        'cboyear.AddItem "2009"  'Null2String(rsCutoff!PERIODYEAR)
        'cboyear.Text = "2009" ' Null2String(rsCutoff!PERIODYEAR)
        
        cboyear.AddItem Null2String(rsCutoff!PERIODYEAR)
        cboyear.Text = Null2String(rsCutoff!PERIODYEAR)
    End If
    InitGrid
    FillGrid
    labName.Caption = ""
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub Grid1_DblClick()
    If Grid1.ActiveCell.Col = 3 Then
        Grid1.RemoveItem (Grid1.ActiveCell.Row)
        Grid1.Refresh
    End If
End Sub
Function IsLeaves(xdate As String) As Boolean

    Dim I As Integer
    Dim leave As New ADODB.Recordset
    Dim reqdesc As String

    I = 1
    
    Set leave = gconDMIS.Execute("SELECT * FROM HRMS_REQUESTLEAVE_OT where '" & xdate & "' between dte_from  and  dte_to  and status = 'A' and empno = '" & Grid1.Cell(I, 1).Text & "'")

    If Not (leave.BOF And leave.EOF) Then
        reqdesc = Null2String(leave!reqdesc)
        IsLeaves = True
    Else
        IsLeaves = False
    End If

End Function

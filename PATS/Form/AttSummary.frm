VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
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
Public DaysSelect As Integer
Public EmploymentStatus As String
Public Filter As String

Private Sub Parse()
Dim H As String
Dim H1, H2, H3, i, v As Integer
            
H1 = 0: H2 = 0: H3 = 0
H = Trim(TxtHolidays)
If H <> "" Then
   v = InStr(H, ",")
   If v = 0 Then
      H1 = Val(H)
   Else
      H1 = Val(Left(H, v - 1))
      i = InStr(v + 1, H, ",")
      If i = 0 Then
         H2 = Val(Mid(H, v + 1))
      Else
         H2 = Val(Mid(H, v + 1, i - v - 1))
         H3 = Val(Mid(H, i + 1))
      End If
   End If
End If
End Sub


Private Sub cmdCancel_Click()
Unload Me
frmLOGIN.TxtEmpNumber.SetFocus
End Sub

Private Sub CmdShow_Click()
Dim c, d, k, X, z, w, u1, u2, u3, u4, tu, st, en, i, d1, d2 As Integer
Dim DaysAbsent, ActualWorkDays As Single
Dim UnderTime, UThrs, UTmins, Late, LateHrs, LateMins, Overtime, OTHrs, OTMins As Integer
Dim DaysOfWeek, Dow, DayList, Pe As String
ReDim da(31, 5) As String
  
Dim tdate As Date
Dim theMonth As Integer
Dim theYear As Integer
tdate = Date
  
If cboMonth.Text = Date2Month(OneMonth(Date, -2)) Then
   tdate = OneMonth(Date, -2)
ElseIf cboMonth.Text = Date2Month(OneMonth(Date, -1)) Then
   tdate = OneMonth(Date, -1)
ElseIf cboMonth.Text = Date2Month(Date) Then
   tdate = Date
ElseIf cboMonth.Text = Date2Month(OneMonth(Date, 1)) Then
   tdate = OneMonth(Date, 1)
End If

theMonth = Month(tdate)
theYear = Year(tdate)
'Parse the holidays
Dim H1(31), H, Dy As String
Dim C1(10) As Integer

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
gconDMIS.Execute "Delete from HRMS_TempSummary"
FrmProgBar.Visible = True
Me.Refresh
Lbl1ProgBar.Caption = "Creating Summary,  Please wait..."
'rsEmpInfo.MoveFirst
i = 0
Set rsEmpInfo = New ADODB.Recordset
Set rsEmpInfo = gconDMIS.Execute("Select * from HRMS_EmpInfo order by EmpNo asc")
Do Until rsEmpInfo.EOF
i = i + 1
   ProgressBar1.Value = (i / rsEmpInfo.RecordCount) * 100
   Lbl2ProgBar.Caption = Int(ProgressBar1.Value) & "%"
   Me.Refresh
   'If Null2String(rsEmpInfo!divcode) = thedivcode Then
      If Null2String(rsEmpInfo!ACTIVEINACTIVE) = "A" Then
         'If Null2String(rsEmpInfo!EMPLEVEL) = EmploymentStatus Then
            Set rsCard = New ADODB.Recordset
                rsCard.Open "Select * from HRMS_Attend Where right(EmpNo,4) = " & N2Str2Null(rsEmpInfo!empno) & " AND Month(DateToday) = " & Month(tdate) & " AND Year(DateToday) = " & Year(tdate), gconDMIS
            If Not rsCard.BOF And Not rsCard.EOF Then
               rsCard.MoveFirst
               Do Until rsCard.EOF
                  d = Day(rsCard!DateToday)
                  da(d, 1) = rsCard!DateToday
                  da(d, 2) = Format(Null2String(rsCard!InAm), "HH:MM:SS AM/PM")
                  da(d, 3) = Format(Null2String(rsCard!OutAm), "HH:MM:SS AM/PM")
                  da(d, 4) = Format(Null2String(rsCard!InPm), "HH:MM:SS AM/PM")
                  da(d, 5) = Format(Null2String(rsCard!OutPM), "HH:MM:SS AM/PM")
                                                              
                  If d > c Then c = d
                  rsCard.MoveNext
               Loop
               For k = 1 To c
                   If da(k, 1) = "" Then
                      da(k, 1) = CDate(Month(tdate) & "/" & Str(k) & "/" & Right(Year(tdate), 2))
                      da(k, 2) = ""
                      da(k, 3) = ""
                      da(k, 4) = ""
                      da(k, 5) = ""
                   End If
               Next k
               ActualWorkDays = 0: DaysAbsent = 0: UnderTime = 0: Late = 0: Overtime = 0: DayList = ""
               If DaysSelect = 0 Then
                  'st = 1: en = 15: Pe = "01-15"
                  If c > 15 Then
                     c = 15
                  End If
                  st = 1: en = c: Pe = "01-15"
               ElseIf DaysSelect = 1 Then
                  st = 16: en = c: Pe = "16-31"
               Else
                  st = 1: en = c: Pe = "01-31"
               End If
               For k = st To en
                   w = Weekday(da(k, 1))
                   'Dow = Mid(DaysOfWeek, (w - 1) * 3 + 1, 3)
                   'If w = 2 Or w = 3 Or w = 4 Or w = 5 Or w = 6 Then (Until Friday Only)
                   If w = 2 Or w = 3 Or w = 4 Or w = 5 Or w = 6 Or w = 7 Then
                      If H1(k) <> "1" Then
                   Stop
                         u1 = 0: u2 = 0: u3 = 0: u4 = 0: tu = 0
                         If da(k, 2) = "" And da(k, 3) = "" And da(k, 4) = "" And da(k, 5) = "" Then
                            If H1(k) = "" Then
                               DaysAbsent = DaysAbsent + 1
                               DayList = DayList + Str(k)
                            ElseIf H1(k) = "A" Then
                               DaysAbsent = DaysAbsent + 0.5
                               DayList = DayList + Str(k) + "P"
                            ElseIf H1(k) = "P" Then
                               DaysAbsent = DaysAbsent + 0.5
                               DayList = DayList + Str(k) + "A"
                            End If
                            Stop
                         ElseIf da(k, 2) = "" Or da(k, 3) = "" Or da(k, 4) = "" Or da(k, 5) = "" Then
                            If H1(k) = "" Or H1(k) = "P" Then
                               d1 = 0
                               If da(k, 2) = "" Or da(k, 3) = "" Then
                                  DaysAbsent = DaysAbsent + 0.5
                                  d1 = k
                               End If
                               If da(k, 2) <> "" And da(k, 3) <> "" Then
                                  If da(k, 2) > #8:00:00 AM# Then
                                     u1 = DateDiff("n", #8:00:00 AM#, da(k, 2))
                                  End If
                                  If da(k, 3) < #11:59:59 AM# Then
                                     u2 = DateDiff("n", da(k, 3), #11:59:00 AM#) + 1
                                  End If
                               End If
                            End If
                            If H1(k) = "" Or H1(k) = "A" Then
                               d2 = 0
                               If da(k, 4) = "" Or da(k, 5) = "" Then
                                  DaysAbsent = DaysAbsent + 0.5
                                  d2 = k
                               End If
                               If da(k, 4) <> "" And da(k, 5) <> "" Then
                                  If da(k, 4) > #1:00:00 PM# Then
                                     u3 = DateDiff("n", #1:00:00 PM#, da(k, 4))
                                  End If
                                  If da(k, 5) < #5:00:00 PM# Then
                                     u4 = DateDiff("n", da(k, 5), #5:00:00 PM#)
                                  End If
                               End If
                            End If
                            If d1 = d2 Then
                               DayList = DayList + Str(k)
                            ElseIf d1 > 0 Then
                               DayList = DayList + Str(k) + "A"
                            ElseIf d2 > 0 Then
                               DayList = DayList + Str(k) + "P"
                            End If
                         Else
                            If H1(k) = "" Or H1(k) = "P" Then
                               If da(k, 2) > #8:00:00 AM# Then
                                  u1 = DateDiff("n", #8:00:00 AM#, da(k, 2))
                               End If
                               If da(k, 3) < #11:59:59 AM# Then
                                  u2 = DateDiff("n", da(k, 3), #11:59:00 AM#) + 1
                               End If
                            End If
                            If H1(k) = "" Or H1(k) = "A" Then
                               If da(k, 4) > #1:00:00 PM# Then
                                  u3 = DateDiff("n", #1:00:00 PM#, da(k, 4))
                               End If
                               If da(k, 5) < #5:00:00 PM# Then
                                  u4 = DateDiff("n", da(k, 5), #5:00:00 PM#)
                               End If
                            End If
                         End If
                         tu = u1 + u2 + u3 + u4
                         UnderTime = UnderTime + tu
                         Late = Late + tu
                         Overtime = Overtime + tu
                       End If
                   End If
               Next k
               'If DaysAbsent > 0 Or UnderTime > 0 Then
                   UThrs = Int(UnderTime / 60)
                   UTmins = UnderTime Mod 60
                   LateHrs = Int(Late / 60)
                   LateMins = Late Mod 60
                   OTHrs = Int(Overtime / 60)
                   OTMins = Overtime Mod 60
                   gconDMIS.Execute "Insert into HRMS_TempSummary (EmpNo, ActualWorkDays, DaysAbsent, UThrs, UTmins, LateHrs, LateMins, OTHrs, OTMins, Days, Period, EMPLEVEL) " & _
                                    "Values (" & N2Str2Null(rsEmpInfo!empno) & "," & ActualWorkDays & ", " & DaysAbsent & "," & UThrs & "," & UTmins & "," & LateHrs & "," & LateMins & "," & OTHrs & "," & OTMins & ",'" & DayList & "','" & Pe & "','" & EmploymentStatus & "')"
               'End If
            End If
         'End If
      End If
   'End If
   rsEmpInfo.MoveNext
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
Dim X As Integer
Dim Thedate As Date
Thedate = OneMonth(Date, -3)
For X = 1 To 4
    Thedate = OneMonth(Thedate, 1)
    cboMonth.AddItem Date2Month(CDate(Thedate))
Next
End Sub

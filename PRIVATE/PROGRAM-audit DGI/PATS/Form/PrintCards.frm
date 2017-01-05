VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrintCards 
   Caption         =   "PATS"
   ClientHeight    =   2655
   ClientLeft      =   4365
   ClientTop       =   3705
   ClientWidth     =   3600
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2655
   ScaleWidth      =   3600
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
      Left            =   1890
      TabIndex        =   7
      Top             =   2130
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Printer Set-Up"
      Height          =   390
      Left            =   1920
      TabIndex        =   6
      Top             =   1365
      Width           =   1485
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1920
      TabIndex        =   5
      Top             =   810
      Width           =   1485
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "O K"
      Height          =   390
      Left            =   1890
      TabIndex        =   4
      Top             =   270
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      Caption         =   "Days Selection"
      Height          =   1845
      Left            =   90
      TabIndex        =   0
      Top             =   210
      Width           =   1710
      Begin VB.OptionButton Option1 
         Caption         =   "01 - 15"
         Height          =   345
         Index           =   0
         Left            =   375
         TabIndex        =   1
         Top             =   270
         Width           =   990
      End
      Begin VB.OptionButton Option1 
         Caption         =   "01 - 31"
         Height          =   345
         Index           =   2
         Left            =   375
         TabIndex        =   3
         Top             =   1005
         Width           =   990
      End
      Begin VB.OptionButton Option1 
         Caption         =   "16 - 31"
         Height          =   300
         Index           =   1
         Left            =   375
         TabIndex        =   2
         Top             =   660
         Width           =   945
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   -45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   225
      TabIndex        =   8
      Top             =   2175
      Width           =   1455
   End
End
Attribute VB_Name = "frmPrintCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DaysSelect As Integer

Private Sub cmdCancel_Click()
Unload Me
frmViewCards.TxtEmpNumber.SetFocus
End Sub

Private Sub CmdOK_Click()
On Error Resume Next
Dim da(31, 5) As String
Dim c, d, k, X, w, u1, u2, u3, u4, tu, st, en, t As Integer
Dim DaysAbsent As Single
Dim UnderTime, UThrs, UTmins As Integer
Dim DaysOfWeek, Dow As String

Dim Criteria As String
Dim tdate As Date
Dim theMonth As Integer
Dim theYear As Integer
tdate = Date
        
If frmViewCards.cboMonth.Text = Date2Month(OneMonth(Date, -2)) Then
   tdate = OneMonth(Date, -2)
ElseIf frmViewCards.cboMonth.Text = Date2Month(OneMonth(Date, -1)) Then
   tdate = OneMonth(Date, -1)
ElseIf frmViewCards.cboMonth.Text = Date2Month(Date) Then
   tdate = Date
ElseIf frmViewCards.cboMonth.Text = Date2Month(OneMonth(Date, 1)) Then
   tdate = OneMonth(Date, 1)
End If
     
theMonth = Month(tdate)
theYear = Year(tdate)
        
        
'Parse the holidays
Dim H1(31), H, Dy As String
Dim C1(10), i  As Integer

H = Trim(TxtHolidays): k = 1
If H <> "" Then
   'find location of commas
   For i = 1 To Len(H)
       If Mid(H, i, 1) = "," Then
          k = k + 1: C1(k) = i
       End If
   Next i
   C1(k + 1) = Len(H) + 1
   'Place date markers ("0" not a holiday, "1" holiday, "A" holiday AM only, "P" holiday PM only)
   For i = 1 To k
       Dy = Mid(H, C1(i) + 1, C1(i + 1) - (C1(i) + 1))
       If Val(Dy) <= 31 Then
          If UCase(Right(Dy, 1)) = "A" Then
             H1(Val(Dy)) = "A"
          ElseIf UCase(Right(Dy, 1)) = "P" Then
             H1(Val(Dy)) = "P"
          Else
             H1(Val(Dy)) = "1"
          End If
       End If
   Next i
End If
        
'****************************************
'added line
Criteria = "Select * from HRMS_Attend Where right(EmpNo,4) = " & rsEmpInfo!empno & " AND Month(DateToday) = " & theMonth & " AND Year(DateToday) = " & theYear

Set rsCard = gconDMIS.OpenADODB.Recordset(Criteria)
If Not (rsCard.BOF) And Not (rsCard.EOF) Then
        
    'up here
    c = 0
    rsCard.MoveFirst
    Do Until rsCard.EOF
       d = Day(rsCard!DateToday)
       da(d, 1) = rsCard!DateToday
       da(d, 2) = Null2String(rsCard!InAm)
       da(d, 3) = Null2String(rsCard!OutAm)
       da(d, 4) = Null2String(rsCard!InPm)
       da(d, 5) = Null2String(rsCard!OutPM)
       If d > c Then c = d
       rsCard.MoveNext
    Loop
    For k = 1 To c
        If da(k, 1) = "" Then
           da(k, 1) = CDate(theMonth & "/" & Str(k) & "/" & Right(theYear, 2))
           da(k, 2) = ""
           da(k, 3) = ""
           da(k, 4) = ""
           da(k, 5) = ""
        End If
    Next k
    DaysAbsent = 0: UnderTime = 0
    DaysOfWeek = "SUNMONTUEWEDTHUFRISAT"
            
    If DaysSelect = 0 Then
       st = 1: en = 15
    ElseIf DaysSelect = 1 Then
       st = 16: en = c
    Else
       st = 1: en = c
    End If
       
    Printer.FontName = "Arial"
    Printer.FontSize = 10
            
    Printer.Print
    Printer.Print
    'Printer.Print Tab(18); "Republic of the Philippines"
    Printer.FontBold = True
    Printer.FontSize = 14
    Printer.Print Tab(5); frmLOGIN.Caption
    Printer.FontSize = 10
    Printer.FontBold = False
    'Printer.Print Tab(25); "City of Naga"
                                    
    Printer.Print
    Printer.Print UCase(frmViewCards.TxtEmpName.Text);
    Printer.Print Tab(56); Format(frmViewCards.TxtEmpNumber.Text, "0000")
    Printer.Print String(80, "-")
            
    For k = st To en
        If da(k, 1) <> "" Then
           w = Weekday(da(k, 1))
           Dow = Mid(DaysOfWeek, (w - 1) * 3 + 1, 3)
                            
           Printer.Print Format(da(k, 1), "mm/dd/yy ");
           Printer.Print Dow;
                            
           If H1(k) = "1" Then
              Printer.Print Tab(18); "HOLIDAY"; Tab(29); "HOLIDAY"; Tab(40); "HOLIDAY"; Tab(51); "HOLIDAY";
           End If
                            
           If w = 2 Or w = 3 Or w = 4 Or w = 5 Or w = 6 Then
              If H1(k) <> "1" Then
                 For X = 2 To 5
                     t = ((X - 1) * 12) + (8 - X)
                     If (X = 2 And H1(k) = "A") Or (X = 4 And H1(k) = "P") Then
                        Printer.Print Tab(t); "HOLIDAY";
                     ElseIf (X = 3 And H1(k) = "A") Or (X = 5 And H1(k) = "P") Then
                        Printer.Print Tab(t); "HOLIDAY";
                     ElseIf da(k, X) = "" Then
                        Printer.Print Tab(t); "";
                     Else
                        Printer.Print Tab(t); Format(da(k, X), "hh:mm AM/PM");
                     End If
                 Next X
                 Printer.Print
                 u1 = 0: u2 = 0: u3 = 0: u4 = 0: tu = 0
                 If da(k, 2) = "" And da(k, 3) = "" And da(k, 4) = "" And da(k, 5) = "" Then
                    If H1(k) = "" Then
                       DaysAbsent = DaysAbsent + 1
                    ElseIf H1(k) = "A" Or H1(k) = "P" Then
                       DaysAbsent = DaysAbsent + 0.5
                    End If
                 ElseIf da(k, 2) = "" Or da(k, 3) = "" Or da(k, 4) = "" Or da(k, 5) = "" Then
                    If H1(k) = "" Or H1(k) = "P" Then
                       If da(k, 2) = "" Or da(k, 3) = "" Then
                          DaysAbsent = DaysAbsent + 0.5
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
                       If da(k, 4) = "" Or da(k, 5) = "" Then
                          DaysAbsent = DaysAbsent + 0.5
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
              Else
                 Printer.Print
              End If
           Else
              Printer.Print
           End If
        End If
    Next k
                    
    UThrs = Int(UnderTime / 60)
    UTmins = UnderTime Mod 60
            
    Printer.Print String(80, "-")
    Printer.Print "Days Absent = "; Format(DaysAbsent, "###0.0"); " Day(s)";
    Printer.Print Tab(31); "UnderTime = "; Format(UThrs, "##0hrs"); Format(UTmins, "  ##0mins")
    
    Printer.Print
    Printer.Print
    
    Printer.Print "________________________" + Space(3) + "________________________"
    'Printer.Print Tab(34); "______________________"
                                            
    Printer.Print rsEmpInfo!EmpName;
    'If UCase(Trim(rsEmpInfo!empno)) = OfficeHeadNo Then
    '   Printer.Print Tab(34); "Sulpicio S. Roco, Jr."
    'Else
    '   Printer.Print Tab(34); Officehead
    'End If
    Erase da
    Printer.EndDoc
End If
frmPrintCards.Hide
frmViewCards.TxtEmpNumber.SetFocus
End Sub

Private Sub Command1_Click()
CommonDialog1.Action = 5
End Sub

Private Sub Option1_Click(Index As Integer)
DaysSelect = Index
CmdOK.SetFocus
End Sub

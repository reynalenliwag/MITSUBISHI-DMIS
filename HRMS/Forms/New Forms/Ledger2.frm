VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Begin VB.Form frmHRMS_Ledger2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ledger"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   14160
   Begin VB.CommandButton Command1 
      Caption         =   "SEARCH"
      Height          =   465
      Left            =   5880
      TabIndex        =   9
      Top             =   870
      Width           =   825
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4680
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   930
      Width           =   1035
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3510
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   930
      Width           =   1035
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   930
      Width           =   1275
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   1950
      TabIndex        =   1
      Top             =   8910
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin FlexCell.Grid Grid1 
      Height          =   6375
      Left            =   60
      TabIndex        =   0
      Top             =   1560
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   11245
      BackColor2      =   12907725
      BackColorBkg    =   -2147483645
      Cols            =   5
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      Rows            =   30
   End
   Begin VB.Label Label6 
      Caption         =   "Label4"
      Height          =   195
      Left            =   4830
      TabIndex        =   12
      Top             =   1350
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Label4"
      Height          =   195
      Left            =   3660
      TabIndex        =   11
      Top             =   1350
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   1350
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Year"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label2 
      Caption         =   "Month"
      Height          =   255
      Left            =   3810
      TabIndex        =   5
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Cut-off"
      Height          =   255
      Left            =   2550
      TabIndex        =   4
      Top             =   600
      Width           =   555
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14145
      _Version        =   655364
      _ExtentX        =   24950
      _ExtentY        =   820
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmHRMS_Ledger2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPayPeriod As ADODB.Recordset

Private Sub Command1_Click()
    rsrefresh
    InitGrid
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    DrawXPCtl Me
    FillCombo
    InitGrid
    Screen.MousePointer = 0
End Sub

Sub InitGrid()
    With Grid1
        .Cols = 23
        .Column(0).Width = 30
        .Column(1).Width = 35
        .Column(2).Width = 70
        .Column(3).Width = 200
        .Column(4).Width = 60
        .Column(5).Width = 50
        .Column(6).Width = 50
        .Column(7).Width = 50
        .Column(8).Width = 50
        .Column(9).Width = 55
        .Column(10).Width = 55
        .Column(11).Width = 50
        .Column(12).Width = 50
        .Column(13).Width = 50
        .Column(14).Width = 50
        .Column(15).Width = 50
        .Column(16).Width = 50
        .Column(17).Width = 50
        .Column(18).Width = 50
        .Column(19).Width = 50
        .Column(20).Width = 50
        .Column(21).Width = 50
        .Column(22).Width = 50
        .Cell(0, 0).Text = "L/N"
        .Cell(0, 1).Text = "LEVEL"
        .Cell(0, 2).Text = "EMPNO"
        .Cell(0, 3).Text = "NAME"
        .Cell(0, 4).Text = "BASIC"
        .Cell(0, 5).Text = "OT"
        .Cell(0, 6).Text = "LWOP"
        .Cell(0, 7).Text = "LT/UT"
        .Cell(0, 8).Text = "TAX ADJ"
        .Cell(0, 9).Text = "NTAX ADJ"
        .Cell(0, 10).Text = "OTHERDED"
        .Cell(0, 11).Text = "PAG"
        .Cell(0, 12).Text = "PHIC"
        .Cell(0, 13).Text = "SSS"
        .Cell(0, 14).Text = "TAX"
        .Cell(0, 15).Text = "LOANS"
        .Cell(0, 16).Text = "ALLOWANCE"
        .Cell(0, 17).Text = "WAGE"
        .Cell(0, 18).Text = "NET"
        .Cell(0, 19).Text = "PAGIBIGR"
        .Cell(0, 20).Text = "PHICR"
        .Cell(0, 21).Text = "SSSR"
        .Cell(0, 22).Text = "OPTION"
        
        .Column(0).Locked = True
        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True
        .Column(6).Locked = True
        .Column(7).Locked = True
        .Column(8).Locked = True
        .Column(9).Locked = True
        .Column(10).Locked = True
        .Column(11).Locked = True
        .Column(12).Locked = True
        .Column(13).Locked = True
        .Column(14).Locked = True
        .Column(15).Locked = True
        .Column(16).Locked = True
        .Column(17).Locked = True
        .Column(18).Locked = True
        .Column(19).Locked = True
        .Column(20).Locked = True
        .Column(21).Locked = True
        .Column(22).Locked = True
        .Column(4).DecimalLength = 2
        .Column(4).Mask = cellValue
        .Column(5).DecimalLength = 2
        .Column(5).Mask = cellValue
        .Column(6).DecimalLength = 2
        .Column(6).Mask = cellValue
        .Column(7).DecimalLength = 2
        .Column(7).Mask = cellValue
        .Column(8).DecimalLength = 2
        .Column(8).Mask = cellValue
        .Column(9).DecimalLength = 2
        .Column(9).Mask = cellValue
        .Column(10).DecimalLength = 2
        .Column(10).Mask = cellValue
        .Column(11).DecimalLength = 2
        .Column(11).Mask = cellValue
        .Column(12).DecimalLength = 2
        .Column(12).Mask = cellValue
        .Column(13).DecimalLength = 2
        .Column(13).Mask = cellValue
        .Column(14).DecimalLength = 2
        .Column(14).Mask = cellValue
        .Column(15).DecimalLength = 2
        .Column(15).Mask = cellValue
        .Column(16).DecimalLength = 2
        .Column(16).Mask = cellValue
        .Column(17).DecimalLength = 2
        .Column(17).Mask = cellValue
        .Column(18).DecimalLength = 2
        .Column(18).Mask = cellValue
        .Column(19).DecimalLength = 2
        .Column(19).Mask = cellValue
        .Column(20).DecimalLength = 2
        .Column(20).Mask = cellValue
        .Column(21).DecimalLength = 2
        .Column(21).Mask = cellValue
        .Column(22).DecimalLength = 2
        .Column(22).Mask = cellValue
        '.Column(23).DecimalLength = 2
        '.Column(23).Mask = cellValue
        .Column(0).Alignment = cellCenterCenter
        .Column(1).Alignment = cellCenterCenter
        .Column(2).Alignment = cellCenterCenter
        .Column(3).Alignment = CellLeft
        .Column(4).Alignment = cellRightGeneral
        .Column(5).Alignment = cellRightGeneral
        .Column(6).Alignment = cellRightGeneral
        .Column(7).Alignment = cellRightGeneral
        .Column(8).Alignment = cellRightGeneral
        .Column(9).Alignment = cellRightGeneral
        .Column(10).Alignment = cellRightGeneral
        .Column(11).Alignment = cellRightGeneral
        .Column(12).Alignment = cellRightGeneral
        .Column(13).Alignment = cellRightGeneral
        .Column(14).Alignment = cellRightGeneral
        .Column(15).Alignment = cellRightGeneral
        .Column(16).Alignment = cellRightGeneral
        .Column(17).Alignment = cellRightGeneral
        .Column(18).Alignment = cellRightGeneral
        .Column(19).Alignment = cellRightGeneral
        .Column(20).Alignment = cellRightGeneral
        .Column(21).Alignment = cellRightGeneral
        .Column(22).Alignment = cellRightGeneral
        '.Column(23).Alignment = cellRightGeneral
        
    End With
End Sub

Sub FillCombo()
    Combo1.AddItem "1st Cut-off"
    Combo1.AddItem "2nd Cut-off"
    Combo2.AddItem MonthName(1)
    Combo2.AddItem MonthName(2)
    Combo2.AddItem MonthName(3)
    Combo2.AddItem MonthName(4)
    Combo2.AddItem MonthName(5)
    Combo2.AddItem MonthName(6)
    Combo2.AddItem MonthName(7)
    Combo2.AddItem MonthName(8)
    Combo2.AddItem MonthName(9)
    Combo2.AddItem MonthName(10)
    Combo2.AddItem MonthName(11)
    Combo2.AddItem MonthName(12)
    Combo3.AddItem YEAR(Now) + 8
    Combo3.AddItem YEAR(Now) + 7
    Combo3.AddItem YEAR(Now) + 6
    Combo3.AddItem YEAR(Now) + 5
    Combo3.AddItem YEAR(Now) + 4
    Combo3.AddItem YEAR(Now) + 3
    Combo3.AddItem YEAR(Now) + 2
    Combo3.AddItem YEAR(Now) + 1
    Combo3.AddItem YEAR(Now)
    Combo3.AddItem YEAR(Now) - 1
    Combo3.AddItem YEAR(Now) - 2
    Combo3.AddItem YEAR(Now) - 3
    Combo3.AddItem YEAR(Now) - 4
    Combo3.AddItem YEAR(Now) - 5
    Combo3.AddItem YEAR(Now) - 6
    Combo3.AddItem YEAR(Now) - 7
    Combo3.AddItem YEAR(Now) - 8
End Sub

Sub rsrefresh()
    Dim vCUTOFF As String
    Dim vMONTH As Integer
    Dim vYEAR As Integer
    If Combo1.Text = "1st Cut-off" Then
        vCUTOFF = "1"
    ElseIf Combo1.Text = "2nd Cut-off" Then
        vCUTOFF = "2"
    Else
        MsgBox "No Cut-off selected"
        Exit Sub
    End If
    If Combo2.Text <> "" Then
        vMONTH = What_month(Combo2.Text)
    Else
        MsgBox "No month selected"
        Exit Sub
    End If
    If Combo3.Text <> "" Then
        vYEAR = Combo3.Text
    Else
        MsgBox "No year selected"
        Exit Sub
    End If
    
    Set rsPayPeriod = New ADODB.Recordset
    Set rsPayPeriod = gconDMIS.Execute("SELECT * FROM HRMS_PAYROLL WHERE CUT_OFF = '" & vCUTOFF & "' AND PAY_MONTH = '" & vMONTH & "' AND PAY_YEAR = " & vYEAR & "")
    FillGrid
End Sub

Sub FillGrid()
    Grid1.Rows = 1
    If Not rsPayPeriod.EOF And Not rsPayPeriod.BOF Then
        rsPayPeriod.MoveFirst
        While Not rsPayPeriod.EOF
            Grid1.AddItem Null2String(rsPayPeriod!EMPLEVEL) & Chr(9) & Null2String(rsPayPeriod!EMPNO) & Chr(9) & GetEmployeeName(Null2String(rsPayPeriod!EMPNO)) & Chr(9) & N2Str2Zero(rsPayPeriod!Rate) & Chr(9) & N2Str2Zero(rsPayPeriod!OVERTIME) & _
                          Chr(9) & N2Str2Zero(rsPayPeriod!ABSENT) & Chr(9) & N2Str2Zero(rsPayPeriod!UNDERTIME) & Chr(9) & N2Str2Zero(rsPayPeriod!TAXABLEADJ) & Chr(9) & N2Str2Zero(rsPayPeriod!NONTAXABLEADJ) & Chr(9) & N2Str2Zero(rsPayPeriod!Others) & _
                          Chr(9) & N2Str2Zero(rsPayPeriod!PAGIBIG) & Chr(9) & N2Str2Zero(rsPayPeriod!PHILHEALTHE) & Chr(9) & N2Str2Zero(rsPayPeriod!SSSE) & Chr(9) & N2Str2Zero(rsPayPeriod!TAX) & _
                          Chr(9) & N2Str2Zero(rsPayPeriod!SSSSALLOAN) + N2Str2Zero(rsPayPeriod!SSSCALLOAN) + Null2String(rsPayPeriod!PAGSALLOAN) + N2Str2Zero(rsPayPeriod!PAGHDMFLOAN) + N2Str2Zero(rsPayPeriod!OTHERLOAN) & _
                          Chr(9) & N2Str2Zero(rsPayPeriod!ALLOWANCE) & Chr(9) & N2Str2Zero(rsPayPeriod!GROSS) & Chr(9) & N2Str2Zero(rsPayPeriod!NETPAY) + N2Str2Zero(rsPayPeriod!ALLOWANCE) & _
                          Chr(9) & "100" & Chr(9) & N2Str2Zero(rsPayPeriod!PHILHEALTHR) & Chr(9) & N2Str2Zero(rsPayPeriod!SSSR) & Chr(9) & "EDIT"
                          
            rsPayPeriod.MoveNext
        Wend
    End If
End Sub

Function GetEmployeeName(EMPNO As String) As String
    GetEmployeeName = ""
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT LASTNAME, FIRSTNAME FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GetEmployeeName = Null2String(rsTemp!lastname) & ", " & Null2String(rsTemp!FIRSTNAME)
    End If
    Set rsTemp = Nothing
End Function

Private Sub Grid1_DblClick()
    If Grid1.ActiveCell.Col = 22 Then
    Grid1.Enabled = True
    Grid1.Cell(Grid1.ActiveCell.Row, 4).Text = "MAT"
        Grid1.Cell(1, 4).Locked = False
        Grid1.Cell(Grid1.ActiveCell.Row, 5).Locked = False
        Grid1.Cell(Grid1.ActiveCell.Row, 6).Locked = False
        Grid1.Cell(Grid1.ActiveCell.Row, 7).Locked = False
        
        MsgBox Grid1.ActiveCell.Row
    End If
End Sub

Private Sub Grid1_LeaveRow(ByVal Row As Long, Cancel As Boolean)
'    With Grid1
'        .Cell(Row, 3).Text = "3"
'        .Cell(Row, 4).Text = "4"
'        .Cell(Row, 5).Text = "5"
'        .Cell(Row, 6).Text = "6"
'        .Cell(Row, 7).Text = "7"
'        .Cell(Row, 8).Text = "8"
'        .Cell(Row, 9).Text = "9"
'        .Cell(Row, 10).Text = "10"
'        .Cell(Row, 11).Text = "11"
'        .Cell(Row, 12).Text = "12"
'        .Cell(Row, 13).Text = "13"
'    End With
End Sub


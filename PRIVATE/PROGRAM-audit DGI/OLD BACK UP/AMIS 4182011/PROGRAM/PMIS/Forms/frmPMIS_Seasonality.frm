VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmPMISReports_Seasonality 
   Caption         =   "Sesonality Report"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13080
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPMIS_Seasonality.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   13080
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin FlexCell.Grid grd_Smoothing 
      Height          =   6255
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   11033
      BackColor2      =   12648384
      Cols            =   9
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   15
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   345
      Left            =   3030
      TabIndex        =   4
      Top             =   75
      Width           =   1155
   End
   Begin wizProgBar.Prg Prg1 
      Height          =   315
      Left            =   4230
      TabIndex        =   2
      Top             =   90
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   556
      Picture         =   "frmPMIS_Seasonality.frx":058A
      ForeColor       =   0
      BorderStyle     =   2
      BarPicture      =   "frmPMIS_Seasonality.frx":05A6
      BackPictureMode =   0
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
   Begin VB.ComboBox cboDate_Gen 
      Height          =   330
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000000&
      Caption         =   "OK"
      Height          =   345
      Left            =   1890
      TabIndex        =   0
      Top             =   75
      Width           =   1155
   End
   Begin VB.Label LABSTATUS 
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   12270
      TabIndex        =   3
      Top             =   105
      Width           =   1155
   End
End
Attribute VB_Name = "frmPMISReports_Seasonality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Prg1.Value = 0

    Dim RS                                             As ADODB.Recordset
    Dim SMOOTHING_VALUE()                              As Double
    Dim LNR()                                          As Double
    Dim RS_SHP                                         As ADODB.Recordset
    Dim ndate
    Dim j                                              As Integer
    Dim LCOUNT                                         As Long
    Dim T_Q1                                           As Double
    Dim T_Q2                                           As Double
    Dim T_Q4                                           As Double

    grd_Smoothing.AutoRedraw = False
    grd_Smoothing.Rows = 2

    Set RS = New ADODB.Recordset
    Call RS.Open("select  * from PMIS_RANKFLE where TYPE='P' AND date_gen='" & cboDate_Gen & "'    order by TYPE ASC,sales12 desc ", gconDMIS, adOpenStatic, adLockReadOnly)
    LCOUNT = RS.RecordCount
    Prg1.Max = LCOUNT

    j = 1

    For i = 0 To 11
        ndate = DateAdd("m", -i, cboDate_Gen)
        grd_Smoothing.Cell(0, 14 + i).Text = MonthName(Month(ndate), True) & " " & Year(ndate)
    Next

    sdate = cboDate_Gen

    ndate = DateAdd("YYYY", -1, cboDate_Gen)
    grd_Smoothing.Cell(0, 26).Text = MonthName(Month(sdate), True) & " " & Year(sdate) & "-" & MonthName(Month(ndate), True) & " " & Year(ndate)

    sdate = ndate: ndate = DateAdd("YYYY", -1, ndate)
    grd_Smoothing.Cell(0, 30).Text = MonthName(Month(sdate), True) & " " & Year(sdate) & "-" & MonthName(Month(ndate), True) & " " & Year(ndate)

    sdate = ndate: ndate = DateAdd("YYYY", -1, ndate)
    grd_Smoothing.Cell(0, 34).Text = MonthName(Month(sdate), True) & " " & Year(sdate) & "-" & MonthName(Month(ndate), True) & " " & Year(ndate)

    Dim Y1TOTAL                                        As Double
    Dim Y2TOTAL                                        As Double
    Dim Y3TOTAL                                        As Double
    Dim Q1_TOTAL                                       As Double
    Dim Q2_TOTAL                                       As Double
    Dim Q3_TOTAL                                       As Double
    Dim Q4_TOTAL                                       As Double
    Dim YTOTAL                                         As Double
    Dim LVAL()                                         As Double
    While Not RS.EOF
        YTOTAL = 0: Y1TOTAL = 0: Y2TOTAL = 0: Y3TOTAL = 0: Q1_TOTAL = 0: Q2_TOTAL = 0: Q3_TOTAL = 0: Q4_TOTAL = 0
        Prg1.Value = Prg1.Value + 1
        labStatus = (FormatNumber((Prg1.Value / LCOUNT) * 100)) & "%"
        With grd_Smoothing
            j = j + 1
            .AddItem RS!PARTNO, False
            DoEvents
            'DOUBLE EXPONENTIAL SMOOTHING ALPHA=0.3,BETA 0.1
            'SMOOTHING_VALUE = DOUBLE_EXPONENTIAL_SMOOTHING(0.1, 0.3, rs!MONTHS_12, rs!MONTHS_11, rs!MONTHS_10, rs!MONTHS_9, rs!MONTHS_8, rs!MONTHS_7, rs!MONTHS_6, rs!MONTHS_5, rs!MONTHS_4, rs!MONTHS_3, rs!MONTHS_2, rs!prev_month)
            '.Cell(j, 3).Text = SMOOTHING_VALUE(0)    'DOUBLE EXPONENTIAL(0)
            '.Cell(j, 4).Text = SMOOTHING_VALUE(1)    'TREND
            '.Cell(j, 5).Text = SMOOTHING_VALUE(2)    'DOUBLE EXPONENTIAL SMOOTHING WITH TREND REFLECTION
            'DOUBLE EXPONENTIAL SMOOTHING ALPHA=0.5,BETA 0.1
            'SMOOTHING_VALUE = DOUBLE_EXPONENTIAL_SMOOTHING(0.1, 0.5, rs!MONTHS_12, rs!MONTHS_11, rs!MONTHS_10, rs!MONTHS_9, rs!MONTHS_8, rs!MONTHS_7, rs!MONTHS_6, rs!MONTHS_5, rs!MONTHS_4, rs!MONTHS_3, rs!MONTHS_2, rs!prev_month)
            '.Cell(j, 6).Text = SMOOTHING_VALUE(0)    'DOUBLE EXPONENTIAL(0)
            '.Cell(j, 7).Text = SMOOTHING_VALUE(1)    'TREND
            '.Cell(j, 8).Text = SMOOTHING_VALUE(2)    'DOUBLE EXPONENTIAL SMOOTHING WITH TREND REFLECTION
            'LINEAR_REGRESSION
            'LNR = LINEAR_REGRESSION(rs!MONTHS_12, rs!MONTHS_11, rs!MONTHS_10, rs!MONTHS_9, rs!MONTHS_8, rs!MONTHS_7, rs!MONTHS_6, rs!MONTHS_5, rs!MONTHS_4, rs!MONTHS_3, rs!MONTHS_2, rs!prev_month)
            '.Cell(j, 9).Text = LNR(2)                'LINEAR REGRESSION
            '.Cell(j, 10).Text = rs!LR6    'LNR(0)               'SLOPE
            '.Cell(j, 11).Text = LNR(1)               'INTERCEPT
            'cn.Execute ("UPDATE PMIS_RANKFLE SET SlopeLine=" & LNR(0) & ", Intercept0=" & LNR(1) & ", LR6=" & LNR(2) & " WHERE ID=" & rs!id)


            .Cell(j, 2).Text = RS!SALES12
            .Cell(j, 14).Text = RS!Prev_Month
            .Cell(j, 15).Text = RS!Months_2
            .Cell(j, 16).Text = RS!Months_3
            .Cell(j, 17).Text = RS!Months_4
            .Cell(j, 18).Text = RS!Months_5
            .Cell(j, 19).Text = RS!Months_6
            .Cell(j, 20).Text = RS!Months_7
            .Cell(j, 21).Text = RS!Months_8
            .Cell(j, 22).Text = RS!Months_9
            .Cell(j, 23).Text = RS!Months_10
            .Cell(j, 24).Text = RS!Months_11
            .Cell(j, 25).Text = RS!months_12

            Set RS_SHP = gconDMIS.Execute("SELECT " & _
                                        " PREV_MONTH + MONTHS_2 + MONTHS_3 AS Q1" & _
                                        " ,MONTHS_4 + MONTHS_5 + MONTHS_6 AS Q2" & _
                                        " ,MONTHS_7 + MONTHS_8 + MONTHS_9 AS Q3" & _
                                        " ,MONTHS_10 + MONTHS_11 + MONTHS_12 AS Q4" & _
                                        " ,MONTHS_13 + MONTHS_14  + MONTHS_15 AS Q5" & _
                                        " ,MONTHS_16 + MONTHS_17  + MONTHS_18 AS Q6" & _
                                        " ,MONTHS_19 + MONTHS_20  + MONTHS_21 AS Q7" & _
                                        " ,MONTHS_22 + MONTHS_23  + MONTHS_24 AS Q8 " & _
                                        " ,MONTHS_25  + MONTHS_26 + MONTHS_27 AS Q9 " & _
                                        " ,MONTHS_28 + MONTHS_29 + MONTHS_30 AS Q10 " & _
                                        " ,MONTHS_31 + MONTHS_32  + MONTHS_33 AS Q11 " & _
                                        " ,MONTHS_34 + MONTHS_35  + MONTHS_36 AS Q12 From PMIS_SHIPPING WHERE TYPE='P' AND PARTNO='" & RS!PARTNO & "'")

            If Not RS_SHP.EOF Or Not RS_SHP.BOF Then

                Y1TOTAL = RS_SHP!Q12 + RS_SHP!Q11 + RS_SHP!Q10 + RS_SHP!Q9
                Y2TOTAL = RS_SHP!Q8 + RS_SHP!Q7 + RS_SHP!Q6 + RS_SHP!Q5
                Y3TOTAL = RS_SHP!Q4 + RS_SHP!Q3 + RS_SHP!Q2 + RS_SHP!Q1

                T_Q1 = RS_SHP!Q1 + RS_SHP!Q5 + RS_SHP!Q9
                T_Q2 = RS_SHP!Q2 + RS_SHP!Q6 + RS_SHP!Q10
                T_Q3 = RS_SHP!Q3 + RS_SHP!Q7 + RS_SHP!Q11
                T_Q4 = RS_SHP!Q4 + RS_SHP!Q8 + RS_SHP!Q12

                YTOTAL = Y1TOTAL + Y2TOTAL + Y3TOTAL

                LVAL = LINEAR_REGRESSION(Y1TOTAL, Y2TOTAL, Y3TOTAL)


                .Cell(j, 26).Text = RS_SHP!Q1
                .Cell(j, 27).Text = RS_SHP!Q2
                .Cell(j, 28).Text = RS_SHP!Q3
                .Cell(j, 29).Text = RS_SHP!Q4
                .Cell(j, 30).Text = RS_SHP!Q5
                .Cell(j, 31).Text = RS_SHP!Q6
                .Cell(j, 32).Text = RS_SHP!Q7
                .Cell(j, 33).Text = RS_SHP!Q8
                .Cell(j, 34).Text = RS_SHP!Q9
                .Cell(j, 35).Text = RS_SHP!Q10
                .Cell(j, 36).Text = RS_SHP!Q11
                .Cell(j, 37).Text = RS_SHP!Q12

                .Cell(j, 38).Text = LVAL(0)
                .Cell(j, 39).Text = LVAL(1)
                .Cell(j, 40).Text = LVAL(2)

                If YTOTAL > 0 Then
                    .Cell(j, 41).Text = FormatNumber(T_Q1 / YTOTAL)
                Else
                    .Cell(j, 41).Text = 0
                End If

                If YTOTAL > 0 Then
                    .Cell(j, 42).Text = FormatNumber(T_Q2 / YTOTAL)
                Else
                    .Cell(j, 42).Text = 0
                End If


                If YTOTAL > 0 Then
                    .Cell(j, 43).Text = FormatNumber(T_Q3 / YTOTAL)
                Else
                    .Cell(j, 43).Text = 0
                End If

                If YTOTAL > 0 Then
                    .Cell(j, 44).Text = FormatNumber(T_Q3 / YTOTAL)
                Else
                    .Cell(j, 44).Text = 0
                End If

                .Cell(j, 45).Text = Round(.Cell(j, 39).DoubleValue * .Cell(j, 41).DoubleValue, 2)
                .Cell(j, 46).Text = Round(.Cell(j, 39).DoubleValue * .Cell(j, 42).DoubleValue, 2)
                .Cell(j, 47).Text = Round(.Cell(j, 39).DoubleValue * .Cell(j, 43).DoubleValue, 2)
                .Cell(j, 48).Text = Round(.Cell(j, 39).DoubleValue * .Cell(j, 44).DoubleValue, 2)
            End If


        End With
        DoEvents
        RS.MoveNext
    Wend
    Dim rg                                             As Range
    'Set RG = grd_Smoothing.Range(2, 26, j, 29)
    '   RG.BackColor = &H80000000
    'RG.Borders(cellEdgeBottom Or cellEdgeTop Or cellEdgeLeft Or cellEdgeRight) = cellThick

    For i = 14 To 37
        '        Set RG = grd_Smoothing.Range(2, i, j, i)
        '        RG.BackColor = &H80000000
        i = i + 1
    Next

    Set rg = Nothing
    grd_Smoothing.Refresh
    grd_Smoothing.AutoRedraw = True
End Sub

Function DOUBLE_EXPONENTIAL_SMOOTHING(BETA As Double, ALPHA As Double, M1 As Double, M2 As Double, M3 As Double, M4 As Double, M5 As Double, M6 As Double, M7 As Double, M8 As Double, M9 As Double, M10 As Double, M11 As Double, M12 As Double) As Double()
    Dim E1                                             As Double
    Dim E2                                             As Double
    Dim E3                                             As Double
    Dim E4                                             As Double
    Dim E5                                             As Double
    Dim E6                                             As Double
    Dim E7                                             As Double
    Dim E8                                             As Double
    Dim E9                                             As Double
    Dim E10                                            As Double
    Dim E11                                            As Double
    Dim E12                                            As Double

    Dim TREND1                                         As Double
    Dim TREND2                                         As Double
    Dim TREND3                                         As Double
    Dim TREND4                                         As Double
    Dim TREND5                                         As Double
    Dim TREND6                                         As Double
    Dim TREND7                                         As Double
    Dim TREND8                                         As Double
    Dim TREND10                                        As Double
    Dim TREND11                                        As Double
    Dim TREND12                                        As Double
    Dim TREND13                                        As Double

    Dim LVAL(2)                                        As Double
    E1 = M1
    E2 = (ALPHA * M1) + ((1 - ALPHA) * E1)
    E3 = (ALPHA * M2) + ((1 - ALPHA) * E2)
    E4 = (ALPHA * M3) + ((1 - ALPHA) * E3)
    E5 = (ALPHA * M4) + ((1 - ALPHA) * E4)
    E6 = (ALPHA * M5) + ((1 - ALPHA) * E5)
    E7 = (ALPHA * M6) + ((1 - ALPHA) * E6)
    E8 = (ALPHA * M7) + ((1 - ALPHA) * E7)
    E9 = (ALPHA * M8) + ((1 - ALPHA) * E8)
    E10 = (ALPHA * M9) + ((1 - ALPHA) * E9)
    E11 = (ALPHA * M10) + ((1 - ALPHA) * E10)
    E12 = (ALPHA * M11) + ((1 - ALPHA) * E11)
    E13 = (ALPHA * M12) + ((1 - ALPHA) * E12)
    TREND1 = 0
    TREND2 = (BETA * (E2 - E1)) + ((1 - BETA) * TREND1)
    TREND3 = (BETA * (E3 - E2)) + ((1 - BETA) * TREND2)
    TREND4 = (BETA * (E4 - E3)) + ((1 - BETA) * TREND3)
    TREND5 = (BETA * (E5 - E4)) + ((1 - BETA) * TREND4)
    TREND6 = (BETA * (E6 - E5)) + ((1 - BETA) * TREND5)
    TREND7 = (BETA * (E7 - E6)) + ((1 - BETA) * TREND6)
    TREND8 = (BETA * (E8 - E7)) + ((1 - BETA) * TREND7)
    TREND9 = (BETA * (E9 - E8)) + ((1 - BETA) * TREND8)
    TREND10 = (BETA * (E10 - E9)) + ((1 - BETA) * TREND9)
    TREND11 = (BETA * (E11 - E10)) + ((1 - BETA) * TREND10)
    TREND12 = (BETA * (E12 - E11)) + ((1 - BETA) * TREND11)
    TREND13 = (BETA * (E13 - E12)) + ((1 - BETA) * TREND12)


    LVAL(0) = FormatNumber(E13)
    LVAL(1) = FormatNumber(TREND13)
    LVAL(2) = FormatNumber(E13 + TREND13)

    DOUBLE_EXPONENTIAL_SMOOTHING = LVAL

End Function

Private Sub Command2_Click()
    FlexGrid_To_Excel grd_Smoothing, grd_Smoothing.Rows, grd_Smoothing.Cols, 8
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1

    Dim RS                                             As ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT DISTINCT DATE_GEN FROM PMIS_RANKFLE ORDER BY 1 DESC")
    While Not RS.EOF
        cboDate_Gen.AddItem RS!DATE_GEN
        RS.MoveNext
    Wend

    With grd_Smoothing
        .Rows = 2
        .Cols = 49
        .Cell(0, 0).Text = "SN"
        .Column(0).Width = 50
        .Cell(0, 1).Text = "PART NUMBER"
        .Column(1).Width = 90
        .Cell(0, 2).Text = "SALES DEMAND"
        .Column(2).Width = 70
        .Cell(0, 3).Text = "Exponential Smoothing"


        .Cell(1, 3).Text = "E1"
        .Column(3).Width = 0
        .Cell(1, 4).Text = "TREND"
        .Column(4).Width = 0
        .Cell(1, 5).Text = "T1"
        .Column(5).Width = 0

        .Cell(1, 6).Text = "E2"
        .Column(6).Width = 0
        .Cell(1, 7).Text = "TREND"
        .Column(7).Width = 0
        .Cell(1, 8).Text = "T2"
        .Column(8).Width = 0
        .Range(0, 2, 1, 2).Merge
        .Range(0, 2, 1, 2).WrapText = True
        .Range(0, 3, 0, 8).Merge
        .Cell(0, 9).Text = "L/R"
        .Range(0, 9, 1, 9).Merge
        .Column(9).Width = 0
        .Range(0, 9, 1, 9).WrapText = True

        .Cell(0, 10).Text = "Averaging"

        .Cell(1, 10).Text = "M1"
        .Column(10).Width = 0
        .Cell(1, 11).Text = "M2"
        .Column(11).Width = 0
        .Cell(1, 12).Text = "M3"
        .Column(12).Width = 0
        .Cell(1, 13).Text = "M4"
        .Column(13).Width = 0
        .Range(0, 10, 0, 13).Merge
        .Range(0, 10, 0, 13).WrapText = True

        Dim j                                          As Integer


        .Cell(0, 26).Text = Year(LOGDATE)

        .Cell(1, 26).Text = "Q4"
        .Cell(1, 27).Text = "Q3"
        .Cell(1, 28).Text = "Q2"
        .Cell(1, 29).Text = "Q1"
        .Range(0, 26, 0, 29).Merge

        .Cell(0, 30).Text = Year(LOGDATE) - 1
        .Cell(1, 30).Text = "Q4"
        .Cell(1, 31).Text = "Q3"
        .Cell(1, 32).Text = "Q2"
        .Cell(1, 33).Text = "Q1"
        .Range(0, 30, 0, 33).Merge

        .Cell(0, 34).Text = Year(LOGDATE) - 2
        .Cell(1, 34).Text = "Q4"
        .Cell(1, 35).Text = "Q3"
        .Cell(1, 36).Text = "Q2"
        .Cell(1, 37).Text = "Q1"
        .Range(0, 34, 0, 37).Merge
        For i = 25 To 36
            .Column(i + 1).Width = 45
        Next


        For i = 13 To 24
            j = j + 1
            .Cell(0, i + 1).Text = "MNT " & j
            .Column(i + 1).Width = 0
            .Range(0, i + 1, 1, i + 1).Merge
            .Range(0, i + 1, 1, i + 1).WrapText = True
        Next



        .Cell(0, 38).Text = "SEASONAL TREND"
        .Cell(1, 38).Text = "SLOPE"
        .Cell(1, 39).Text = "TREND"
        .Cell(1, 40).Text = "INTERCEPT"
        .Range(0, 38, 0, 40).Merge
        .Cell(0, 41).Text = "SESONAL INDEX"
        .Cell(1, 41).Text = "SQ4"
        .Cell(1, 42).Text = "SQ3"
        .Cell(1, 43).Text = "SQ2"
        .Cell(1, 44).Text = "SQ1"
        .Range(0, 41, 0, 44).Merge
        .Cell(0, 45).Text = "FORECAST"
        .Cell(1, 45).Text = "Q4"
        .Cell(1, 46).Text = "Q3"
        .Cell(1, 47).Text = "Q2"
        .Cell(1, 48).Text = "Q1"
        .Range(0, 45, 0, 48).Merge

        For i = 2 To grd_Smoothing.Cols - 1
            grd_Smoothing.Column(i).Alignment = cellCenterCenter
            'If I > 37 Then
            '    .Column(I).Width = 35

            'End If

        Next

    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = 1 Then Exit Sub
    grd_Smoothing.Width = (Me.ScaleWidth) - 50
    grd_Smoothing.Height = (Me.ScaleHeight) - 480
End Sub












Private Function Sum(ParamArray arguments() As Variant) As Single
    Dim i                                              As Integer
    Dim total                                          As Single

    For i = LBound(arguments) To UBound(arguments)
        total = total + arguments(i)
    Next i
    Sum = total
End Function


Private Function Average(ParamArray arguments() As Variant) As Single
    Dim i                                              As Integer
    Dim total                                          As Single

    For i = LBound(arguments) To UBound(arguments)
        total = total + arguments(i)
    Next i
    Average = total / (UBound(arguments) - LBound(arguments) + 1)
End Function





Public Function LINEAR_REGRESSION(ParamArray Values() As Variant) As Double()
    Dim X                                              As Integer
    Dim y()                                            As Double
    Dim INTLOOP                                        As Integer
    Dim N                                              As Integer

    Dim Q1                                             As Double
    Dim Q2                                             As Double
    Dim Q3                                             As Double
    Dim XY                                             As Double

    Dim XSQUARED                                       As Double
    Dim YSQUARED                                       As Double
    Dim XSUM                                           As Double
    Dim YSUM                                           As Double
    Dim XSQUAREDSUM                                    As Double
    Dim YSQUAREDSUM                                    As Double
    Dim XYSUM                                          As Double
    Dim LVAL(3)                                        As Double
    X = UBound(Values) + 1
    ReDim y(1 To X) As Double
    For INTLOOP = 1 To X
        y(INTLOOP) = Values(INTLOOP - 1)              'Copy values to X
    Next INTLOOP

    For INTLOOP = 1 To X
        XSUM = XSUM + (INTLOOP)
        YSUM = YSUM + y(INTLOOP)
        XSQUAREDSUM = XSQUAREDSUM + (INTLOOP * INTLOOP)
        YSQUAREDSUM = YSQUAREDSUM + (y(INTLOOP) * y(INTLOOP))
        XYSUM = XYSUM + (y(INTLOOP) * INTLOOP)
    Next INTLOOP

    N = X                                             'Number of periods in calculation
    Q1 = (XYSUM - ((XSUM * YSUM) / N))
    Q2 = (XSQUAREDSUM - ((XSUM * XSUM) / N))
    Q3 = (YSQUAREDSUM - ((YSUM * YSUM) / N))
    LVAL(0) = FormatNumber((Q1 / Q2))                 'Slope
    LVAL(1) = FormatNumber((YSUM - LVAL(0) * XSUM) / N)    'Intercept
    LVAL(2) = FormatNumber(((N + 1) * LVAL(0)) + LVAL(1))    'Forecast
    LINEAR_REGRESSION = LVAL
End Function


Public Sub FlexGrid_To_Excel(TheFlexgrid, TheRows As Integer, TheCols As Integer, Optional GridStyle As Integer = 1, Optional WorkSheetName As String)

    Dim objXL                                          As New Excel.Application
    Dim wbXL                                           As New Excel.Workbook
    Dim wsXL                                           As New Excel.Worksheet
    Dim intRow                                         As Integer    ' counter
    Dim intCol                                         As Integer    ' counter

    If Not IsObject(objXL) Then
        MsgBox "You need Microsoft Excel to use this function", _
               vbExclamation, "Print to Excel"
        Exit Sub
    End If

    'On Error Resume Next is necessary because
    'someone may pass more rows
    'or columns than the flexgrid has

    'you can instead check for this,
    'or rewrite the function so that
    'it exports all non-fixed cells
    'to Excel

    On Error Resume Next

    ' open Excel

    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet

    ' name the worksheet
    With wsXL
        If Not WorkSheetName = "" Then
            .Name = WorkSheetName
        End If
    End With

    ' fill worksheet
    For intRow = 1 To TheRows

        For intCol = 1 To TheCols
            With TheFlexgrid
                If wsXL.Columns(intRow).Visible = True Then
                    wsXL.Cells(intRow, intCol).Value = .Cell(intRow - 1, intCol - 1).Text & " "
                End If
            End With
        Next
    Next

    ' format the look
    For intCol = 1 To TheCols
        wsXL.Columns(intCol).AutoFit
        'wsXL.Columns(intCol).AutoFormat (1)
        wsXL.Range("A1", Right(wsXL.Columns(TheCols).AddressLocal, 1) & TheRows).AutoFormat GridStyle
    Next
    objXL.Visible = True
End Sub




'Private Function LINEAR_REGRESSION(SALES_1, SALES_2, SALES_3, SALES_4, SALES_5, SALES_6, SALES_7, SALES_8, SALES_9, SALES_10, SALES_11, SALES_12) As Double()
'    Dim SUMMATIONOFXY                                 As Double
'    Dim SUMMATIONOFX                                  As Double
'    Dim SUMMATIONOFY                                  As Double
'    Dim SUMMATIONOFX2                                 As Double
'    Dim SUMMATIONOFY2                                 As Double
'    Dim MEANOFX                                       As Double
'    Dim MEANOFY                                       As Double
'    Dim XSQR                                          As Double
'    Dim SLOPEOFLINE                                   As Double
'    Dim INTERCEPT                                     As Double
'    Dim i                                             As Integer
'    Dim LVAL(2)                                       As Double
'
'    SUMMATIONOFX = 78
'
'    SUMMATIONOFY = (Val(SALES_1) + Val(SALES_2) + Val(SALES_3) + Val(SALES_4) + Val(SALES_5) + Val(SALES_6) + Val(SALES_7) + Val(SALES_8) + Val(SALES_9) + Val(SALES_10) + Val(SALES_11) + Val(SALES_12))
'
'
'    SUMMATIONOFXY = SALES_1 * 1
'    SUMMATIONOFXY = SUMMATIONOFXY + (SALES_2 * 2)
'    SUMMATIONOFXY = SUMMATIONOFXY + (SALES_3 * 3)
'    SUMMATIONOFXY = SUMMATIONOFXY + (SALES_4 * 4)
'    SUMMATIONOFXY = SUMMATIONOFXY + (SALES_5 * 5)
'    SUMMATIONOFXY = SUMMATIONOFXY + (SALES_6 * 6)
'    SUMMATIONOFXY = SUMMATIONOFXY + (SALES_7 * 7)
'    SUMMATIONOFXY = SUMMATIONOFXY + (SALES_8 * 8)
'    SUMMATIONOFXY = SUMMATIONOFXY + (SALES_9 * 9)
'    SUMMATIONOFXY = SUMMATIONOFXY + (SALES_10 * 10)
'    SUMMATIONOFXY = SUMMATIONOFXY + (SALES_11 * 11)
'    SUMMATIONOFXY = SUMMATIONOFXY + (SALES_12 * 12)
'
'    SUMMATIONOFX2 = 650
'
'    SUMMATIONOFY2 = (Val(SALES_1) ^ 2)
'    SUMMATIONOFY2 = SUMMATIONOFY2 + (Val(SALES_2) ^ 2)
'    SUMMATIONOFY2 = SUMMATIONOFY2 + (Val(SALES_3) ^ 2)
'    SUMMATIONOFY2 = SUMMATIONOFY2 + (Val(SALES_4) ^ 2)
'    SUMMATIONOFY2 = SUMMATIONOFY2 + (Val(SALES_5) ^ 2)
'    SUMMATIONOFY2 = SUMMATIONOFY2 + (Val(SALES_6) ^ 2)
'    SUMMATIONOFY2 = SUMMATIONOFY2 + (Val(SALES_7) ^ 2)
'    SUMMATIONOFY2 = SUMMATIONOFY2 + (Val(SALES_8) ^ 2)
'    SUMMATIONOFY2 = SUMMATIONOFY2 + (Val(SALES_9) ^ 2)
'    SUMMATIONOFY2 = SUMMATIONOFY2 + (Val(SALES_10) ^ 2)
'    SUMMATIONOFY2 = SUMMATIONOFY2 + (Val(SALES_11) ^ 2)
'    SUMMATIONOFY2 = SUMMATIONOFY2 + (Val(SALES_12) ^ 2)
'
'
'    MEANOFX = 6.5
'
'    MEANOFY = SUMMATIONOFY / 12
'    XSQR = 650
'    SLOPEOFLINE = (SUMMATIONOFXY - (12 * MEANOFX * MEANOFY)) / (SUMMATIONOFX2 - (12 * (MEANOFX ^ 2)))
'    INTERCEPT = MEANOFY - (SLOPEOFLINE * MEANOFX)
'    LVAL(0) = FormatNumber(SLOPEOFLINE)
'    LVAL(1) = FormatNumber(INTERCEPT)
'    LVAL(2) = FormatNumber(INTERCEPT + (SLOPEOFLINE * 13))
'
'    LINEAR_REGRESSION = LVAL
'End Function

'Function DOUBLE_EXPONENTIAL_SMOOTHING_TREND(ALPHA As Double, M1 As Double, M2 As Double, M3 As Double, M4 As Double, M5 As Double, M6 As Double, M7 As Double, M8 As Double, M9 As Double, M10 As Double, M11 As Double, M12 As Double) As Double
'    Dim E1                                            As Double
'    Dim E2                                            As Double
'    Dim E3                                            As Double
'    Dim E4                                            As Double
'    Dim E5                                            As Double
'    Dim E6                                            As Double
'    Dim E7                                            As Double
'    Dim E8                                            As Double
'    Dim E9                                            As Double
'    Dim E10                                           As Double
'    Dim E11                                           As Double
'    Dim E12                                           As Double
'    E1 = M1
'    E2 = (ALPHA * M1) + ((1 - ALPHA) * E1)
'    E3 = (ALPHA * M2) + ((1 - ALPHA) * E2)
'    E4 = (ALPHA * M3) + ((1 - ALPHA) * E3)
'    E5 = (ALPHA * M4) + ((1 - ALPHA) * E4)
'    E6 = (ALPHA * M5) + ((1 - ALPHA) * E5)
'    E7 = (ALPHA * M6) + ((1 - ALPHA) * E6)
'    E8 = (ALPHA * M7) + ((1 - ALPHA) * E7)
'    E9 = (ALPHA * M8) + ((1 - ALPHA) * E8)
'    E10 = (ALPHA * M9) + ((1 - ALPHA) * E9)
'    E11 = (ALPHA * M10) + ((1 - ALPHA) * E10)
'    E12 = (ALPHA * M11) + ((1 - ALPHA) * E11)
'    DOUBLE_EXPONENTIAL_SMOOTHING = FormatNumber((ALPHA * M12) + ((1 - ALPHA) * E12))
'End Function

Private Sub grd_Smoothing_AfterUserSort(ByVal Col As Long)

End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmMACTool 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MAC Checking and Fixing Tool"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   Begin wizProgBar.Prg Prg 
      Height          =   345
      Left            =   60
      TabIndex        =   4
      Top             =   6300
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   609
      Picture         =   "MacTool.frx":0000
      ForeColor       =   0
      BarPicture      =   "MacTool.frx":001C
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin VB.CommandButton cmdUpdateMAC 
      Caption         =   "Update Parts with Correct MAC"
      Height          =   465
      Left            =   3090
      TabIndex        =   3
      Top             =   90
      Width           =   2835
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   465
      Left            =   12720
      TabIndex        =   2
      Top             =   90
      Width           =   1725
   End
   Begin MSFlexGridLib.MSFlexGrid grdPartsLedger 
      Height          =   5565
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   9816
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Show Parts with Incorrect MAC"
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2835
   End
End
Attribute VB_Name = "frmMACTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTrans                                                           As ADODB.Recordset

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11

    Dim xlApp                                                         As Excel.Application
    Dim xlBook                                                        As Excel.Workbook
    Dim xlSheet                                                       As Excel.Worksheet

    Set xlApp = CreateObject("Excel.Application")
    
    If Len(Dir(App.Path & "\MACMAC.xlt")) <= 0 Then
        If EXTRACT_FILES(104, "MACMAC.xlt") = False Then
            MsgBox "Please Put MACMAC.xlt on " & vbCrLf & App.Path, vbInformation
            Exit Sub
        End If
    End If
    
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\MACMAC.xlt")
    Set xlSheet = xlBook.Worksheets(1)

    Dim SUM_COMP_MAC, EXT_MAC, EXT_COMP_MAC                           As Double
    SUM_COMP_MAC = 0: EXT_MAC = 0: EXT_COMP_MAC = 0
    Dim rowCtr, xlrCtr                                                As Long
    xlrCtr = 5
    For rowCtr = 1 To grdPartsLedger.Rows - 1
        With grdPartsLedger
            xlSheet.Cells(xlrCtr, "A") = .TextMatrix(rowCtr, 0)
            xlSheet.Cells(xlrCtr, "B") = .TextMatrix(rowCtr, 1)
            xlSheet.Cells(xlrCtr, "C") = .TextMatrix(rowCtr, 2)
            xlSheet.Cells(xlrCtr, "D") = .TextMatrix(rowCtr, 3)
            xlSheet.Cells(xlrCtr, "E") = .TextMatrix(rowCtr, 4)
            xlSheet.Cells(xlrCtr, "F") = .TextMatrix(rowCtr, 5)
            xlSheet.Cells(xlrCtr, "G") = .TextMatrix(rowCtr, 6)
            xlSheet.Cells(xlrCtr, "H") = .TextMatrix(rowCtr, 7)
            xlSheet.Cells(xlrCtr, "I") = .TextMatrix(rowCtr, 8)
            xlSheet.Cells(xlrCtr, "J") = .TextMatrix(rowCtr, 9)
            xlSheet.Cells(xlrCtr, "K") = .TextMatrix(rowCtr, 10)
            xlSheet.Cells(xlrCtr, "L") = .TextMatrix(rowCtr, 11)
            xlSheet.Cells(xlrCtr, "M") = .TextMatrix(rowCtr, 12)
            xlSheet.Cells(xlrCtr, "N") = .TextMatrix(rowCtr, 13)
            xlSheet.Cells(xlrCtr, "O") = .TextMatrix(rowCtr, 14)
            xlSheet.Cells(xlrCtr, "P") = .TextMatrix(rowCtr, 15)
            If NumericVal(.TextMatrix(rowCtr, 3)) > 0 Then
            Else
               xlSheet.Cells(xlrCtr, "Q") = NumericVal(.TextMatrix(rowCtr, 8)) - NumericVal(.TextMatrix(rowCtr, 6))
            End If

            'SUM_COMP_MAC = SUM_COMP_MAC + NumericVal(xlSheet.Cells(xlrCtr, "I"))
            'EXT_MAC = EXT_MAC + NumericVal(xlSheet.Cells(xlrCtr, "J"))
            'EXT_COMP_MAC = EXT_COMP_MAC + NumericVal(xlSheet.Cells(xlrCtr, "K"))
            xlrCtr = xlrCtr + 1
        End With
    Next
    '    xlSheet.Cells(xlrCtr, "I").Font.Bold = True
    '    xlSheet.Cells(xlrCtr, "I").Font.Underline = True
    '    xlSheet.Cells(xlrCtr, "I").Font.Size = 11
    '    xlSheet.Cells(xlrCtr, "I") = Format(SUM_COMP_MAC, MAXIMUM_DIGIT)
    '    xlSheet.Cells(xlrCtr, "J").Font.Bold = True
    '    xlSheet.Cells(xlrCtr, "J").Font.Underline = True
    '    xlSheet.Cells(xlrCtr, "J").Font.Size = 11
    '    xlSheet.Cells(xlrCtr, "J") = Format(EXT_MAC, MAXIMUM_DIGIT)
    '    xlSheet.Cells(xlrCtr, "K").Font.Bold = True
    '    xlSheet.Cells(xlrCtr, "K").Font.Underline = True
    '    xlSheet.Cells(xlrCtr, "K").Font.Size = 11
    '    xlSheet.Cells(xlrCtr, "K") = Format(EXT_COMP_MAC, MAXIMUM_DIGIT)


    xlApp.Visible = True
    Set xlApp = Nothing
    Screen.MousePointer = 0
CloseExcel:
    Set xlApp = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdProcess_Click()
    Screen.MousePointer = 11
    cmdPrint.Enabled = True
    Dim BALANSE                                                       As Integer
    Dim COMPUTED_MAC                                                  As Double
    Dim DistinctTrans                                                 As ADODB.Recordset
    Dim rsStkMas                                                      As ADODB.Recordset
    Dim cnt                                                           As Integer
    gconDMIS.Execute ("Update PMIS_STOCKMAS set VALID_ICC = NULL")
    cleargrid grdPartsLedger
    InitGrid
    Set DistinctTrans = New ADODB.Recordset
    Set DistinctTrans = gconDMIS.Execute("Select DISTINCT STOCK_ORD from PMIS_AllDaytran Order by STOCK_ORD asc")
    If Not DistinctTrans.EOF And Not DistinctTrans.BOF Then
        DistinctTrans.MoveFirst:
        Do While Not DistinctTrans.EOF
            Set rsTrans = New ADODB.Recordset
            Set rsTrans = gconDMIS.Execute("select id,ItemNo,trandate,STOCK_ORD,trantype,tranno,itemno,tranqty,tranucost,mac,status,in_out,TRANUPRICE,usercode from PMIS_AllDaytran where (IN_OUT = 'I' OR IN_OUT = 'O') AND TYPE = 'P' and (STATUS = 'P' OR STATUS = 'B') AND STOCK_ORD = '" & DistinctTrans!STOCK_ORD & "' order by trandate asc, id asc,tranno asc")
            If Not rsTrans.EOF And Not rsTrans.BOF Then
                rsTrans.MoveFirst: COMPUTED_MAC = 0: BALANSE = 0
                Screen.MousePointer = 11
                Do While Not rsTrans.EOF
                    If Null2String(rsTrans!IN_OUT) = "I" Then
                        If BALANSE <= 0 Then
                            COMPUTED_MAC = N2Str2Zero(rsTrans!tranucost)
                        Else
                            COMPUTED_MAC = ((BALANSE * Round(COMPUTED_MAC, 2)) + (N2Str2Zero(rsTrans!tranucost) * N2Str2Zero(rsTrans!tranqty))) / (BALANSE + N2Str2Zero(rsTrans!tranqty))
                        End If
                        BALANSE = BALANSE + N2Str2Zero(rsTrans!tranqty)
                        If Round(COMPUTED_MAC, 0) - Round(N2Str2Zero(rsTrans!Mac), 0) > 0.2 Then
                            gconDMIS.Execute ("Update PMIS_StockMas set VALID_ICC = 'U' where STOCKNO = '" & Null2String(rsTrans!STOCK_ORD) & "'")
                        End If
                    Else
                        BALANSE = BALANSE - N2Str2Zero(rsTrans!tranqty)
                    End If
                    rsTrans.MoveNext
                Loop
                Screen.MousePointer = 0
            End If
            Prg.Value = (DistinctTrans.AbsolutePosition / DistinctTrans.RecordCount) * 100: DoEvents
            Prg.Text = Round((DistinctTrans.AbsolutePosition / DistinctTrans.RecordCount) * 100, 0) & "% Completed"
            DistinctTrans.MoveNext
        Loop
        Set rsStkMas = New ADODB.Recordset
        Set rsStkMas = gconDMIS.Execute("Select STOCKNO from PMIS_STOCKMAS where VALID_ICC = 'U' order by STOCKNO asc")
        If Not rsStkMas.EOF And Not rsStkMas.BOF Then
            rsStkMas.MoveFirst: cnt = 0
            Do While Not rsStkMas.EOF
                cnt = cnt + 1
                grdPartsLedger.AddItem cnt & Chr(9) & Null2String(rsStkMas!STOCKNO)
                Set rsTrans = New ADODB.Recordset
                Set rsTrans = gconDMIS.Execute("select id,ItemNo,trandate,STOCK_ORD,trantype,tranno,itemno,tranqty,tranucost,mac,status,in_out,TRANUPRICE,usercode from PMIS_AllDaytran where (IN_OUT = 'I' OR IN_OUT = 'O') AND TYPE = 'P' and (STATUS = 'P' OR STATUS = 'B') AND STOCK_ORD = '" & rsStkMas!STOCKNO & "' order by trandate asc, id asc,tranno asc")
                If Not rsTrans.EOF And Not rsTrans.BOF Then
                    rsTrans.MoveFirst: BALANSE = 0: COMPUTED_MAC = 0:
                    Do While Not rsTrans.EOF
                        If Null2String(rsTrans!IN_OUT) = "I" Then
                            If BALANSE <= 0 Then
                                COMPUTED_MAC = N2Str2Zero(rsTrans!tranucost)
                            Else
                                COMPUTED_MAC = ((BALANSE * COMPUTED_MAC) + (N2Str2Zero(rsTrans!tranucost) * N2Str2Zero(rsTrans!tranqty))) / (BALANSE + N2Str2Zero(rsTrans!tranqty))
                            End If
                            BALANSE = BALANSE + N2Str2Zero(rsTrans!tranqty)
                        Else
                            BALANSE = BALANSE - N2Str2Zero(rsTrans!tranqty)
                        End If
                        If Null2String(rsTrans!IN_OUT) = "I" Then
                            If (N2Str2Zero(rsTrans!Mac)) = 0 Then
                                If BALANSE > 0 Then
                                    grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & Null2String(rsTrans!trandate) & Chr(9) & _
                                                           Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & _
                                                           N2Str2Zero(rsTrans!tranqty) & Chr(9) & "" & Chr(9) & BALANSE & Chr(9) & _
                                                           ToDoubleNumber(N2Str2Zero(rsTrans!tranucost)) & Chr(9) & _
                                                           ToDoubleNumber(N2Str2Zero(rsTrans!Mac)) & Chr(9) & _
                                                           ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
                                                           Round(BALANSE * N2Str2Zero(rsTrans!Mac), 2) & Chr(9) & _
                                                           Round(BALANSE, 2) * Round(COMPUTED_MAC, 2) & Chr(9) & _
                                                           N2Str2Zero(rsTrans!Mac) - Round(COMPUTED_MAC, 2) & Chr(9) & _
                                                           Round(Round(BALANSE * N2Str2Zero(rsTrans!Mac), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)), 2) & Chr(9) & _
                                                         0 & Chr(9) & _
                                                         0 & Chr(9) & _
                                                           N2Str2Zero(rsTrans!ID)
                                Else
                                    grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & Null2String(rsTrans!trandate) & Chr(9) & _
                                                           Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & _
                                                           N2Str2Zero(rsTrans!tranqty) & Chr(9) & "" & Chr(9) & BALANSE & Chr(9) & _
                                                           ToDoubleNumber(N2Str2Zero(rsTrans!tranucost)) & Chr(9) & _
                                                           ToDoubleNumber(N2Str2Zero(rsTrans!Mac)) & Chr(9) & _
                                                           ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
                                                           Round(BALANSE * N2Str2Zero(rsTrans!Mac), 2) & Chr(9) & _
                                                           Round(BALANSE, 2) * Round(COMPUTED_MAC, 2) & Chr(9) & _
                                                           N2Str2Zero(rsTrans!Mac) - Round(COMPUTED_MAC, 2) & Chr(9) & _
                                                         0 & Chr(9) & _
                                                         0 & Chr(9) & _
                                                         0 & Chr(9) & _
                                                           N2Str2Zero(rsTrans!ID)
                                End If
                            Else
                                If BALANSE > 0 Then
                                    grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & Null2String(rsTrans!trandate) & Chr(9) & _
                                                           Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & _
                                                           N2Str2Zero(rsTrans!tranqty) & Chr(9) & "" & Chr(9) & BALANSE & Chr(9) & _
                                                           ToDoubleNumber(N2Str2Zero(rsTrans!tranucost)) & Chr(9) & _
                                                           ToDoubleNumber(N2Str2Zero(rsTrans!Mac)) & Chr(9) & _
                                                           ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
                                                           Round(BALANSE * N2Str2Zero(rsTrans!Mac), 2) & Chr(9) & _
                                                           Round(BALANSE, 2) * Round(COMPUTED_MAC, 2) & Chr(9) & _
                                                           Round(N2Str2Zero(rsTrans!Mac) - Round(COMPUTED_MAC, 2), 2) & Chr(9) & _
                                                           Round(Round(BALANSE * N2Str2Zero(rsTrans!Mac), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)), 2) & Chr(9) & _
                                                           Round(((N2Str2Zero(rsTrans!Mac) - Round(COMPUTED_MAC, 2)) / N2Str2Zero(rsTrans!Mac)) * 100, 2) & "%" & Chr(9) & _
                                                           Round(((Round(BALANSE * N2Str2Zero(rsTrans!Mac), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2))) / Round(BALANSE * N2Str2Zero(rsTrans!Mac), 2)) * 100, 2) & "%" & Chr(9) & _
                                                           N2Str2Zero(rsTrans!ID)
                                Else
                                    grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & Null2String(rsTrans!trandate) & Chr(9) & _
                                                           Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & _
                                                           N2Str2Zero(rsTrans!tranqty) & Chr(9) & "" & Chr(9) & BALANSE & Chr(9) & _
                                                           ToDoubleNumber(N2Str2Zero(rsTrans!tranucost)) & Chr(9) & _
                                                           ToDoubleNumber(N2Str2Zero(rsTrans!Mac)) & Chr(9) & _
                                                           ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
                                                           Round(BALANSE * N2Str2Zero(rsTrans!Mac), 2) & Chr(9) & _
                                                           Round(BALANSE, 2) * Round(COMPUTED_MAC, 2) & Chr(9) & _
                                                           Round(N2Str2Zero(rsTrans!Mac) - Round(COMPUTED_MAC, 2), 2) & Chr(9) & _
                                                           Round(Round(BALANSE * N2Str2Zero(rsTrans!Mac), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)), 2) & Chr(9) & _
                                                           Round(((N2Str2Zero(rsTrans!Mac) - Round(COMPUTED_MAC, 2)) / N2Str2Zero(rsTrans!Mac)) * 100, 2) & "%" & Chr(9) & _
                                                         0 & "%" & Chr(9) & _
                                                           N2Str2Zero(rsTrans!ID)
                                End If
                            End If
                            gconDMIS.Execute "INSERT INTO PMIS_MACTOOL " & _
                                             "(PARTNO,TRANDATE,TRANTYPE,TRANNO,RECEIVED,ISSUED,BALANCE,UNITCOST,MAC,COMPUTED_MAC,EXT_MAC,EXT_COMP_MAC,DIFF_MAC,DIFF_EXT_MAC,ID) VALUES ('" & Null2String(rsTrans!STOCK_ORD) & "','" & Null2String(rsTrans!trandate) & "'," & _
                                           " '" & Null2String(rsTrans!TranType) & "','" & Null2String(rsTrans!TRANNO) & "'," & N2Str2Zero(rsTrans!tranqty) & ",0," & BALANSE & "," & N2Str2Zero(rsTrans!tranucost) & "," & N2Str2Zero(rsTrans!Mac) & "," & Round(COMPUTED_MAC, 2) & "," & Round(BALANSE & N2Str2Zero(rsTrans!Mac), 2) & "," & Round(Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)) & "," & _
                                           " " & Round(N2Str2Zero(rsTrans!Mac) - Round(COMPUTED_MAC, 2), 2) & "," & Round(Round(BALANSE * N2Str2Zero(rsTrans!Mac), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)), 2) & ", " & N2Str2Zero(rsTrans!ID) & ")"
                        Else
                            If (N2Str2Zero(rsTrans!Mac)) = 0 Then
                                grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & Null2String(rsTrans!trandate) & Chr(9) & _
                                                       Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & "" & Chr(9) & _
                                                       N2Str2Zero(rsTrans!tranqty) & Chr(9) & BALANSE & Chr(9) & _
                                                       ToDoubleNumber(N2Str2Zero(rsTrans!tranucost)) & Chr(9) & _
                                                           ToDoubleNumber(N2Str2Zero(rsTrans!Mac)) & Chr(9) & _
                                                           ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
                                                           "" & Chr(9) & _
                                                           "" & Chr(9) & _
                                                           "" & Chr(9) & _
                                                           "" & Chr(9) & _
                                                           "" & Chr(9) & _
                                                         0 & "%" & Chr(9) & _
                                                           N2Str2Zero(rsTrans!ID)
                            Else
                                grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & Null2String(rsTrans!trandate) & Chr(9) & _
                                                       Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & "" & Chr(9) & _
                                                       N2Str2Zero(rsTrans!tranqty) & Chr(9) & BALANSE & Chr(9) & _
                                                       ToDoubleNumber(N2Str2Zero(rsTrans!tranucost)) & Chr(9) & _
                                                           ToDoubleNumber(N2Str2Zero(rsTrans!Mac)) & Chr(9) & _
                                                           ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
                                                           "" & Chr(9) & _
                                                           "" & Chr(9) & _
                                                           "" & Chr(9) & _
                                                           "" & Chr(9) & _
                                                           "" & Chr(9) & _
                                                         0 & "%" & Chr(9) & _
                                                           N2Str2Zero(rsTrans!ID)
                            End If
                            gconDMIS.Execute "INSERT INTO PMIS_MACTOOL " & _
                                             "(PARTNO,TRANDATE,TRANTYPE,TRANNO,RECEIVED,ISSUED,BALANCE,UNITCOST,MAC,COMPUTED_MAC,EXT_MAC,EXT_COMP_MAC,DIFF_MAC,DIFF_EXT_MAC,ID) VALUES ('" & Null2String(rsTrans!STOCK_ORD) & "','" & Null2String(rsTrans!trandate) & "'," & _
                                           " '" & Null2String(rsTrans!TranType) & "','" & Null2String(rsTrans!TRANNO) & "'," & N2Str2Zero(rsTrans!tranqty) & ",0," & BALANSE & "," & N2Str2Zero(rsTrans!tranucost) & "," & N2Str2Zero(rsTrans!Mac) & "," & Round(COMPUTED_MAC, 2) & ",NULL,NULL," & _
                                           " NULL,NULL, " & N2Str2Zero(rsTrans!ID) & ")"
                            
                        End If
                        rsTrans.MoveNext
                    Loop
                End If
                If cnt = 1 Then grdPartsLedger.RemoveItem 1
                Prg.Value = (rsStkMas.AbsolutePosition / rsStkMas.RecordCount) * 100: DoEvents
                Prg.Text = Round((rsStkMas.AbsolutePosition / rsStkMas.RecordCount) * 100, 0) & "% Completed"
                rsStkMas.MoveNext
            Loop
        End If
    End If
    MsgBox "Show Parts Done."
    Screen.MousePointer = 0
End Sub

Sub UpdateMaster()
    Dim rsStkMas                                       As ADODB.Recordset
    Dim cnt                                            As Long
    Dim BALANSE                                        As Double
    Dim COMPUTED_MAC                                   As Double
    Dim rsTrans                                        As ADODB.Recordset

    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet                                        As Excel.Worksheet

    If Len(Dir(App.Path & "\CHANGES IN MAC.xlt")) <= 0 Then
        If EXTRACT_FILES(102, "CHANGES IN MAC.xlt") = False Then
            MsgBox "Please Put CHANGES IN MAC.xlt on " & vbCrLf & App.Path, vbInformation
            Exit Sub
        End If
    End If

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\CHANGES IN MAC.xlt")
    Set xlSheet = xlBook.Worksheets(1)
    Dim rowCtr, xlrCtr                                 As Long
    xlrCtr = 4

    Set rsStkMas = New ADODB.Recordset
    Set rsStkMas = gconDMIS.Execute("Select STOCKNO,MAC from PMIS_STOCKMAS WHERE VALID_ICC = 'U' order by STOCKNO asc")
    If Not rsStkMas.EOF And Not rsStkMas.BOF Then
        rsStkMas.MoveFirst: cnt = 0
        Do While Not rsStkMas.EOF
            cnt = cnt + 1
            grdPartsLedger.AddItem cnt & Chr(9) & Null2String(rsStkMas!STOCKNO)
            Set rsTrans = New ADODB.Recordset
            Set rsTrans = gconDMIS.Execute("select id,ItemNo,trandate,STOCK_ORD,trantype,tranno,itemno,tranqty,tranucost,mac,status,in_out,TRANUPRICE,usercode from PMIS_AllDaytran where (IN_OUT = 'I' OR IN_OUT = 'O') AND TYPE = 'P' and (STATUS = 'P' OR STATUS = 'B') AND STOCK_ORD = '" & rsStkMas!STOCKNO & "' order by trandate asc, id asc,tranno asc")
            If Not rsTrans.EOF And Not rsTrans.BOF Then
                rsTrans.MoveFirst: BALANSE = 0: COMPUTED_MAC = 0:
                Do While Not rsTrans.EOF
                    If Null2String(rsTrans!IN_OUT) = "I" Then
                        If BALANSE <= 0 Then
                            COMPUTED_MAC = N2Str2Zero(rsTrans!tranucost)
                        Else
                            COMPUTED_MAC = ((BALANSE * COMPUTED_MAC) + (N2Str2Zero(rsTrans!tranucost) * N2Str2Zero(rsTrans!tranqty))) / (BALANSE + N2Str2Zero(rsTrans!tranqty))
                        End If
                        BALANSE = BALANSE + N2Str2Zero(rsTrans!tranqty)
                    Else
                        BALANSE = BALANSE - N2Str2Zero(rsTrans!tranqty)
                    End If
                    rsTrans.MoveNext
                Loop

                xlSheet.Cells(xlrCtr, "A") = rsStkMas!STOCKNO
                xlSheet.Cells(xlrCtr, "B") = BALANSE
                xlSheet.Cells(xlrCtr, "C") = rsStkMas!Mac
                xlSheet.Cells(xlrCtr, "D") = COMPUTED_MAC
                xlSheet.Cells(xlrCtr, "E") = BALANSE * rsStkMas!Mac
                xlSheet.Cells(xlrCtr, "F") = BALANSE * COMPUTED_MAC
                xlSheet.Cells(xlrCtr, "G") = (BALANSE * COMPUTED_MAC) - BALANSE * rsStkMas!Mac
                gconDMIS.Execute ("Update PMIS_StockMas set MAC = " & COMPUTED_MAC & " where STOCKNO = '" & rsStkMas!STOCKNO & "'")
                xlrCtr = xlrCtr + 1
            End If
            If cnt = 1 Then grdPartsLedger.RemoveItem 1
            Prg.Value = (rsStkMas.AbsolutePosition / rsStkMas.RecordCount) * 100: DoEvents
            Prg.Text = Round((rsStkMas.AbsolutePosition / rsStkMas.RecordCount) * 100, 0) & "% Completed"
            rsStkMas.MoveNext
        Loop
    End If

    xlApp.Visible = True
    Set xlApp = Nothing
    Screen.MousePointer = 0

CloseExcel:
    Set xlApp = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdUpdateMAC_Click()
    If MsgBox("Update Transactions History and Parts Master File with Correct MAC?", vbQuestion + vbYesNo, "Please Confirm Action...") = vbYes Then
        If InputBox("Please Input Verification Keyword for Update", "Input Keyword", "") = "ALONE" Then
        Else
            MsgBox "Invalid Verification Keyword!", vbCritical, "Not Allowed!"
            Exit Sub
        End If
        Dim KIM                                                       As Integer
        Dim COMP_MAC                                                  As Double
        Dim YUNIT_KOST                                                As Double
        Dim aydi                                                      As Long

        Dim DB_MAC                                                    As Double


        For KIM = 1 To grdPartsLedger.Rows - 1
            grdPartsLedger.Row = KIM
            grdPartsLedger.Col = 15
            If NumericVal(grdPartsLedger.Text) > 0 Then
                Screen.MousePointer = 11
                grdPartsLedger.Col = 6
                YUNIT_KOST = NumericVal(grdPartsLedger.Text)
                grdPartsLedger.Col = 7
                DB_MAC = NumericVal(grdPartsLedger.Text)
                grdPartsLedger.Col = 8
                COMP_MAC = NumericVal(grdPartsLedger.Text)
                grdPartsLedger.Col = 15
                aydi = grdPartsLedger.Text

                grdPartsLedger.Col = 4
                If NumericVal(grdPartsLedger.Text) > 0 Then
                    gconDMIS.Execute ("Update PMIS_Daytran set TRANUCOST = " & COMP_MAC & ", MAC = " & COMP_MAC & " where ID = " & aydi)
                    gconDMIS.Execute ("Update PMIS_TDaytran set TRANUCOST = " & COMP_MAC & ", MAC = " & COMP_MAC & " where ID = " & aydi)
                Else
                    If YUNIT_KOST <> COMP_MAC Then
                        grdPartsLedger.Col = 2
                        If Left(grdPartsLedger.Text, 3) = "BEG" Or Left(grdPartsLedger.Text, 3) = "ADJ" Then
                            gconDMIS.Execute ("Update PMIS_Daytran set TRANUCOST = " & YUNIT_KOST & ", MAC = " & YUNIT_KOST & " where ID = " & aydi)
                            gconDMIS.Execute ("Update PMIS_TDaytran set TRANUCOST = " & YUNIT_KOST & ", MAC = " & YUNIT_KOST & " where ID = " & aydi)
                        Else
                            gconDMIS.Execute ("Update PMIS_Daytran set TRANUCOST = " & YUNIT_KOST & ", MAC = " & COMP_MAC & " where ID = " & aydi)
                            gconDMIS.Execute ("Update PMIS_TDaytran set TRANUCOST = " & YUNIT_KOST & ", MAC = " & COMP_MAC & " where ID = " & aydi)
                        End If
                    Else
                        gconDMIS.Execute ("Update PMIS_Daytran set MAC = " & COMP_MAC & " where ID = " & aydi)
                        gconDMIS.Execute ("Update PMIS_TDaytran set MAC = " & COMP_MAC & " where ID = " & aydi)
                    End If
                End If
                Screen.MousePointer = 0
            End If
            Prg.Value = (KIM / (grdPartsLedger.Rows - 1)) * 100: DoEvents
            Prg.Text = Round((KIM / (grdPartsLedger.Rows - 1)) * 100, 0) & "% Completed"
        Next
        Call UpdateMaster
        MsgBox "Update for Parts to Correct MAC Successfully Completed!", vbInformation, "Done"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    CenterMe Me, frmMain, 1
    InitGrid
    cmdPrint.Enabled = False
End Sub

Sub InitGrid()
    With grdPartsLedger
        .Row = 0
        .FormatString = "Part Number         | Tran Date   | Tran No         | " & _
                        "Received | Issued | Balance | " & _
                        "Unit Cost |  MAC         | Comp. MAC     |   EXT. MAC | Balance * Comp. MAC |    Diff. MAC | Diff. EXT. MAC| Diff. MAC % | Diff. EXT. MAC % |  ID  "
    End With
End Sub

Private Sub grdPartsLedger_KeyPress(KeyAscii As Integer)
    If grdPartsLedger.Col = 8 Or grdPartsLedger.Col = 6 Then
        If IsNumeric(KeyAscii) = True Then
            grdPartsLedger.CellForeColor = vbWhite
            grdPartsLedger.CellBackColor = vbBlue
            grdPartsLedger.Text = grdPartsLedger.Text & Chr(KeyAscii)
        End If
    End If
End Sub

Private Sub grdPartsLedger_KeyUp(KeyCode As Integer, Shift As Integer)
    If grdPartsLedger.Col = 8 Then
        If KeyCode = vbKeyDelete Then
            grdPartsLedger.Text = ""
        End If
    End If
    If grdPartsLedger.Col = 6 Then
        If KeyCode = vbKeyDelete Then
            grdPartsLedger.Text = ""
        End If
    End If
End Sub

Private Sub grdPartsLedger_LeaveCell()
    If grdPartsLedger.Col = 8 Then
        grdPartsLedger.Text = ToDoubleNumber(NumericVal(grdPartsLedger.Text))
    End If
    If grdPartsLedger.Col = 6 Then
        grdPartsLedger.Text = ToDoubleNumber(NumericVal(grdPartsLedger.Text))
    End If
End Sub

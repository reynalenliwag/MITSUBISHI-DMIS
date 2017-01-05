VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmMACTool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MAC Checking and Fixing Tool"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14520
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MacTool.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   14520
   Begin XtremeReportControl.ReportControl rptRO 
      Height          =   7155
      Left            =   60
      TabIndex        =   5
      Top             =   1110
      Width           =   14445
      _Version        =   655364
      _ExtentX        =   25479
      _ExtentY        =   12621
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnReorder=   0   'False
      MultipleSelection=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   9810
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update Parts with Correct MAC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2910
      TabIndex        =   8
      Top             =   90
      Width           =   2835
   End
   Begin VB.TextBox txtsearch 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   720
      Width           =   5925
   End
   Begin wizProgBar.Prg Prg 
      Height          =   345
      Left            =   60
      TabIndex        =   4
      Top             =   8280
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   609
      Picture         =   "MacTool.frx":1082
      ForeColor       =   0
      BarPicture      =   "MacTool.frx":109E
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6630
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12720
      TabIndex        =   2
      Top             =   90
      Width           =   1725
   End
   Begin MSFlexGridLib.MSFlexGrid grdPartsLedger 
      Height          =   3525
      Left            =   90
      TabIndex        =   1
      Top             =   4320
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   6218
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Show Parts with Incorrect MAC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2835
   End
   Begin VB.Label labStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading. . . Initializing Part Master File, Please wait. . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   6240
      TabIndex        =   9
      Top             =   810
      Visible         =   0   'False
      Width           =   4470
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   705
      Left            =   -30
      TabIndex        =   6
      Top             =   -30
      Width           =   14595
      _Version        =   655364
      _ExtentX        =   25744
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
Attribute VB_Name = "frmMACTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTrans                                            As ADODB.Recordset

Private Sub cmdPrint_Click()
    MessagePop InfoStop, "Action Denied", "Module under Revision"
    Exit Sub
    Screen.MousePointer = 11

    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet                                        As Excel.Worksheet

    Set xlApp = CreateObject("Excel.Application")

    If Len(Dir(App.Path & "\MACMAC.xlt")) <= 0 Then
        If EXTRACT_FILES(104, "MACMAC.xlt") = False Then
            MsgBox "Please Put MACMAC.xlt on " & vbCrLf & App.Path, vbInformation
            Exit Sub
        End If
    End If

    Set xlBook = xlApp.Workbooks.Open(App.Path & "\MACMAC.xlt")
    Set xlSheet = xlBook.Worksheets(1)

    Dim SUM_COMP_MAC, EXT_MAC, EXT_COMP_MAC            As Double
    SUM_COMP_MAC = 0: EXT_MAC = 0: EXT_COMP_MAC = 0
    Dim rowCtr, xlrCtr                                 As Long
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
    'On Error GoTo ErrorCode:
    
    Screen.MousePointer = 11
    txtsearch.Enabled = False
    cmdPrint.Enabled = True
    Timer1.Enabled = True
    
    Dim BALANSE                                        As Integer
    Dim COMPUTED_MAC                                   As Double
    Dim DistinctTrans                                  As ADODB.Recordset
    Dim rsStkMas                                       As ADODB.Recordset
    Dim CNT                                            As Integer
    Dim Value                                          As Double
    Dim rec                                            As XtremeReportControl.ReportRecord
    rptRO.Records.DeleteAll
    rptRO.Populate
    gconDMIS.Execute ("Update PMIS_STOCKMAS set VALID_ICC = NULL")
    Call cleargrid(grdPartsLedger)
    Call InitGrid
    
    Set DistinctTrans = New ADODB.Recordset
    Set DistinctTrans = gconDMIS.Execute("Select DISTINCT STOCK_ORD from PMIS_AllDaytran Order by STOCK_ORD asc")
    If Not DistinctTrans.EOF And Not DistinctTrans.BOF Then
        DistinctTrans.MoveFirst:
        Do While Not DistinctTrans.EOF
            DoEvents
            Set rsTrans = New ADODB.Recordset
            Set rsTrans = gconDMIS.Execute("select id,ItemNo,trandate,STOCK_ORD,trantype,tranno,itemno,tranqty,tranucost,mac,status,in_out,TRANUPRICE,usercode from PMIS_AllDaytran where (IN_OUT = 'I' OR IN_OUT = 'O')  and (STATUS = 'P' OR STATUS = 'B') AND STOCK_ORD = '" & DistinctTrans!STOCK_ORD & "' order by trandate asc, id asc,tranno asc")
            If Not rsTrans.EOF And Not rsTrans.BOF Then
                rsTrans.MoveFirst: COMPUTED_MAC = 0: BALANSE = 0
                Screen.MousePointer = 11
                Do While Not rsTrans.EOF
                    DoEvents
                    If Null2String(rsTrans!IN_OUT) = "I" Then
                        If BALANSE <= 0 Then
                            COMPUTED_MAC = N2Str2Zero(rsTrans!TRANUCOST)
                        Else
                            COMPUTED_MAC = ((BALANSE * Round(COMPUTED_MAC, 2)) + (N2Str2Zero(rsTrans!TRANUCOST) * N2Str2Zero(rsTrans!TRANQTY))) / (BALANSE + N2Str2Zero(rsTrans!TRANQTY))
                        End If
                        DoEvents
                        BALANSE = BALANSE + N2Str2Zero(rsTrans!TRANQTY)
                        If Round(COMPUTED_MAC, 0) - Round(N2Str2Zero(rsTrans!MAC), 0) > 0.2 Then
                            gconDMIS.Execute ("Update PMIS_StockMas set VALID_ICC = 'U' where STOCKNO = '" & Null2String(rsTrans!STOCK_ORD) & "'")
                            Debug.Print Null2String(rsTrans!STOCK_ORD)
                            GoTo NEXT_ITEM:
                        End If
                    Else
                        BALANSE = BALANSE - N2Str2Zero(rsTrans!TRANQTY)
                    End If
                    rsTrans.MoveNext
                Loop
                Screen.MousePointer = 0
            End If
NEXT_ITEM:
            Set rsTrans = Nothing
            Prg.Value = (DistinctTrans.AbsolutePosition / DistinctTrans.RecordCount) * 100: DoEvents
            Prg.Text = Round((DistinctTrans.AbsolutePosition / DistinctTrans.RecordCount) * 100, 0) & "% Completed"
            DistinctTrans.MoveNext
        Loop
        
        Set rsStkMas = New ADODB.Recordset
        Set rsStkMas = gconDMIS.Execute("Select STOCKNO from PMIS_STOCKMAS where VALID_ICC = 'U' order by STOCKNO asc")
        If Not rsStkMas.EOF And Not rsStkMas.BOF Then
            rsStkMas.MoveFirst: CNT = 0
            Do While Not rsStkMas.EOF
                DoEvents
                CNT = CNT + 1
                'grdPartsLedger.AddItem CNT & Chr(9) & Null2String(rsStkMas!STOCKNO)
                
                Set rsTrans = New ADODB.Recordset
                Set rsTrans = gconDMIS.Execute("select id,ItemNo,trandate,STOCK_ORD,trantype,tranno,itemno,tranqty,tranucost,mac,status,in_out,TRANUPRICE,usercode from PMIS_AllDaytran where (IN_OUT = 'I' OR IN_OUT = 'O') AND   (STATUS = 'P' OR STATUS = 'B') AND STOCK_ORD = '" & rsStkMas!STOCKNO & "' order by trandate asc, id asc,tranno asc")
                If Not rsTrans.EOF And Not rsTrans.BOF Then
                    rsTrans.MoveFirst: BALANSE = 0: COMPUTED_MAC = 0:
                    Do While Not rsTrans.EOF
                        DoEvents
                        Value = 0
                        If Null2String(rsTrans!IN_OUT) = "I" Then
                            If BALANSE <= 0 Then
                                COMPUTED_MAC = N2Str2Zero(rsTrans!TRANUCOST)
                            Else
                                COMPUTED_MAC = ((BALANSE * COMPUTED_MAC) + (N2Str2Zero(rsTrans!TRANUCOST) * N2Str2Zero(rsTrans!TRANQTY))) / (BALANSE + N2Str2Zero(rsTrans!TRANQTY))
                            End If
                            BALANSE = BALANSE + N2Str2Zero(rsTrans!TRANQTY)
                        Else
                            BALANSE = BALANSE - N2Str2Zero(rsTrans!TRANQTY)
                        End If
                        
                        If Null2String(rsTrans!IN_OUT) = "I" Then
                            If (N2Str2Zero(rsTrans!MAC)) = 0 Then
                                If BALANSE > 0 Then
                                    Value = Round(Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)), 2)
'                                    grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & _
'                                        Null2String(rsTrans!TRANDATE) & Chr(9) & _
'                                        Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & _
'                                        N2Str2Zero(rsTrans!TRANQTY) & Chr(9) & _
'                                        "" & Chr(9) & _
'                                        BALANSE & Chr(9) & _
'                                        ToDoubleNumber(N2Str2Zero(rsTrans!TRANUCOST)) & Chr(9) & _
'                                        ToDoubleNumber(N2Str2Zero(rsTrans!MAC)) & Chr(9) & _
'                                        ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
'                                        Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) & Chr(9) & _
'                                        Round(BALANSE, 2) * Round(COMPUTED_MAC, 2) & Chr(9) & _
'                                        N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2) & Chr(9) & _
'                                        Round(Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)), 2) & Chr(9) & _
'                                        0 & Chr(9) & _
'                                        0 & Chr(9) & _
'                                        N2Str2Zero(rsTrans!ID)
                                Else
                                    Value = 0
'                                    grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & Null2String(rsTrans!TRANDATE) & Chr(9) & _
'                                        Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & _
'                                        N2Str2Zero(rsTrans!TRANQTY) & Chr(9) & "" & Chr(9) & BALANSE & Chr(9) & _
'                                        ToDoubleNumber(N2Str2Zero(rsTrans!TRANUCOST)) & Chr(9) & _
'                                        ToDoubleNumber(N2Str2Zero(rsTrans!MAC)) & Chr(9) & _
'                                        ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
'                                        Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) & Chr(9) & _
'                                        Round(BALANSE, 2) * Round(COMPUTED_MAC, 2) & Chr(9) & _
'                                        N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2) & Chr(9) & _
'                                        0 & Chr(9) & _
'                                        0 & Chr(9) & _
'                                        0 & Chr(9) & _
'                                        N2Str2Zero(rsTrans!ID)
                                End If
                                
                                Set rec = rptRO.Records.Add
                                rec.AddItem Null2String(rsTrans!STOCK_ORD)
                                rec.AddItem Null2String(rsTrans!TRANDATE)
                                rec.AddItem Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO)
                                rec.AddItem N2Str2Zero(rsTrans!TRANQTY)
                                rec.AddItem Null2String("")
                                rec.AddItem Null2String(BALANSE)
                                rec.AddItem ToDoubleNumber(N2Str2Zero(rsTrans!TRANUCOST))
                                rec.AddItem ToDoubleNumber(N2Str2Zero(rsTrans!MAC))
                                rec.AddItem ToDoubleNumber(COMPUTED_MAC)
                                rec.AddItem Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2)
                                rec.AddItem Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)
                                rec.AddItem N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2)
                                rec.AddItem Null2String(Value)
                                rec.AddItem Null2String(0)
                                rec.AddItem Null2String(0)
                                rec.AddItem N2Str2Zero(rsTrans!ID)
                                
                                'rptRO.Populate
'                                grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & _
'                                    Null2String(rsTrans!TRANDATE) & Chr(9) & _
'                                    Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & _
'                                    N2Str2Zero(rsTrans!TRANQTY) & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    BALANSE & Chr(9) & _
'                                    ToDoubleNumber(N2Str2Zero(rsTrans!TRANUCOST)) & Chr(9) & _
'                                    ToDoubleNumber(N2Str2Zero(rsTrans!MAC)) & Chr(9) & _
'                                    ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
'                                    Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) & Chr(9) & _
'                                    Round(BALANSE, 2) * Round(COMPUTED_MAC, 2) & Chr(9) & _
'                                    N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2) & Chr(9) & _
'                                    VALUE & Chr(9) & _
'                                    0 & Chr(9) & _
'                                    0 & Chr(9) & _
'                                    N2Str2Zero(rsTrans!ID), False
                            Else
                                If BALANSE > 0 Then
                                    Value = Round(((Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2))) / Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2)) * 100, 2)
'                                    grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & Null2String(rsTrans!TRANDATE) & Chr(9) & _
'                                        Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & _
'                                        N2Str2Zero(rsTrans!TRANQTY) & Chr(9) & "" & Chr(9) & BALANSE & Chr(9) & _
'                                        ToDoubleNumber(N2Str2Zero(rsTrans!TRANUCOST)) & Chr(9) & _
'                                        ToDoubleNumber(N2Str2Zero(rsTrans!MAC)) & Chr(9) & _
'                                        ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
'                                        Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) & Chr(9) & _
'                                        Round(BALANSE, 2) * Round(COMPUTED_MAC, 2) & Chr(9) & _
'                                        Round(N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2), 2) & Chr(9) & _
'                                        Round(Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)), 2) & Chr(9) & _
'                                        Round(((N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2)) / N2Str2Zero(rsTrans!MAC)) * 100, 2) & "%" & Chr(9) & _
'                                        Round(((Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2))) / Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2)) * 100, 2) & "%" & Chr(9) & _
'                                        N2Str2Zero(rsTrans!ID)
                                Else
                                    Value = 0
'                                    grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & _
'                                        Null2String(rsTrans!TRANDATE) & Chr(9) & _
'                                        Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & _
'                                        N2Str2Zero(rsTrans!TRANQTY) & Chr(9) & _
'                                        "" & Chr(9) & _
'                                        BALANSE & Chr(9) & _
'                                        ToDoubleNumber(N2Str2Zero(rsTrans!TRANUCOST)) & Chr(9) & _
'                                        ToDoubleNumber(N2Str2Zero(rsTrans!MAC)) & Chr(9) & _
'                                        ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
'                                        Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) & Chr(9) & _
'                                        Round(BALANSE, 2) * Round(COMPUTED_MAC, 2) & Chr(9) & _
'                                        Round(N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2), 2) & Chr(9) & _
'                                        Round(Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)), 2) & Chr(9) & _
'                                        Round(((N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2)) / N2Str2Zero(rsTrans!MAC)) * 100, 2) & "%" & Chr(9) & _
'                                        0 & "%" & Chr(9) & _
'                                        N2Str2Zero(rsTrans!ID)
                                End If
                                
                                Set rec = rptRO.Records.Add
                                rec.AddItem Null2String(rsTrans!STOCK_ORD)
                                rec.AddItem Null2String(rsTrans!TRANDATE)
                                rec.AddItem Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO)
                                rec.AddItem N2Str2Zero(rsTrans!TRANQTY)
                                rec.AddItem Null2String("")
                                rec.AddItem Null2String(BALANSE)
                                rec.AddItem ToDoubleNumber(N2Str2Zero(rsTrans!TRANUCOST))
                                rec.AddItem ToDoubleNumber(N2Str2Zero(rsTrans!MAC))
                                rec.AddItem ToDoubleNumber(COMPUTED_MAC)
                                rec.AddItem Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2)
                                rec.AddItem Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)
                                rec.AddItem Round(N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2), 2)
                                rec.AddItem Round(Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)), 2)
                                rec.AddItem Round(((N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2)) / N2Str2Zero(rsTrans!MAC)) * 100, 2) & "%"
                                rec.AddItem Null2String(Value) & "%"
                                rec.AddItem N2Str2Zero(rsTrans!ID)
                                
                                'rptRO.Populate
'                                grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & _
'                                    Null2String(rsTrans!TRANDATE) & Chr(9) & _
'                                    Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & _
'                                    N2Str2Zero(rsTrans!TRANQTY) & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    BALANSE & Chr(9) & _
'                                    ToDoubleNumber(N2Str2Zero(rsTrans!TRANUCOST)) & Chr(9) & _
'                                    ToDoubleNumber(N2Str2Zero(rsTrans!MAC)) & Chr(9) & _
'                                    ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
'                                    Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) & Chr(9) & _
'                                    Round(BALANSE, 2) * Round(COMPUTED_MAC, 2) & Chr(9) & _
'                                    round(RN2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2), 2) & Chr(9) & _
'                                    Round(Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)), 2) & Chr(9) & _
'                                    Round(((N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2)) / N2Str2Zero(rsTrans!MAC)) * 100, 2) & "%" & Chr(9) & _
'                                    VALUE & "%" & Chr(9) & _
'                                    N2Str2Zero(rsTrans!ID)
                            End If
'                            gconDMIS.Execute "INSERT INTO PMIS_MACTOOL " & _
'                                "(PARTNO,TRANDATE,TRANTYPE,TRANNO,RECEIVED,ISSUED,BALANCE,UNITCOST,MAC,COMPUTED_MAC,EXT_MAC,EXT_COMP_MAC,DIFF_MAC,DIFF_EXT_MAC,ID) VALUES ('" & Null2String(rsTrans!STOCK_ORD) & "','" & Null2String(rsTrans!TRANDATE) & "'," & _
'                                " '" & Null2String(rsTrans!TranType) & _
'                                "', '" & Null2String(rsTrans!TRANNO) & _
'                                "', " & N2Str2Zero(rsTrans!TRANQTY) & _
'                                ", 0 " & _
'                                ", " & BALANSE & _
'                                ", " & N2Str2Zero(rsTrans!TRANUCOST) & _
'                                ", " & N2Str2Zero(rsTrans!MAC) & _
'                                ", " & Round(COMPUTED_MAC, 2) & _
'                                ", " & Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) & _
'                                ", " & Round(Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)) & _
'                                ", " & Round(N2Str2Zero(rsTrans!MAC) - Round(COMPUTED_MAC, 2), 2) & _
'                                ", " & Round(Round(BALANSE * N2Str2Zero(rsTrans!MAC), 2) - (Round(BALANSE, 2) * Round(COMPUTED_MAC, 2)), 2) & ", " & N2Str2Zero(rsTrans!ID) & ")"
                        Else
                            Set rec = rptRO.Records.Add
                            rec.AddItem Null2String(rsTrans!STOCK_ORD)
                            rec.AddItem Null2String(rsTrans!TRANDATE)
                            rec.AddItem Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO)
                            rec.AddItem Null2String("")
                            rec.AddItem N2Str2Zero(rsTrans!TRANQTY)
                            rec.AddItem Null2String(BALANSE)
                            rec.AddItem ToDoubleNumber(N2Str2Zero(rsTrans!TRANUCOST))
                            rec.AddItem ToDoubleNumber(N2Str2Zero(rsTrans!MAC))
                            rec.AddItem ToDoubleNumber(COMPUTED_MAC)
                            rec.AddItem Null2String("")
                            rec.AddItem Null2String("")
                            rec.AddItem Null2String("")
                            rec.AddItem Null2String("")
                            rec.AddItem Null2String("")
                            rec.AddItem Null2String(0 & "%")
                            rec.AddItem N2Str2Zero(rsTrans!ID)
'
                            If (N2Str2Zero(rsTrans!MAC)) = 0 Then
'                                grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & _
'                                    Null2String(rsTrans!TRANDATE) & Chr(9) & _
'                                    Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    N2Str2Zero(rsTrans!TRANQTY) & Chr(9) & _
'                                    BALANSE & Chr(9) & _
'                                    ToDoubleNumber(N2Str2Zero(rsTrans!TRANUCOST)) & Chr(9) & _
'                                    ToDoubleNumber(N2Str2Zero(rsTrans!MAC)) & Chr(9) & _
'                                    ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    0 & "%" & Chr(9) & _
'                                    N2Str2Zero(rsTrans!ID)
                            Else
'                                grdPartsLedger.AddItem Null2String(rsTrans!STOCK_ORD) & Chr(9) & Null2String(rsTrans!TRANDATE) & Chr(9) & _
'                                    Null2String(rsTrans!TranType) & "-" & Null2String(rsTrans!TRANNO) & Chr(9) & "" & Chr(9) & _
'                                    N2Str2Zero(rsTrans!TRANQTY) & Chr(9) & BALANSE & Chr(9) & _
'                                    ToDoubleNumber(N2Str2Zero(rsTrans!TRANUCOST)) & Chr(9) & _
'                                    ToDoubleNumber(N2Str2Zero(rsTrans!MAC)) & Chr(9) & _
'                                    ToDoubleNumber(COMPUTED_MAC) & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    "" & Chr(9) & _
'                                    0 & "%" & Chr(9) & _
'                                    N2Str2Zero(rsTrans!ID)
                            End If

                            
'                            gconDMIS.Execute "INSERT INTO PMIS_MACTOOL " & _
'                                "(PARTNO,TRANDATE,TRANTYPE,TRANNO,RECEIVED,ISSUED,BALANCE,UNITCOST,MAC,COMPUTED_MAC,EXT_MAC,EXT_COMP_MAC,DIFF_MAC,DIFF_EXT_MAC,ID) VALUES ('" & Null2String(rsTrans!STOCK_ORD) & "','" & Null2String(rsTrans!TRANDATE) & "'," & _
'                                " '" & Null2String(rsTrans!TranType) & _
'                                "', '" & Null2String(rsTrans!TRANNO) & _
'                                "', " & N2Str2Zero(rsTrans!TRANQTY) & _
'                                ", 0 " & _
'                                ", " & BALANSE & _
'                                ", " & N2Str2Zero(rsTrans!TRANUCOST) & _
'                                ", " & N2Str2Zero(rsTrans!MAC) & _
'                                ", " & Round(COMPUTED_MAC, 2) & _
'                                ", NULL " & _
'                                ", NULL " & _
'                                ", NULL " & _
'                                ", NULL " & _
'                                ", " & N2Str2Zero(rsTrans!ID) & ")"

                        End If
                        rsTrans.MoveNext
                    Loop
                End If
                
                'If CNT = 1 Then grdPartsLedger.RemoveItem 1
                
                Prg.Value = (rsStkMas.AbsolutePosition / rsStkMas.RecordCount) * 100: DoEvents
                Prg.Text = Round((rsStkMas.AbsolutePosition / rsStkMas.RecordCount) * 100, 0) & "% Completed"
                rsStkMas.MoveNext
            Loop
        End If
    End If
    
    txtsearch.Enabled = True
    rptRO.Populate
    MsgBox "Show Parts Done.", vbInformation, "Information"
    Timer1.Enabled = False
    labStatus.Visible = False
    Screen.MousePointer = 0
    
    Exit Sub
ErrorCode:
    MsgBox err.Description
    Timer1.Enabled = False
    labStatus.Visible = False
    Exit Sub
End Sub

Sub UpdateMaster()
    Dim rsStkMas                                       As ADODB.Recordset
    Dim CNT                                            As Long
    Dim BALANSE                                        As Double
    Dim COMPUTED_MAC                                   As Double
    Dim rsTrans                                        As ADODB.Recordset

    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet                                        As Excel.Worksheet
    Dim DEFAULT_MAC                                    As Double
    Dim rsTMP As New ADODB.Recordset
    
    DEFAULT_MAC = 0

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
        rsStkMas.MoveFirst: CNT = 0
        Do While Not rsStkMas.EOF
            CNT = CNT + 1
            grdPartsLedger.AddItem CNT & Chr(9) & Null2String(rsStkMas!STOCKNO)
            Set rsTrans = New ADODB.Recordset
            Set rsTrans = gconDMIS.Execute("select id,ItemNo,trandate,STOCK_ORD,trantype,tranno,itemno,tranqty,tranucost,mac,status,in_out,TRANUPRICE,usercode from PMIS_AllDaytran where (IN_OUT = 'I' OR IN_OUT = 'O') and (STATUS = 'P' OR STATUS = 'B') AND STOCK_ORD = '" & rsStkMas!STOCKNO & "' order by trandate asc, id asc,tranno asc")
            If Not rsTrans.EOF And Not rsTrans.BOF Then
                rsTrans.MoveFirst: BALANSE = 0: COMPUTED_MAC = 0:
                Do While Not rsTrans.EOF
                    If Null2String(rsTrans!IN_OUT) = "I" Then
                        If BALANSE <= 0 Then
                            COMPUTED_MAC = N2Str2Zero(rsTrans!TRANUCOST)
                        Else
                            COMPUTED_MAC = ((BALANSE * COMPUTED_MAC) + (N2Str2Zero(rsTrans!TRANUCOST) * N2Str2Zero(rsTrans!TRANQTY))) / (BALANSE + N2Str2Zero(rsTrans!TRANQTY))
                        End If
                        BALANSE = BALANSE + N2Str2Zero(rsTrans!TRANQTY)
                    Else
                        BALANSE = BALANSE - N2Str2Zero(rsTrans!TRANQTY)
                    End If
                    rsTrans.MoveNext
                Loop
                
                Set rsTMP = gconDMIS.Execute("SELECT ISNULL(MAC,0) FROM PMIS_DAYTRAN WHERE TRANTYPE = 'BEG' AND STOCK_ORD =  '" & rsStkMas!STOCKNO & "'")
                If Not (rsTMP.EOF And rsTMP.BOF) Then
                
                    DEFAULT_MAC = rsTMP.Fields(0).Value
                Else
                    DEFAULT_MAC = COMPUTED_MAC
                End If

                xlSheet.Cells(xlrCtr, "A") = rsStkMas!STOCKNO
                xlSheet.Cells(xlrCtr, "B") = BALANSE
                xlSheet.Cells(xlrCtr, "C") = rsStkMas!MAC
                xlSheet.Cells(xlrCtr, "D") = COMPUTED_MAC
                xlSheet.Cells(xlrCtr, "E") = BALANSE * rsStkMas!MAC
                xlSheet.Cells(xlrCtr, "F") = BALANSE * COMPUTED_MAC
                xlSheet.Cells(xlrCtr, "G") = (BALANSE * COMPUTED_MAC) - BALANSE * rsStkMas!MAC
                
               
                gconDMIS.Execute ("Update PMIS_StockMas set MAC = " & COMPUTED_MAC & " where STOCKNO = '" & rsStkMas!STOCKNO & "'")
                
                xlrCtr = xlrCtr + 1
            End If
            If CNT = 1 Then grdPartsLedger.RemoveItem 1
            Prg.Value = (rsStkMas.AbsolutePosition / rsStkMas.RecordCount) * 100: DoEvents
            Prg.Text = Round((rsStkMas.AbsolutePosition / rsStkMas.RecordCount) * 100, 0) & "% Completed"
            rsStkMas.MoveNext
        Loop
    End If

    xlApp.Visible = True
    Set xlApp = Nothing
    Screen.MousePointer = 0
    Set rsTMP = Nothing

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
        
        Dim KIM                                        As Integer
        Dim COMP_MAC                                   As Double
        Dim YUNIT_KOST                                 As Double
        Dim aydi                                       As Long
        Dim XTYPE                                      As String
        Dim xTRANTYPE                                  As String
        Dim XSTOCK_ORD                                 As String
        Dim DB_MAC                                     As Double


        For KIM = 1 To grdPartsLedger.Rows - 1
            grdPartsLedger.Row = KIM
            grdPartsLedger.Col = 15
            DoEvents
            If NumericVal(grdPartsLedger.Text) > 0 Then
                Screen.MousePointer = 11
                
                '----------------------------------------------------
                'Update by NVB: 04/20/2010
                'Update Logical bug in MAC tool:
                'Duplicate ID is not Validated
                '-------------------------------------------------------
                grdPartsLedger.Col = 0
                XSTOCK_ORD = grdPartsLedger.Text
                XTYPE = F_GETTYPE(XSTOCK_ORD)
                grdPartsLedger.Col = 2
                'XTYPE = Mid(grdPartsLedger.Text, 3, 1)
                If Mid(grdPartsLedger.Text, 3, 1) = "-" Then
                    xTRANTYPE = Mid(grdPartsLedger.Text, 1, 2)
                Else
                    xTRANTYPE = Mid(grdPartsLedger.Text, 1, 3)
                End If
                grdPartsLedger.Col = 6
                YUNIT_KOST = NumericVal(grdPartsLedger.Text)
                grdPartsLedger.Col = 7
                DB_MAC = NumericVal(grdPartsLedger.Text)
                grdPartsLedger.Col = 8
                COMP_MAC = NumericVal(grdPartsLedger.Text)
                grdPartsLedger.Col = 15
                aydi = grdPartsLedger.Text
            

                grdPartsLedger.Col = 4
                DoEvents
                If NumericVal(grdPartsLedger.Text) > 0 Then
                    gconDMIS.Execute ("Update PMIS_Daytran set " & _
                        " TRANUCOST = " & COMP_MAC & _
                        ", MAC = " & COMP_MAC & _
                        " where STOCK_ORD = '" & XSTOCK_ORD & _
                        "' AND [TYPE] = '" & XTYPE & _
                        "' AND TRANTYPE = '" & xTRANTYPE & _
                        "' AND ID = " & aydi)
                        
                    gconDMIS.Execute ("Update PMIS_TDaytran set " & _
                        " TRANUCOST = " & COMP_MAC & _
                        ", MAC = " & COMP_MAC & _
                        " where STOCK_ORD = '" & XSTOCK_ORD & _
                        "' AND [TYPE] = '" & XTYPE & _
                        "' AND TRANTYPE = '" & xTRANTYPE & _
                        "'  AND ID = " & aydi)
                Else
                    If YUNIT_KOST <> COMP_MAC Then
                        grdPartsLedger.Col = 2
                        If Left(grdPartsLedger.Text, 3) = "BEG" Or Left(grdPartsLedger.Text, 3) = "ADJ" Then
                            gconDMIS.Execute ("Update PMIS_Daytran set " & _
                                " TRANUCOST = " & YUNIT_KOST & _
                                ", MAC = " & YUNIT_KOST & _
                                " where STOCK_ORD = '" & XSTOCK_ORD & _
                                "' AND [TYPE] = '" & XTYPE & _
                                "' AND TRANTYPE = '" & xTRANTYPE & _
                                "' AND ID = " & aydi)
                                
                            gconDMIS.Execute ("Update PMIS_TDaytran set " & _
                                " TRANUCOST = " & YUNIT_KOST & _
                                ", MAC = " & YUNIT_KOST & _
                                " where STOCK_ORD = '" & XSTOCK_ORD & _
                                "' AND [TYPE] = '" & XTYPE & _
                                "' AND TRANTYPE = '" & xTRANTYPE & _
                                "'  AND ID = " & aydi)
                        Else
                            gconDMIS.Execute ("Update PMIS_Daytran set " & _
                                " TRANUCOST = " & YUNIT_KOST & _
                                ", MAC = " & COMP_MAC & _
                                " where STOCK_ORD = '" & XSTOCK_ORD & _
                                "' AND [TYPE] = '" & XTYPE & _
                                "' AND TRANTYPE = '" & xTRANTYPE & _
                                "' AND ID = " & aydi)
                                
                            gconDMIS.Execute ("Update PMIS_TDaytran set " & _
                                " TRANUCOST = " & YUNIT_KOST & _
                                ", MAC = " & COMP_MAC & _
                                " where STOCK_ORD = '" & XSTOCK_ORD & _
                                "' AND [TYPE] = '" & XTYPE & _
                                "' AND TRANTYPE = '" & xTRANTYPE & _
                                "' AND ID = " & aydi)
                        End If
                    Else
                        gconDMIS.Execute ("Update PMIS_Daytran set " & _
                            " MAC = " & COMP_MAC & _
                            " where STOCK_ORD = '" & XSTOCK_ORD & _
                            "' AND [TYPE] = '" & XTYPE & _
                            "' AND TRANTYPE = '" & xTRANTYPE & _
                            "' AND ID = " & aydi)
                            
                        gconDMIS.Execute ("Update PMIS_TDaytran set " & _
                            " MAC = " & COMP_MAC & _
                            " where STOCK_ORD = '" & XSTOCK_ORD & _
                            "' AND [TYPE] = '" & XTYPE & _
                            "' AND TRANTYPE = '" & xTRANTYPE & _
                            "' AND ID = " & aydi)
                    End If
                End If
                Screen.MousePointer = 0
            End If
            
            DoEvents
            Prg.Value = (KIM / (grdPartsLedger.Rows - 1)) * 100: DoEvents
            Prg.Text = Round((KIM / (grdPartsLedger.Rows - 1)) * 100, 0) & "% Completed"
        Next
        
        Call UpdateMaster
        MsgBox "Update for Parts to Correct MAC Successfully Completed!", vbInformation, "Done"
    End If
    Screen.MousePointer = 0
End Sub

Function F_GETTYPE(PARTNUMBER As String) As String
    Dim SQL As String
    Dim rsTMP As New ADODB.Recordset
    
    SQL = "SELECT TYPE FROM PMIS_STOCKMAS WHERE STOCKNO = '" & LTrim(RTrim(PARTNUMBER)) & "'"
    Set rsTMP = gconDMIS.Execute(SQL)
    
    If Not (rsTMP.EOF And rsTMP.BOF) Then
        F_GETTYPE = rsTMP.Fields(0).Value
    End If
    
    Set rsTMP = Nothing
End Function

Private Sub Command1_Click()
    txtsearch.Text = ""
        
    If MsgBox("Update Transactions History and Parts Master File with Correct MAC?", vbQuestion + vbYesNo, "Please Confirm Action...") = vbYes Then
        If InputBox("Please Input Verification Keyword for Update", "Input Keyword", "") <> "ALONE" Then
            MsgBox "Invalid Verification Keyword!", vbCritical, "Not Allowed!"
            Exit Sub
        End If
        
        Dim KIM                                        As Integer
        Dim COMP_MAC                                   As Double
        Dim YUNIT_KOST                                 As Double
        Dim aydi                                       As Long
        Dim XTYPE                                      As String
        Dim xTRANTYPE                                  As String
        Dim XSTOCK_ORD                                 As String
        Dim DB_MAC                                     As Double
    
        Timer1.Enabled = True
        labStatus.Caption = "Updating Transaction Data, Please wait..."
        
        Screen.MousePointer = 11
        For KIM = 0 To rptRO.Rows.Count - 1
            DoEvents
            XSTOCK_ORD = rptRO.Rows(KIM).Record(0).Value
            XTYPE = F_GETTYPE(XSTOCK_ORD)
            
            DoEvents
            Me.Caption = "Part no : " & XSTOCK_ORD & "    Tran no. : " & rptRO.Rows(KIM).Record(2).Value
            If Mid(rptRO.Rows(KIM).Record(2).Value, 3, 1) = "-" Then
                xTRANTYPE = Mid(rptRO.Rows(KIM).Record(2).Value, 1, 2)
            Else
                xTRANTYPE = Mid(rptRO.Rows(KIM).Record(2).Value, 1, 3)
            End If
            YUNIT_KOST = rptRO.Rows(KIM).Record(6).Value
            DB_MAC = NumericVal(rptRO.Rows(KIM).Record(7).Value)
            COMP_MAC = rptRO.Rows(KIM).Record(8).Value
            aydi = rptRO.Rows(KIM).Record(15).Value

            DoEvents
            If rptRO.Rows(KIM).Record(4).Value <> "" Then
                gconDMIS.Execute ("Update PMIS_Daytran set " & _
                    " TRANUCOST = " & COMP_MAC & _
                    ", MAC = " & COMP_MAC & _
                    " where STOCK_ORD = '" & XSTOCK_ORD & _
                    "' AND [TYPE] = '" & XTYPE & _
                    "' AND TRANTYPE = '" & xTRANTYPE & _
                    "' AND ID = " & aydi)
                    
                gconDMIS.Execute ("Update PMIS_TDaytran set " & _
                    " TRANUCOST = " & COMP_MAC & _
                    ", MAC = " & COMP_MAC & _
                    " where STOCK_ORD = '" & XSTOCK_ORD & _
                    "' AND [TYPE] = '" & XTYPE & _
                    "' AND TRANTYPE = '" & xTRANTYPE & _
                    "'  AND ID = " & aydi)
            Else
                If YUNIT_KOST <> COMP_MAC Then
                    If Left(rptRO.Rows(KIM).Record(2).Value, 3) = "BEG" Or Left(rptRO.Rows(KIM).Record(2).Value, 3) = "ADJ" Then
                        gconDMIS.Execute ("Update PMIS_Daytran set " & _
                            " TRANUCOST = " & YUNIT_KOST & _
                            ", MAC = " & YUNIT_KOST & _
                            " where STOCK_ORD = '" & XSTOCK_ORD & _
                            "' AND [TYPE] = '" & XTYPE & _
                            "' AND TRANTYPE = '" & xTRANTYPE & _
                            "' AND ID = " & aydi)
                            
                        gconDMIS.Execute ("Update PMIS_TDaytran set " & _
                            " TRANUCOST = " & YUNIT_KOST & _
                            ", MAC = " & YUNIT_KOST & _
                            " where STOCK_ORD = '" & XSTOCK_ORD & _
                            "' AND [TYPE] = '" & XTYPE & _
                            "' AND TRANTYPE = '" & xTRANTYPE & _
                            "'  AND ID = " & aydi)
                    Else
                        gconDMIS.Execute ("Update PMIS_Daytran set " & _
                            " TRANUCOST = " & YUNIT_KOST & _
                            ", MAC = " & COMP_MAC & _
                            " where STOCK_ORD = '" & XSTOCK_ORD & _
                            "' AND [TYPE] = '" & XTYPE & _
                            "' AND TRANTYPE = '" & xTRANTYPE & _
                            "' AND ID = " & aydi)
                            
                        gconDMIS.Execute ("Update PMIS_TDaytran set " & _
                            " TRANUCOST = " & YUNIT_KOST & _
                            ", MAC = " & COMP_MAC & _
                            " where STOCK_ORD = '" & XSTOCK_ORD & _
                            "' AND [TYPE] = '" & XTYPE & _
                            "' AND TRANTYPE = '" & xTRANTYPE & _
                            "' AND ID = " & aydi)
                    End If
                Else
                    gconDMIS.Execute ("Update PMIS_Daytran set " & _
                        " MAC = " & COMP_MAC & _
                        " where STOCK_ORD = '" & XSTOCK_ORD & _
                        "' AND [TYPE] = '" & XTYPE & _
                        "' AND TRANTYPE = '" & xTRANTYPE & _
                        "' AND ID = " & aydi)
                        
                    gconDMIS.Execute ("Update PMIS_TDaytran set " & _
                        " MAC = " & COMP_MAC & _
                        " where STOCK_ORD = '" & XSTOCK_ORD & _
                        "' AND [TYPE] = '" & XTYPE & _
                        "' AND TRANTYPE = '" & xTRANTYPE & _
                        "' AND ID = " & aydi)
                End If
            End If
            Screen.MousePointer = 0
            
            DoEvents
            Prg.Value = (KIM / (rptRO.Rows.Count)) * 100: DoEvents
            Prg.Text = Round((KIM / (rptRO.Rows.Count)) * 100, 0) & "% Completed"
        Next
        
        Call UpdateMaster
        MsgBox "Update for Parts to Correct MAC Successfully Completed!", vbInformation, "Done"
    End If
    
    DoEvents
    Timer1.Enabled = False
    labStatus.Visible = False
    Me.Caption = "MAC Checking and Fixing Tool"
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    CenterMe Me, frmMain, 1
    Call InitGrid
    cmdPrint.Enabled = False
End Sub

Sub InitGrid()
    With grdPartsLedger
        .Row = 0
        .FormatString = "Part Number         | Tran Date   | Tran No         | " & _
                        "Received | Issued | Balance | " & _
                        "Unit Cost |  MAC         | Comp. MAC     |   EXT. MAC | Balance * Comp. MAC |    Diff. MAC | Diff. EXT. MAC| Diff. MAC % | Diff. EXT. MAC % |  ID  "
    End With
    
    
    With rptRO
      .Columns.DeleteAll
        .Columns.Add 0, "PART NO", 75, True::               .Columns(0).Resizable = True:                  .Columns(0).AllowRemove = False
        .Columns.Add 1, "TRANDATE", 110, True:              .Columns(1).Alignment = xtpAlignmentCenter:     .Columns(1).AllowRemove = False:
        .Columns.Add 2, "TRANNO", 100, True:                .Columns(2).Alignment = xtpAlignmentCenter:     .Columns(2).AllowRemove = False:
        .Columns.Add 3, "RECEIVED", 60, True:               .Columns(3).Alignment = xtpAlignmentCenter:     .Columns(3).AllowRemove = False
        .Columns.Add 4, "ISSUED", 75, True:                 .Columns(4).Alignment = xtpAlignmentCenter:       .Columns(4).AllowRemove = False
        .Columns.Add 5, "BALANCE", 60, True:                .Columns(5).Alignment = xtpAlignmentCenter:     .Columns(5).AllowRemove = False
        .Columns.Add 6, "UNIT COST", 70, True:              .Columns(6).Alignment = xtpAlignmentRight:     .Columns(6).AllowRemove = False
        .Columns.Add 7, "MAC", 60, True:                    .Columns(7).Alignment = xtpAlignmentRight:     .Columns(7).AllowRemove = False
        .Columns.Add 8, "COMP. MAC", 80, True:              .Columns(8).Alignment = xtpAlignmentRight:       .Columns(8).AllowRemove = False
        .Columns.Add 9, "EXT. MAC", 80, True:              .Columns(9).Alignment = xtpAlignmentRight:       .Columns(9).AllowRemove = False
        .Columns.Add 10, "BALANCE * COMP. MAC", 140, True:  .Columns(10).Alignment = xtpAlignmentRight:      .Columns(10).AllowRemove = False
        .Columns.Add 11, "DIFF. MAC", 70, True:             .Columns(11).Alignment = xtpAlignmentRight:      .Columns(11).AllowRemove = False
        .Columns.Add 12, "DIFF. EXT. MAC", 100, True:       .Columns(12).Alignment = xtpAlignmentRight:      .Columns(12).AllowRemove = False
        .Columns.Add 13, "DIFF. MAC %", 90, False:         .Columns(13).Alignment = xtpAlignmentRight:      .Columns(13).AllowRemove = False
        .Columns.Add 14, "DIFF. EXT %", 90, False:        .Columns(14).Alignment = xtpAlignmentRight:      .Columns(14).AllowRemove = False
        .Columns.Add 15, "ID %", 100, False:                  .Columns(15).Alignment = xtpAlignmentRight:      .Columns(15).AllowRemove = False
        
        '.GroupsOrder.Add .Columns(0)
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.TextFont.Weight = 540
        .PaintManager.CaptionFont.Bold = True
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

Private Sub Timer1_Timer()
    If labStatus.Visible = True Then
        labStatus.Visible = False
    Else
        labStatus.Visible = True
    End If
End Sub

Private Sub txtSEARCH_Change()
    rptRO.FilterText = txtsearch.Text
    rptRO.Populate
End Sub

Private Sub txtSearch_GotFocus()
    txtsearch.BackColor = &HC0FFC0
End Sub

Private Sub txtsearch_LinkClose()
    txtsearch.BackColor = vbWhite
End Sub

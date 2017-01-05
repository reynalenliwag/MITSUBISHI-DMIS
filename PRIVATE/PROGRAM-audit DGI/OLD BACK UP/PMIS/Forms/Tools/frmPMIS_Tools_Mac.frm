VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmPMIS_Tools_Mac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mac Issue Finder"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   9540
   Begin XtremeReportControl.ReportControl rptDet 
      Height          =   6045
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9495
      _Version        =   655364
      _ExtentX        =   16748
      _ExtentY        =   10663
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnReorder=   0   'False
      MultipleSelection=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4020
      TabIndex        =   2
      Top             =   90
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmPMIS_Tools_Mac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    Call InitGrid
End Sub

Sub InitGrid()
    With rptDet
        .Columns.DeleteAll
        .Columns.Add 0, "TRAN DATE", 60, True::   .Columns(0).Alignment = xtpAlignmentCenter:       .Columns(0).AllowRemove = False
        .Columns.Add 1, "STOCK ORD", 70, True:        .Columns(1).Alignment = xtpAlignmentCenter:       .Columns(1).AllowRemove = False
        .Columns.Add 2, "UNIT COST", 70, True:          .Columns(2).Alignment = xtpAlignmentCenter:       .Columns(2).AllowRemove = False
        .Columns.Add 3, "MAC.", 50, True:        .Columns(3).Alignment = xtpAlignmentCenter:       .Columns(3).AllowRemove = False
        .Columns.Add 4, "DIFF", 50, True:        .Columns(4).Alignment = xtpAlignmentCenter:       .Columns(4).AllowRemove = False
        
        .GroupsOrder.Add rptDet.Columns(1)
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
    End With

End Sub

Sub FillGrid()
    Dim rsTMP                                           As New ADODB.Recordset
    Dim rec                                             As XtremeReportControl.ReportRecord
    Dim rec2                                            As XtremeReportControl.ReportRecord
    Dim XXX                                             As String
    Dim UNITCOST                                        As Currency
    Dim MAC                                             As Currency
    Dim PREV_MAC                                        As Currency
    Dim PREV_UNITCOST                                   As Currency
    Dim CNT                                             As Long
    Dim PARTNO                                          As String
    Dim PREV_PARTNO                                     As String
    
    CNT = 1
    Screen.MousePointer = 11
    DoEvents
    rptDet.Records.DeleteAll
    
    XXX = ""
    XXX = " SELECT TRANDATE, STOCK_ORD, TRANUCOST, MAC FROM PMIS_AllDayTran "
    XXX = XXX & " WHERE TYPE = 'P' AND TRANTYPE IN ('BEG','RR','ADJ')"
    XXX = XXX & " AND STATUS = 'P' ORDER BY STOCK_ORD, ID"
    
    Set rsTMP = gconDMIS.Execute(XXX)
    rptDet.Records.DeleteAll
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        Do While Not rsTMP.EOF
            Set rec = rptDet.Records.Add
            With rec
                DoEvents
                .AddItem Null2String(rsTMP.Fields(0))
                .AddItem Null2String(rsTMP.Fields(1))
                .AddItem Null2String(rsTMP.Fields(2))
                .AddItem Null2String(rsTMP.Fields(3))
                
                If CNT = 1 Then
                    .AddItem Null2String("")
                    
                    PARTNO = Null2String(rsTMP.Fields(1))
                    MAC = NumericVal(rsTMP.Fields(3))
                    UNITCOST = NumericVal(rsTMP.Fields(2))
                Else
                    If Null2String(rsTMP.Fields(1)) = PREV_PARTNO Then
                        If NumericVal(rsTMP.Fields(2)) > PREV_UNITCOST Then
                            If (NumericVal(rsTMP.Fields(2)) - PREV_UNITCOST) > 100 Then
                                .AddItem Null2String("ERROR")
                            End If
                        ElseIf (NumericVal(rsTMP.Fields(2)) < PREV_UNITCOST) Then
                            If (PREV_UNITCOST - NumericVal(rsTMP.Fields(2))) > 100 Then
                                .AddItem Null2String("ERROR")
                            End If
                        Else
                            .AddItem Null2String("")
                        End If
                    End If
                End If
                
                CNT = CNT + 1
                PREV_PARTNO = Null2String(rsTMP.Fields(1))
                PREV_MAC = NumericVal(rsTMP.Fields(3))
                PREV_UNITCOST = NumericVal(rsTMP.Fields(2))
            End With
            rsTMP.MoveNext
        Loop
    End If
    
    rptDet.Populate
    Set rsTMP = Nothing:
    Screen.MousePointer = 0
End Sub

Private Sub Text1_Change()
    rptDet.FilterText = Text1.Text
    rptDet.Populate
    
End Sub

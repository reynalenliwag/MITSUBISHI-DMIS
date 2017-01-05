VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmcsms_grossprofitreport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service Gross Profit Report"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4830
   Icon            =   "frmgrossprofitreports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1275
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3720
         MouseIcon       =   "frmgrossprofitreports.frx":415B6
         MousePointer    =   99  'Custom
         Picture         =   "frmgrossprofitreports.frx":41708
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Close Window"
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2880
         MouseIcon       =   "frmgrossprofitreports.frx":41B53
         MousePointer    =   99  'Custom
         Picture         =   "frmgrossprofitreports.frx":41CA5
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print Report"
         Top             =   240
         Width           =   795
      End
      Begin VB.ComboBox cboMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Select month from the list"
         Top             =   240
         Width           =   1725
      End
      Begin VB.ComboBox cboYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select year from the list"
         Top             =   720
         Width           =   1725
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport rptService_Advisor 
      Left            =   0
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Service Advisor Performance Report"
      PrintFileLinesPerPage=   60
   End
   Begin wizProgBar.Prg prgExcelGen 
      Height          =   270
      Left            =   80
      TabIndex        =   1
      Top             =   1605
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   476
      Picture         =   "frmgrossprofitreports.frx":42144
      BackColor       =   -2147483629
      ForeColor       =   255
      BorderStyle     =   2
      BarForeColor    =   8454016
      BarPicture      =   "frmgrossprofitreports.frx":42160
      Max             =   200
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
   Begin wizProgBar.Prg prg1 
      Height          =   270
      Left            =   80
      TabIndex        =   8
      Top             =   1320
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   476
      Picture         =   "frmgrossprofitreports.frx":4217C
      BackColor       =   -2147483629
      ForeColor       =   255
      BorderStyle     =   2
      BarForeColor    =   8454016
      BarPicture      =   "frmgrossprofitreports.frx":42198
      Max             =   200
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
End
Attribute VB_Name = "frmcsms_grossprofitreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPOR                                                     As ADODB.Recordset
Dim rsrepordetail                                               As ADODB.Recordset
Dim xlApp                                                       As Excel.Application
Dim xlBook                                                      As Excel.Workbook
Dim xlSheet                                                     As Excel.Worksheet
Dim i                                                           As Double
Dim sqlcommand                                                  As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Call showprint
End Sub


Sub showprint()
    Dim Row                                                     As Integer
    Dim xRO                                                     As String
    Dim XtotalLAbor_billed                                      As Double
    Dim XtotalLAbor_GP                                          As Double
    Dim XtotalLAbor_cost                                        As Double
    Dim YtotalLAbor_billed                                      As Double
    Dim YtotalLAbor_cost                                        As Double
    Dim ictr                                                    As Long
    
    Dim XtotalLAbor_billed_sublet                               As Double
    Dim XtotalLAbor_GP_sublet                                   As Double
    Dim XtotalLAbor_cost_sublet                                 As Double
    Dim YtotalLAbor_billed_sublet                               As Double
    Dim YtotalLAbor_GP_sublet                                   As Double
    Dim YtotalLAbor_cost_sublet                                 As Double
    
    Dim Ytotal_billed_part                                      As Double
    Dim Xtotal_billed_part                                      As Double
    Dim Ytotal_cost_part                                        As Double
    Dim Xtotal_Cost_past                                        As Double
    Dim Xtotal_part_GP                                          As Double
    
    Dim Ytotal_billed_part_sublet                               As Double
    Dim Xtotal_billed_part_sublet                               As Double
    Dim Ytotal_cost_part_sublet                                 As Double
    Dim Xtotal_Cost_past_sublet                                 As Double
    Dim Xtotal_part_GP_sublet                                   As Double
    
    Dim Ytotal_billed_mat                                       As Double
    Dim Xtotal_billed_mat                                       As Double
    Dim Ytotal_cost_mat                                         As Double
    Dim Xtotal_Cost_mat                                         As Double
    Dim Xtotal_mat_GP                                           As Double
    
    Dim Ytotal_billed_mat_sublet                                As Double
    Dim Xtotal_billed_mat_sublet                                As Double
    Dim Ytotal_cost_mat_sublet                                  As Double
    Dim Xtotal_Cost_mat_sublet                                  As Double
    Dim Xtotal_mat_GP_sublet                                    As Double
    
    Dim xLIVILx                                                 As Integer
    Dim xparts_ctr                                              As Integer
    Dim xjobs_ctr                                               As Integer
    Dim xmat_ctr                                                As Integer
    
    Dim xLIVILx_sublet                                          As Integer
    Dim xparts_ctr_sublet                                       As Integer
    Dim xjobs_ctr_sublet                                        As Integer
    Dim xmat_ctr_sublet                                         As Integer
    
    Dim Xgrndtotal_billed_job                                   As Double
    Dim Xgrndtotal_cost_job                                     As Double
    Dim Xgrndtotal_GP_job                                       As Double
    
    Dim Xgrndtotal_billed_part                                  As Double
    Dim Xgrndtotal_cost_part                                    As Double
    Dim Xgrndtotal_GP_part                                      As Double
    
    Dim Xgrndtotal_billed_mat                                   As Double
    Dim Xgrndtotal_cost_mat                                     As Double
    Dim Xgrndtotal_GP_mat                                       As Double
    
    
    
    Xgrndtotal_billed_mat = 0
    Xgrndtotal_cost_mat = 0
    Xgrndtotal_GP_mat = 0
    On Error GoTo IVANEXEQUIELVALENCIA
    If Len(Dir(CSMS_REPORT_PATH & "service_gross_profit_report.xlt")) <= 0 Then
        If EXTRACT_FILES(113, "service_gross_profit_report.xlt") = False Then
            MsgBox "Please Put service_gross_profit_report.xlt on " & vbCrLf & CSMS_REPORT_PATH, vbInformation
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = 11
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "service_gross_profit_report.xlt")
    Set xlSheet = xlBook.Worksheets(1)
    Set rsREPOR = New ADODB.Recordset
    
    sqlcommand = "Select * from csms_repor where Month(dte_rel) = '" & What_month(cboMonth.Text) & "' and year(dte_rel) = '" & cboYear.Text & "' order by dte_rel asc"
    
    Set rsREPOR = gconDMIS.Execute(sqlcommand)
    
    
    If Not (rsREPOR.EOF And rsREPOR.BOF) Then
        
        MousePointer = 11
        prgExcelGen.Max = rsREPOR.RecordCount
        prgExcelGen.Value = 0
        Call enableIT(False)
        xlSheet.Cells(1, "A") = COMPANY_NAME
        xlSheet.Range("A" & 1 & ":" & "A" & 1).Font.Bold = True
        xlSheet.Range("A" & 1 & ":" & "A" & 1).Font.Size = "16"
        xlSheet.Range("A" & 2 & ":" & "A" & 2).Font.Bold = True
        xlSheet.Range("A" & 2 & ":" & "A" & 2).Font.Size = "12"
        xlSheet.Range("A" & 3 & ":" & "A" & 3).Font.Underline = True
        xlSheet.Cells(2, "A") = COMPANY_ADDRESS
        xlSheet.Cells(3, "A") = "SERVICE GROSS PROFIT REPORT FOR : " & cboMonth.Text & " " & cboYear
        Row = 6
        DoEvents
        rsREPOR.MoveFirst
START:
        Do While Not rsREPOR.EOF
            DoEvents
            prgExcelGen.Value = prgExcelGen.Value + 1
            prgExcelGen.Text = "Overall Progress (" & Round(((prgExcelGen.Value / prgExcelGen.Max) * 100), 0) & "%)"
            xlSheet.Range("A" & Row & ":" & "R" & Row).Borders(xlTop).LineStyle = xlContinuous
            xlSheet.Cells(Row, "A") = Trim(rsREPOR!DTE_rel)
            xlSheet.Cells(Row, "B") = setcusname(rsREPOR!ACCT_NO)
            xlSheet.Cells(Row, "C") = Null2String(rsREPOR!REP_OR)
            
            XtotalLAbor_billed = 0
            XtotalLAbor_GP = 0
            XtotalLAbor_cost = 0
            YtotalLAbor_billed = 0
            YtotalLAbor_cost = 0
            
            XtotalLAbor_billed_sublet = 0
            XtotalLAbor_GP_sublet = 0
            XtotalLAbor_cost_sublet = 0
            YtotalLAbor_billed_sublet = 0
            YtotalLAbor_GP_sublet = 0
            YtotalLAbor_cost_sublet = 0
            
            Ytotal_billed_part = 0: Ytotal_billed_part_sublet = 0
            Xtotal_billed_part = 0: Xtotal_billed_part_sublet = 0
            Ytotal_cost_part = 0: Ytotal_cost_part_sublet = 0
            Xtotal_Cost_past = 0: Xtotal_Cost_past_sublet = 0
            Xtotal_part_GP = 0: Xtotal_part_GP_sublet = 0
           
            Ytotal_billed_mat = 0: Ytotal_billed_mat_sublet = 0
            Xtotal_billed_mat = 0: Xtotal_billed_mat_sublet = 0
            Ytotal_cost_mat = 0: Ytotal_cost_mat_sublet = 0
            Xtotal_Cost_mat = 0: Xtotal_Cost_mat_sublet = 0
            Xtotal_mat_GP = 0: Xtotal_mat_GP_sublet = 0
            
            xLIVILx = 0: xLIVILx_sublet = 0
            xparts_ctr = 0: xparts_ctr_sublet = 0
            xjobs_ctr = 0: xjobs_ctr_sublet = 0
            xmat_ctr = 0: xmat_ctr_sublet = 0
            
    
            Set rsrepordetail = New ADODB.Recordset
            'Set rsrepordetail = gconDMIS.Execute("Select case rotype   when 'SR' then 'SR' when 'SR ' then 'SR' when 'SR  ' then 'SR' when 'SR   ' then 'SR' end as rotype,* csms_ro_det where rep_or = '" & rsREPOR!REP_OR & "' and livil not in ('4') order by rotype,livil,line_no asc")
            Set rsrepordetail = gconDMIS.Execute("Select case rotype   when 'SR' then 'SR' when 'SR ' then 'SR' when 'SR  ' then 'SR' when 'SR   ' then 'SR' end as rotype,livil,line_no,DETCDE,wCode,DET_AMT,Discount_2,DetCost,DetCost,detvol,REP_OR,techcode,HRSWRK from csms_ro_det where rep_or = '" & rsREPOR!REP_OR & "' and livil not in ('4') order by line_no asc")
            If Not (rsrepordetail.EOF And rsrepordetail.BOF) Then
                Prg1.Max = rsrepordetail.RecordCount
                Prg1.Value = 0
                DoEvents
                rsrepordetail.MoveFirst
START1:
                    Do While Not rsrepordetail.EOF
                        Prg1.Value = Prg1.Value + 1
                        Prg1.Text = "Processing: ( " & Null2String(rsREPOR!REP_OR) & " )"
                        If Trim(rsrepordetail!ROTYPE) = "SR" Then
                            If rsrepordetail!LIVIL = "1" Then
                                ictr = ictr + 1
                                xjobs_ctr_sublet = xjobs_ctr_sublet + 1
                                xLIVILx_sublet = rsrepordetail!LIVIL
                                
                                xlSheet.Cells(Row, "D") = (rsrepordetail!DETCDE)
                                xlSheet.Range("D" & Row & ":" & "D" & Row).HorizontalAlignment = xlRight
                                
                                If Trim(rsrepordetail!wCode) = "W" Or Trim(rsrepordetail!wCode) = "S" Or Trim(rsrepordetail!wCode) = "C" Then
                                    YtotalLAbor_billed = 0
                                Else
                                    YtotalLAbor_billed = NumericVal(rsrepordetail!DET_AMT - NumericVal(rsrepordetail!Discount_2))
                                End If
                                XtotalLAbor_billed = XtotalLAbor_billed + YtotalLAbor_billed
                                XtotalLAbor_billed_sublet = XtotalLAbor_billed_sublet + YtotalLAbor_billed
                                xlSheet.Cells(Row, "E") = NumericVal(YtotalLAbor_billed)
        
                                'YtotalLAbor_cost = NumericVal(setCost_of_labor(rsrepordetail!REP_OR, rsrepordetail!LINE_NO, Trim(rsrepordetail!TechCode)))
                                YtotalLAbor_cost = Round(NumericVal(rsrepordetail!DetCost * 1.12), 2)
                                XtotalLAbor_cost_sublet = XtotalLAbor_cost_sublet + YtotalLAbor_cost
                                
                                XtotalLAbor_cost = NumericVal(XtotalLAbor_cost + YtotalLAbor_cost)
                                xlSheet.Cells(Row, "F") = NumericVal(YtotalLAbor_cost)
                                
                                XtotalLAbor_GP = ToDoubleNumber(XtotalLAbor_GP + (YtotalLAbor_billed - YtotalLAbor_cost))
                                XtotalLAbor_GP_sublet = ToDoubleNumber(XtotalLAbor_GP_sublet + (YtotalLAbor_billed - YtotalLAbor_cost))
                                
                                xlSheet.Cells(Row, "G") = YtotalLAbor_billed - YtotalLAbor_cost
                                If YtotalLAbor_billed > 0 Then
                                    xlSheet.Cells(Row, "H") = " " & Round(NumericVal((((YtotalLAbor_billed - YtotalLAbor_cost) / NumericVal(rsrepordetail!DET_AMT)) * 100)), 0) & " %"
                                Else
                                    xlSheet.Cells(Row, "H") = "0%"
                                End If
                                Row = Row + 1
                            ElseIf rsrepordetail!LIVIL = "2" Then
                            
                                ictr = ictr + 1
                                xparts_ctr_sublet = xparts_ctr_sublet + 1
                                xLIVILx_sublet = rsrepordetail!LIVIL
                                xlSheet.Cells(Row, "I") = (rsrepordetail!DETCDE)
                                xlSheet.Range("I" & Row & ":" & "I" & Row).HorizontalAlignment = xlRight
                                
                                If Trim(rsrepordetail!wCode) = "W" Or Trim(rsrepordetail!wCode) = "S" Or Trim(rsrepordetail!wCode) = "C" Then
                                    Ytotal_billed_part_sublet = 0
                                Else
                                    Ytotal_billed_part_sublet = NumericVal(rsrepordetail!DET_AMT - NumericVal(rsrepordetail!Discount_2))
                                End If
                                Xtotal_billed_part_sublet = Xtotal_billed_part_sublet + Ytotal_billed_part_sublet
                                Xtotal_billed_part = Xtotal_billed_part + Ytotal_billed_part_sublet
                                xlSheet.Cells(Row, "J") = Ytotal_billed_part_sublet

                                Ytotal_cost_part_sublet = Round(NumericVal((rsrepordetail!DetCost * rsrepordetail!detvol) * 1.12), 2)
                                Xtotal_Cost_past_sublet = Xtotal_Cost_past_sublet + Ytotal_cost_part_sublet
                                
                                Xtotal_Cost_past = Xtotal_Cost_past + Ytotal_cost_part_sublet
                                xlSheet.Cells(Row, "K") = Ytotal_cost_part_sublet
                                
                                Xtotal_part_GP = NumericVal(Xtotal_part_GP + (Ytotal_billed_part_sublet - Ytotal_cost_part_sublet))
                                Xtotal_part_GP_sublet = (Xtotal_part_GP_sublet + (Ytotal_billed_part_sublet - Ytotal_cost_part_sublet))
                                
                                xlSheet.Cells(Row, "L") = (Ytotal_billed_part_sublet - Ytotal_cost_part_sublet)
                                If Ytotal_billed_part_sublet > 0 Then
                                    xlSheet.Cells(Row, "M") = " " & Round(NumericVal(((Ytotal_billed_part_sublet - Ytotal_cost_part_sublet) / Ytotal_billed_part_sublet) * 100), 0) & " %"
                                Else
                                    xlSheet.Cells(Row, "M") = "0%"
                                End If
                                Row = Row + 1
                            ElseIf rsrepordetail!LIVIL = "3" Then
                            
                                ictr = ictr + 1
                                xmat_ctr_sublet = xmat_ctr_sublet + 1
                                xLIVILx_sublet = rsrepordetail!LIVIL
                                xlSheet.Cells(Row, "N") = (rsrepordetail!DETCDE)
                                xlSheet.Range("N" & Row & ":" & "N" & Row).HorizontalAlignment = xlRight
                                If Trim(rsrepordetail!wCode) = "W" Or Trim(rsrepordetail!wCode) = "S" Or Trim(rsrepordetail!wCode) = "C" Then
                                    Ytotal_billed_mat_sublet = 0
                                Else
                                    Ytotal_billed_mat_sublet = NumericVal(rsrepordetail!DET_AMT - NumericVal(rsrepordetail!Discount_2))
                                End If
                                Xtotal_billed_mat_sublet = Xtotal_billed_mat_sublet + Ytotal_billed_mat_sublet
                                Xtotal_billed_mat = Xtotal_billed_mat + Ytotal_billed_mat_sublet
                                xlSheet.Cells(Row, "O") = Ytotal_billed_mat_sublet
                                
                                Ytotal_cost_mat_sublet = Round(NumericVal((rsrepordetail!DetCost * rsrepordetail!detvol) * 1.12), 2)
                                
                                Xtotal_Cost_mat_sublet = Xtotal_Cost_mat_sublet + Ytotal_cost_mat_sublet
                                
                                Xtotal_Cost_mat = Xtotal_Cost_mat + Ytotal_cost_mat_sublet
                                xlSheet.Cells(Row, "P") = Ytotal_cost_mat_sublet
                                
                                xlSheet.Cells(Row, "Q") = (Ytotal_billed_mat_sublet - Ytotal_cost_mat_sublet)
                                Xtotal_mat_GP = NumericVal(Xtotal_mat_GP + (Ytotal_billed_mat_sublet - Ytotal_cost_mat_sublet))
                                Xtotal_mat_GP_sublet = NumericVal(Xtotal_mat_GP_sublet + (Ytotal_billed_mat_sublet - Ytotal_cost_mat_sublet))
                                If Ytotal_billed_mat_sublet > 0 Then
                                    xlSheet.Cells(Row, "R") = " " & Round(NumericVal(((Ytotal_billed_mat_sublet - Ytotal_cost_mat_sublet) / Ytotal_billed_mat_sublet) * 100), 0) & " %"
                                Else
                                    xlSheet.Cells(Row, "R") = "0%"
                                End If
                                Row = Row + 1
                            End If
                            rsrepordetail.MoveNext
                            If rsrepordetail.EOF = True Then
                                
                                If (xjobs_ctr_sublet > xparts_ctr_sublet) And (xjobs_ctr_sublet > xmat_ctr_sublet) Then
                                    Row = Row - xjobs_ctr_sublet
                                    Row = Row + xjobs_ctr_sublet
                                ElseIf (xparts_ctr_sublet > xjobs_ctr_sublet) And (xparts_ctr_sublet > xmat_ctr_sublet) Then
                                    Row = Row - xparts_ctr_sublet
                                    Row = Row + xparts_ctr_sublet
                                ElseIf (xmat_ctr_sublet > xjobs_ctr_sublet) And (xmat_ctr_sublet > xparts_ctr_sublet) Then
                                    Row = Row - xmat_ctr_sublet
                                    Row = Row + xmat_ctr_sublet
                                ElseIf (xjobs_ctr_sublet = xparts_ctr_sublet) And (xjobs_ctr_sublet > xmat_ctr_sublet) Then
                                    Row = Row - xjobs_ctr_sublet
                                    Row = Row + xjobs_ctr_sublet
                                ElseIf (xjobs_ctr_sublet = xmat_ctr_sublet) And (xjobs_ctr_sublet > xparts_ctr_sublet) Then
                                    Row = Row - xjobs_ctr_sublet
                                    Row = Row + xjobs_ctr_sublet
                                ElseIf (xparts_ctr_sublet = xjobs_ctr_sublet) And (xparts_ctr_sublet > xmat_ctr_sublet) Then
                                    Row = Row - xparts_ctr_sublet
                                    Row = Row + xparts_ctr_sublet
                                ElseIf (xparts_ctr_sublet = xmat_ctr_sublet) And (xparts_ctr_sublet > xjobs_ctr_sublet) Then
                                    Row = Row - xparts_ctr_sublet
                                    Row = Row + xparts_ctr_sublet
                                ElseIf (xmat_ctr_sublet = xjobs_ctr_sublet) And (xmat_ctr_sublet > xparts_ctr_sublet) Then
                                    Row = Row - xmat_ctr_sublet
                                    Row = Row + xmat_ctr_sublet
                                ElseIf (xmat_ctr_sublet = xparts_ctr_sublet) And (xmat_ctr_sublet > xjobs_ctr_sublet) Then
                                    Row = Row - xmat_ctr_sublet
                                    Row = Row + xmat_ctr_sublet
                                End If
                                Row = Row + 1
                                
                                'sub total for job sublet
                                xlSheet.Cells(Row, "D") = "Subtotal"
                                xlSheet.Cells(Row, "E") = Round(XtotalLAbor_billed_sublet, 2)
                                xlSheet.Cells(Row, "F") = XtotalLAbor_cost_sublet
                                xlSheet.Cells(Row, "G") = XtotalLAbor_GP_sublet
                                If XtotalLAbor_billed_sublet > 0 Then
                                    xlSheet.Cells(Row, "H") = " " & Round(((XtotalLAbor_GP_sublet / XtotalLAbor_billed_sublet) * 100), 0) & "% "
                                Else
                                    xlSheet.Cells(Row, "H") = "0%"
                                End If
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Font.Bold = True
                                xlSheet.Range("E" & Row & ":" & "G" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Borders(xlBottom).LineStyle = xlContinuous
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Borders(xlTop).LineStyle = xlContinuous
                                
                                'sub total for part sublet
                                xlSheet.Cells(Row, "J") = Round(Xtotal_billed_part_sublet, 2)
                                xlSheet.Cells(Row, "K") = Xtotal_Cost_past_sublet
                                xlSheet.Cells(Row, "L") = Xtotal_part_GP_sublet
                                If Xtotal_billed_part_sublet > 0 Then
                                    xlSheet.Cells(Row, "M") = " " & Round(((Xtotal_part_GP_sublet / Xtotal_billed_part_sublet) * 100), 0) & "% "
                                Else
                                    xlSheet.Cells(Row, "M") = "0%"
                                End If
                                xlSheet.Range("I" & Row & ":" & "M" & Row).Font.Bold = True
                                xlSheet.Range("J" & Row & ":" & "L" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("J" & Row & ":" & "M" & Row).Borders(xlBottom).LineStyle = xlContinuous
                                xlSheet.Range("J" & Row & ":" & "M" & Row).Borders(xlTop).LineStyle = xlContinuous
                                
                                'sub total for material sublet
                                
                                xlSheet.Cells(Row, "O") = Round(Xtotal_billed_mat_sublet, 2)
                                xlSheet.Cells(Row, "P") = Xtotal_Cost_mat_sublet
                                xlSheet.Cells(Row, "Q") = Xtotal_mat_GP_sublet
                                If Xtotal_billed_mat_sublet > 0 Then
                                    xlSheet.Cells(Row, "R") = " " & Round(((Xtotal_mat_GP_sublet / Xtotal_billed_mat_sublet) * 100), 0) & "% "
                                Else
                                    xlSheet.Cells(Row, "R") = "0%"
                                End If
                                xlSheet.Range("N" & Row & ":" & "R" & Row).Font.Bold = True
                                xlSheet.Range("O" & Row & ":" & "Q" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("O" & Row & ":" & "R" & Row).Borders(xlBottom).LineStyle = xlContinuous
                                xlSheet.Range("O" & Row & ":" & "R" & Row).Borders(xlTop).LineStyle = xlContinuous
                                
                                'total RO PART
                                Row = Row + 2
                                
                                xlSheet.Cells(Row, "D") = "TOTAl RO"
                                xlSheet.Cells(Row, "E") = Round(XtotalLAbor_billed, 2)
                                'xlSheet.Cells(Row, "E") = Round((XtotalLAbor_billed - NumericVal(rsREPOR!PartLabor)), 2)
                                xlSheet.Cells(Row, "F") = XtotalLAbor_cost
                                xlSheet.Cells(Row, "G") = XtotalLAbor_GP
                                
                                '---------------------------------------------------------------------
                                'Xgrndtotal_billed_job = Xgrndtotal_billed_job + Round((XtotalLAbor_billed - NumericVal(rsREPOR!PartLabor)), 2)
                                Xgrndtotal_billed_job = Xgrndtotal_billed_job + Round(XtotalLAbor_billed, 2)
                                Xgrndtotal_cost_job = Xgrndtotal_cost_job + XtotalLAbor_cost
                                Xgrndtotal_GP_job = Xgrndtotal_GP_job + XtotalLAbor_GP
                                '---------------------------------------------------------------------
                                
                                If Round(XtotalLAbor_billed, 2) > 0 Then
                                'If Round((XtotalLAbor_billed - NumericVal(rsREPOR!PartLabor)), 2) > 0 Then
                                    'xlSheet.Cells(Row, "H") = " " & Round(((XtotalLAbor_GP / Round((XtotalLAbor_billed - NumericVal(rsREPOR!PartLabor)), 2)) * 100), 0) & "% "
                                    xlSheet.Cells(Row, "H") = " " & Round(((XtotalLAbor_GP / XtotalLAbor_billed) * 100), 0) & "% "
                                Else
                                    xlSheet.Cells(Row, "H") = "0%"
                                End If
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Font.Bold = True
                                xlSheet.Range("E" & Row & ":" & "G" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Borders(xlBottom).LineStyle = xlDouble
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Borders(xlTop).LineStyle = xlDouble
                                
                                'total RO PARTS

                                'xlSheet.Cells(Row, "J") = Round((Xtotal_billed_part - NumericVal(rsREPOR!PartParts)), 2)
                                xlSheet.Cells(Row, "J") = Round(Xtotal_billed_part, 2)
                                xlSheet.Cells(Row, "K") = Xtotal_Cost_past
                                xlSheet.Cells(Row, "L") = Xtotal_part_GP
                                
                                '---------------------------------------------------------------------
                                'Xgrndtotal_billed_part = Xgrndtotal_billed_part + Round((Xtotal_billed_part - NumericVal(rsREPOR!PartParts)), 2)
                                Xgrndtotal_billed_part = Xgrndtotal_billed_part + Round(Xtotal_billed_part, 2)
                                Xgrndtotal_cost_part = Xgrndtotal_cost_part + Xtotal_Cost_past
                                Xgrndtotal_GP_part = Xgrndtotal_GP_part + Xtotal_part_GP
                                '---------------------------------------------------------------------
                                
                                If Round(Xtotal_billed_part, 2) > 0 Then
                                'If Round((Xtotal_billed_part - NumericVal(rsREPOR!PartParts)), 2) > 0 Then
                                    xlSheet.Cells(Row, "M") = " " & Round(((Xtotal_part_GP / Xtotal_billed_part) * 100), 0) & "% "
                                    'xlSheet.Cells(Row, "M") = " " & Round(((Xtotal_part_GP / Round((Xtotal_billed_part - NumericVal(rsREPOR!PartParts)), 2)) * 100), 0) & "% "
                                Else
                                    xlSheet.Cells(Row, "M") = "0%"
                                End If
                                xlSheet.Range("I" & Row & ":" & "M" & Row).Font.Bold = True
                                xlSheet.Range("J" & Row & ":" & "L" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("I" & Row & ":" & "M" & Row).Borders(xlBottom).LineStyle = xlDouble
                                xlSheet.Range("I" & Row & ":" & "M" & Row).Borders(xlTop).LineStyle = xlDouble
                                
                                'total RO MAterial
                              
                                'xlSheet.Cells(Row, "O") = Round((Xtotal_billed_mat - NumericVal(rsREPOR!PartMaterials)), 2)
                                xlSheet.Cells(Row, "O") = Round(Xtotal_billed_mat, 2)
                                xlSheet.Cells(Row, "P") = Xtotal_Cost_mat
                                xlSheet.Cells(Row, "Q") = Xtotal_mat_GP
                                
                                '---------------------------------------------------------------------
                                'Xgrndtotal_billed_mat = Xgrndtotal_billed_mat + Round((Xtotal_billed_mat - NumericVal(rsREPOR!PartMaterials)), 2)
                                Xgrndtotal_billed_mat = Xgrndtotal_billed_mat + Round(Xtotal_billed_mat, 2)
                                Xgrndtotal_cost_mat = Xgrndtotal_cost_mat + Xtotal_Cost_mat
                                Xgrndtotal_GP_mat = Xgrndtotal_GP_mat + Xtotal_mat_GP
                                '---------------------------------------------------------------------
                                'If Round((Xtotal_billed_mat - NumericVal(rsREPOR!PartMaterials)), 2) > 0 Then
                                If Round(Xtotal_billed_mat, 2) > 0 Then
                                    'xlSheet.Cells(Row, "R") = " " & Round(((Xtotal_mat_GP / Round((Xtotal_billed_mat - NumericVal(rsREPOR!PartMaterials)), 2)) * 100), 0) & "% "
                                    xlSheet.Cells(Row, "R") = " " & Round(((Xtotal_mat_GP / Xtotal_billed_mat) * 100), 0) & "% "
                                Else
                                    xlSheet.Cells(Row, "R") = "0%"
                                End If
                                xlSheet.Range("N" & Row & ":" & "R" & Row).Font.Bold = True
                                xlSheet.Range("O" & Row & ":" & "Q" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("N" & Row & ":" & "R" & Row).Borders(xlBottom).LineStyle = xlDouble
                                xlSheet.Range("N" & Row & ":" & "R" & Row).Borders(xlTop).LineStyle = xlDouble
                                
                                Row = Row + 2
                                ictr = 0
                                rsREPOR.MoveNext
                                GoTo START
                            Else
                                If rsrepordetail!LIVIL > xLIVILx_sublet Then
                                    Row = Row - ictr
                                    ictr = 0
                                End If
                                GoTo START1
                            End If
                        Else
                            If rsrepordetail!LIVIL = "1" Then
                    
                                ictr = ictr + 1
                                xjobs_ctr = xjobs_ctr + 1
                                xLIVILx = rsrepordetail!LIVIL
                                
                                xlSheet.Cells(Row, "D") = (rsrepordetail!DETCDE)
                                xlSheet.Range("D" & Row & ":" & "D" & Row).HorizontalAlignment = xlRight
                                
                                If Trim(rsrepordetail!wCode) = "W" Or Trim(rsrepordetail!wCode) = "S" Or Trim(rsrepordetail!wCode) = "C" Then
                                    YtotalLAbor_billed = 0
                                Else
                                    YtotalLAbor_billed = NumericVal(rsrepordetail!DET_AMT - NumericVal(rsrepordetail!Discount_2))
                                End If
                                XtotalLAbor_billed = XtotalLAbor_billed + YtotalLAbor_billed
                                xlSheet.Cells(Row, "E") = NumericVal(YtotalLAbor_billed)
        
                                'YtotalLAbor_cost = NumericVal(setCost_of_labor(rsrepordetail!REP_OR, rsrepordetail!LINE_NO, Trim(rsrepordetail!TechCode)))
                                YtotalLAbor_cost = Round(NumericVal(setCost_of_labor(rsrepordetail!REP_OR, rsrepordetail!LINE_NO, Trim(rsrepordetail!TechCode))), 2)
                                
                                'YtotalLAbor_cost = NumericVal(rsrepordetail!HRSWRK * getrateperhour(rsrepordetail!TechCode))
                                XtotalLAbor_cost = NumericVal(XtotalLAbor_cost + YtotalLAbor_cost)
                                
                                xlSheet.Cells(Row, "F") = NumericVal(YtotalLAbor_cost)
                                XtotalLAbor_GP = ToDoubleNumber(XtotalLAbor_GP + (YtotalLAbor_billed - YtotalLAbor_cost))
                                xlSheet.Cells(Row, "G") = YtotalLAbor_billed - YtotalLAbor_cost
                                If YtotalLAbor_billed > 0 Then
                                    xlSheet.Cells(Row, "H") = " " & Round(NumericVal((((YtotalLAbor_billed - YtotalLAbor_cost) / NumericVal(rsrepordetail!DET_AMT)) * 100)), 0) & " %"
                                Else
                                    xlSheet.Cells(Row, "H") = "0%"
                                End If
                                Row = Row + 1
                                
                            ElseIf rsrepordetail!LIVIL = "2" Then
                                
                                ictr = ictr + 1
                                xparts_ctr = xparts_ctr + 1
                                xLIVILx = rsrepordetail!LIVIL
                                
                                xlSheet.Cells(Row, "I") = (rsrepordetail!DETCDE)
                                xlSheet.Range("I" & Row & ":" & "I" & Row).HorizontalAlignment = xlRight
                                If Trim(rsrepordetail!wCode) = "W" Or Trim(rsrepordetail!wCode) = "S" Or Trim(rsrepordetail!wCode) = "C" Then
                                    Ytotal_billed_part = 0
                                Else
                                    Ytotal_billed_part = NumericVal(rsrepordetail!DET_AMT - rsrepordetail!Discount_2)
                                End If
                                'sum
                                Xtotal_billed_part = Xtotal_billed_part + Ytotal_billed_part
                                xlSheet.Cells(Row, "J") = Ytotal_billed_part
                                
                                Ytotal_cost_part = NumericVal(rsrepordetail!DetCost * rsrepordetail!detvol)
                                Xtotal_Cost_past = Xtotal_Cost_past + Ytotal_cost_part
                                xlSheet.Cells(Row, "K") = Ytotal_cost_part
                                
                                xlSheet.Cells(Row, "L") = (Ytotal_billed_part - Ytotal_cost_part)
                                Xtotal_part_GP = NumericVal(Xtotal_part_GP + (Ytotal_billed_part - Ytotal_cost_part))
                                If Ytotal_billed_part > 0 Then
                                    xlSheet.Cells(Row, "M") = " " & Round(NumericVal(((Ytotal_billed_part - Ytotal_cost_part) / Ytotal_billed_part) * 100), 0) & " %"
                                Else
                                    xlSheet.Cells(Row, "M") = "0%"
                                End If
                                Row = Row + 1
                            ElseIf rsrepordetail!LIVIL = "3" Then
                            
                                ictr = ictr + 1
                                xmat_ctr = xmat_ctr + 1
                                xLIVILx = rsrepordetail!LIVIL
                                
                                xlSheet.Cells(Row, "N") = (rsrepordetail!DETCDE)
                                xlSheet.Range("N" & Row & ":" & "N" & Row).HorizontalAlignment = xlRight
                                If Trim(rsrepordetail!wCode) = "W" Or Trim(rsrepordetail!wCode) = "S" Or Trim(rsrepordetail!wCode) = "C" Then
                                    Ytotal_billed_mat = 0
                                Else
                                    Ytotal_billed_mat = NumericVal(rsrepordetail!DET_AMT - rsrepordetail!Discount_2)
                                End If
                                'sum
                                Xtotal_billed_mat = Xtotal_billed_mat + Ytotal_billed_mat
                                xlSheet.Cells(Row, "O") = Ytotal_billed_mat
                                
                                Ytotal_cost_mat = NumericVal(rsrepordetail!DetCost * rsrepordetail!detvol)
                                Xtotal_Cost_mat = Xtotal_Cost_mat + Ytotal_cost_mat
                                xlSheet.Cells(Row, "P") = Ytotal_cost_mat
                                
                                xlSheet.Cells(Row, "Q") = (Ytotal_billed_mat - Ytotal_cost_mat)
                                Xtotal_mat_GP = NumericVal(Xtotal_mat_GP + (Ytotal_billed_mat - Ytotal_cost_mat))
                                If Ytotal_billed_mat > 0 Then
                                    xlSheet.Cells(Row, "R") = " " & Round(NumericVal(((Ytotal_billed_mat - Ytotal_cost_mat) / Ytotal_billed_mat) * 100), 0) & " %"
                                Else
                                    xlSheet.Cells(Row, "R") = "0%"
                                End If
                                Row = Row + 1
                            End If
                            
                            rsrepordetail.MoveNext
                            If rsrepordetail.EOF Then
                                'total RO_JOB
                               
                                If (xjobs_ctr > xparts_ctr) And (xjobs_ctr > xmat_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xjobs_ctr
                                ElseIf (xparts_ctr > xjobs_ctr) And (xparts_ctr > xmat_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xparts_ctr
                                ElseIf (xmat_ctr > xparts_ctr) And (xmat_ctr > xjobs_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xmat_ctr
                                ElseIf (xjobs_ctr = xparts_ctr) And (xjobs_ctr > xmat_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xjobs_ctr
                                ElseIf (xjobs_ctr = xmat_ctr) And (xjobs_ctr > xparts_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xjobs_ctr
                                ElseIf (xparts_ctr = xjobs_ctr) And (xparts_ctr > xmat_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xparts_ctr
                                ElseIf (xparts_ctr = xmat_ctr) And (xparts_ctr > xjobs_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xparts_ctr
                                ElseIf (xmat_ctr = xjobs_ctr) And (xmat_ctr > xparts_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xmat_ctr
                                ElseIf (xmat_ctr = xparts_ctr) And (xmat_ctr > xjobs_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xmat_ctr
                                End If
                                
                                Row = Row + 1
                                
                                xlSheet.Cells(Row, "D") = "TOTAl RO"
                                'xlSheet.Cells(Row, "E") = Round((XtotalLAbor_billed - NumericVal(rsREPOR!PartLabor)), 2)
                                xlSheet.Cells(Row, "E") = Round(XtotalLAbor_billed, 2)
                                xlSheet.Cells(Row, "F") = XtotalLAbor_cost
                                xlSheet.Cells(Row, "G") = XtotalLAbor_GP
                                
                                '---------------------------------------------------------------------
                                'Xgrndtotal_billed_job = Xgrndtotal_billed_job + Round((XtotalLAbor_billed - NumericVal(rsREPOR!PartLabor)), 2)
                                Xgrndtotal_billed_job = Xgrndtotal_billed_job + Round(XtotalLAbor_billed, 2)
                                Xgrndtotal_cost_job = Xgrndtotal_cost_job + XtotalLAbor_cost
                                Xgrndtotal_GP_job = Xgrndtotal_GP_job + XtotalLAbor_GP
                                '---------------------------------------------------------------------
                                
                                'If Round((XtotalLAbor_billed - NumericVal(rsREPOR!PartLabor)), 2) > 0 Then
                                If XtotalLAbor_billed > 0 Then
                                    'xlSheet.Cells(Row, "H") = " " & Round(((XtotalLAbor_GP / Round((XtotalLAbor_billed - NumericVal(rsREPOR!PartLabor)), 2)) * 100), 0) & "% "
                                    xlSheet.Cells(Row, "H") = " " & Round(((XtotalLAbor_GP / XtotalLAbor_billed) * 100), 0) & "% "
                                Else
                                    xlSheet.Cells(Row, "H") = "0%"
                                End If
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Font.Bold = True
                                xlSheet.Range("E" & Row & ":" & "G" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Borders(xlBottom).LineStyle = xlDouble
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Borders(xlTop).LineStyle = xlDouble
                                
                                'total Parts
                                
                                'xlSheet.Cells(Row, "J") = Round((Xtotal_billed_part - NumericVal(rsREPOR!PartParts)), 2)
                                xlSheet.Cells(Row, "J") = Round(Xtotal_billed_part, 2)
                                xlSheet.Cells(Row, "K") = Xtotal_Cost_past
                                xlSheet.Cells(Row, "L") = Xtotal_part_GP
                                
                                '---------------------------------------------------------------------
                                'Xgrndtotal_billed_part = Xgrndtotal_billed_part + Round((Xtotal_billed_part - NumericVal(rsREPOR!PartParts)), 2)
                                Xgrndtotal_billed_part = Xgrndtotal_billed_part + Round(Xtotal_billed_part, 2)
                                Xgrndtotal_cost_part = Xgrndtotal_cost_part + Xtotal_Cost_past
                                Xgrndtotal_GP_part = Xgrndtotal_GP_part + Xtotal_part_GP
                                '---------------------------------------------------------------------
                                
                                'If Round((Xtotal_billed_part - NumericVal(rsREPOR!PartParts)), 2) > 0 Then
                                If Xtotal_billed_part > 0 Then
                                     'xlSheet.Cells(Row, "M") = " " & Round(((Xtotal_part_GP / Round((Xtotal_billed_part - NumericVal(rsREPOR!PartParts)), 2)) * 100), 0) & "% "
                                     xlSheet.Cells(Row, "M") = " " & Round(((Xtotal_part_GP / Xtotal_billed_part) * 100), 0) & "% "
                                Else
                                     xlSheet.Cells(Row, "M") = "0%"
                                End If
                                xlSheet.Range("I" & Row & ":" & "M" & Row).Font.Bold = True
                                xlSheet.Range("J" & Row & ":" & "L" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("I" & Row & ":" & "M" & Row).Borders(xlBottom).LineStyle = xlDouble
                                xlSheet.Range("I" & Row & ":" & "M" & Row).Borders(xlTop).LineStyle = xlDouble
                                
                                'total Mat
                                
                                'xlSheet.Cells(Row, "O") = Round((Xtotal_billed_mat - NumericVal(rsREPOR!PartMaterials)), 2)
                                xlSheet.Cells(Row, "O") = Round(Xtotal_billed_mat, 2)
                                xlSheet.Cells(Row, "P") = Xtotal_Cost_mat
                                xlSheet.Cells(Row, "Q") = Xtotal_mat_GP
                                
                                '---------------------------------------------------------------------
                                'Xgrndtotal_billed_mat = Xgrndtotal_billed_mat + Round((Xtotal_billed_mat - NumericVal(rsREPOR!PartMaterials)), 2)
                                Xgrndtotal_billed_mat = Xgrndtotal_billed_mat + Round(Xtotal_billed_mat, 2)
                                Xgrndtotal_cost_mat = Xgrndtotal_cost_mat + Xtotal_Cost_mat
                                Xgrndtotal_GP_mat = Xgrndtotal_GP_mat + Xtotal_mat_GP
                                '---------------------------------------------------------------------
                                
                                'If Round((Xtotal_billed_mat - NumericVal(rsREPOR!PartMaterials)), 2) > 0 Then
                                If Xtotal_billed_mat > 0 Then
                                     xlSheet.Cells(Row, "R") = " " & Round(((Xtotal_mat_GP / Xtotal_billed_mat) * 100), 0) & "% "
                                     'xlSheet.Cells(Row, "R") = " " & Round(((Xtotal_mat_GP / Round((Xtotal_billed_mat - NumericVal(rsREPOR!PartMaterials)), 2)) * 100), 0) & "% "
                                Else
                                     xlSheet.Cells(Row, "R") = "0%"
                                End If
                                xlSheet.Range("N" & Row & ":" & "R" & Row).Font.Bold = True
                                xlSheet.Range("O" & Row & ":" & "Q" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("N" & Row & ":" & "R" & Row).Borders(xlBottom).LineStyle = xlDouble
                                xlSheet.Range("N" & Row & ":" & "R" & Row).Borders(xlTop).LineStyle = xlDouble
                                
                                ictr = 0
                                Row = Row + 2
                                rsREPOR.MoveNext
                                GoTo START
                            End If
                            If Trim(rsrepordetail!ROTYPE) = "SR" Then
                                'JOB sub total
                                
                                If (xjobs_ctr > xparts_ctr) And (xjobs_ctr > xmat_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xjobs_ctr
                                ElseIf (xparts_ctr > xjobs_ctr) And (xparts_ctr > xmat_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xparts_ctr
                                ElseIf (xmat_ctr > xparts_ctr) And (xmat_ctr > xjobs_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xmat_ctr
                                ElseIf (xjobs_ctr = xparts_ctr) And (xjobs_ctr > xmat_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xjobs_ctr
                                ElseIf (xjobs_ctr = xmat_ctr) And (xjobs_ctr > xparts_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xjobs_ctr
                                ElseIf (xparts_ctr = xjobs_ctr) And (xparts_ctr > xmat_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xparts_ctr
                                ElseIf (xparts_ctr = xmat_ctr) And (xparts_ctr > xjobs_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xparts_ctr
                                ElseIf (xmat_ctr = xjobs_ctr) And (xmat_ctr > xparts_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xmat_ctr
                                ElseIf (xmat_ctr = xparts_ctr) And (xmat_ctr > xjobs_ctr) Then
                                    Row = Row - ictr
                                    Row = Row + xmat_ctr
                                End If

                                
                                'Job Sub total
                                Row = Row + 1
                                xlSheet.Cells(Row, "D") = "Subtotal"
                                'xlSheet.Cells(Row, "E") = Round((XtotalLAbor_billed - NumericVal(rsREPOR!PartLabor)), 2)
                                xlSheet.Cells(Row, "E") = Round(XtotalLAbor_billed, 2)
                                xlSheet.Cells(Row, "F") = XtotalLAbor_cost
                                xlSheet.Cells(Row, "G") = XtotalLAbor_GP
                                'If Round((XtotalLAbor_billed - NumericVal(rsREPOR!PartLabor)), 2) > 0 Then
                                If XtotalLAbor_billed > 0 Then
                                    xlSheet.Cells(Row, "H") = " " & Round(((XtotalLAbor_GP / XtotalLAbor_billed) * 100), 0) & "% "
                                    'xlSheet.Cells(Row, "H") = " " & Round(((XtotalLAbor_GP / Round((XtotalLAbor_billed - NumericVal(rsREPOR!PartLabor)), 2)) * 100), 0) & "% "
                                Else
                                    xlSheet.Cells(Row, "H") = "0%"
                                End If
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Font.Bold = True
                                xlSheet.Range("E" & Row & ":" & "G" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Borders(xlBottom).LineStyle = xlContinuous
                                xlSheet.Range("D" & Row & ":" & "H" & Row).Borders(xlTop).LineStyle = xlContinuous
                                
                                'Part sub total
                                'xlSheet.Cells(Row, "J") = Round((Xtotal_billed_part - NumericVal(rsREPOR!PartParts)), 2)
                                xlSheet.Cells(Row, "J") = Round(Xtotal_billed_part, 2)
                                xlSheet.Cells(Row, "K") = Xtotal_Cost_past
                                xlSheet.Cells(Row, "L") = Xtotal_part_GP
                                
                                'If Round((Xtotal_billed_part - NumericVal(rsREPOR!PartParts)), 2) > 0 Then
                                If Xtotal_billed_part > 0 Then
                                     'xlSheet.Cells(Row, "M") = " " & Round(((Xtotal_part_GP / Round((Xtotal_billed_part - NumericVal(rsREPOR!PartParts)), 2)) * 100), 0) & "% "
                                     xlSheet.Cells(Row, "M") = " " & Round(((Xtotal_part_GP / Xtotal_billed_part) * 100), 0) & "% "
                                Else
                                     xlSheet.Cells(Row, "M") = "0%"
                                End If
                                xlSheet.Range("J" & Row & ":" & "M" & Row).Font.Bold = True
                                xlSheet.Range("J" & Row & ":" & "L" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("J" & Row & ":" & "M" & Row).Borders(xlBottom).LineStyle = xlContinuous
                                xlSheet.Range("J" & Row & ":" & "M" & Row).Borders(xlTop).LineStyle = xlContinuous

                                'Mat sub total
                                'xlSheet.Cells(Row, "O") = Round((Xtotal_billed_mat - NumericVal(rsREPOR!PartMaterials)), 2)
                                xlSheet.Cells(Row, "O") = Round(Xtotal_billed_mat, 2)
                                xlSheet.Cells(Row, "P") = Xtotal_Cost_mat
                                xlSheet.Cells(Row, "Q") = Xtotal_mat_GP
                                
                                'If Round((Xtotal_billed_mat - NumericVal(rsREPOR!PartMaterials)), 2) > 0 Then
                                If Xtotal_billed_mat > 0 Then
                                     'xlSheet.Cells(Row, "R") = " " & Round(((Xtotal_mat_GP / Round((Xtotal_billed_mat - NumericVal(rsREPOR!PartMaterials)), 2)) * 100), 0) & "% "
                                     xlSheet.Cells(Row, "R") = " " & Round(((Xtotal_mat_GP / Xtotal_billed_mat) * 100), 0) & "% "
                                Else
                                     xlSheet.Cells(Row, "R") = "0%"
                                End If
                                xlSheet.Range("O" & Row & ":" & "R" & Row).Font.Bold = True
                                xlSheet.Range("O" & Row & ":" & "Q" & Row).NumberFormat = MAXIMUM_DIGIT
                                xlSheet.Range("O" & Row & ":" & "R" & Row).Borders(xlBottom).LineStyle = xlContinuous
                                xlSheet.Range("O" & Row & ":" & "R" & Row).Borders(xlTop).LineStyle = xlContinuous
                                
                                ictr = 0
                                Row = Row + 2
                                GoTo START1
                            End If
                            If rsrepordetail!LIVIL > xLIVILx Then
                                Row = Row - ictr
                                ictr = 0
                            End If
                        End If
                    Loop
                rsREPOR.MoveNext
            End If
            rsREPOR.MoveNext
        Loop
        
        'GRAND TOTAL JOBS
        Row = Row + 1
        
        xlSheet.Cells(Row, "D") = "GRAND TOTAL"
        xlSheet.Cells(Row, "E") = Xgrndtotal_billed_job
        xlSheet.Cells(Row, "F") = Xgrndtotal_cost_job
        xlSheet.Cells(Row, "G") = Xgrndtotal_GP_job
        If Xgrndtotal_billed_job > 0 Then
             xlSheet.Cells(Row, "H") = " " & Round(((Xgrndtotal_GP_job / Xgrndtotal_billed_job) * 100), 0) & "% "
        Else
             xlSheet.Cells(Row, "H") = "0%"
        End If
        xlSheet.Range("D" & Row & ":" & "R" & Row).Interior.Color = RGB(200, 160, 10)
        xlSheet.Range("D" & Row & ":" & "R" & Row).Font.Size = "14"
        xlSheet.Range("D" & Row & ":" & "H" & Row).Font.Bold = True
        xlSheet.Range("D" & Row & ":" & "H" & Row).Borders(xlBottom).LineStyle = xlContinuous
        
        '----------------------------------------------------------------------------------
        'GRAND TOTAL PARTS
        xlSheet.Cells(Row, "J") = Xgrndtotal_billed_part
        xlSheet.Cells(Row, "K") = Xgrndtotal_cost_part
        xlSheet.Cells(Row, "L") = Xgrndtotal_GP_part
        If Xgrndtotal_billed_part > 0 Then
             xlSheet.Cells(Row, "M") = " " & Round(((Xgrndtotal_GP_part / Xgrndtotal_billed_part) * 100), 0) & "% "
        Else
             xlSheet.Cells(Row, "M") = "0%"
        End If
        xlSheet.Range("J" & Row & ":" & "M" & Row).Font.Bold = True
        xlSheet.Range("J" & Row & ":" & "M" & Row).Borders(xlBottom).LineStyle = xlContinuous
        
        'GRAND TOTAL MATERIALS
        '----------------------------------------------------------------------------------
        xlSheet.Cells(Row, "O") = Xgrndtotal_billed_mat
        xlSheet.Cells(Row, "P") = Xgrndtotal_cost_mat
        xlSheet.Cells(Row, "Q") = Xgrndtotal_GP_mat
        If Xgrndtotal_billed_mat > 0 Then
             xlSheet.Cells(Row, "R") = " " & Round(((Xgrndtotal_GP_mat / Xgrndtotal_billed_mat) * 100), 0) & "% "
        Else
             xlSheet.Cells(Row, "R") = "0%"
        End If
        
        xlSheet.Range("O" & Row & ":" & "R" & Row).Font.Bold = True
        xlSheet.Range("O" & Row & ":" & "R" & Row).Borders(xlBottom).LineStyle = xlContinuous
        xlSheet.Range("D" & Row & ":" & "R" & Row).Borders(xlBottom).Weight = xlThick
        xlSheet.Range("E" & Row & ":" & "G" & Row).NumberFormat = MAXIMUM_DIGIT
        xlSheet.Range("J" & Row & ":" & "L" & Row).NumberFormat = MAXIMUM_DIGIT
        xlSheet.Range("O" & Row & ":" & "Q" & Row).NumberFormat = MAXIMUM_DIGIT
        
        Row = Row + 3
        xlSheet.Range("A" & Row & ":" & "D" & Row).Merge
        xlSheet.Range("A" & Row & ":" & "A" & Row).Font.Color = vbRed
        xlSheet.Range("A" & Row & ":" & "A" & Row).Font.Italic = True
        xlSheet.Cells(Row, "A") = "Note: *COST = ((Total Hours Rendered) * (Rate Per Hour))"
        
        Prg1.Text = "Processing: ( - )"
        Call enableIT(True)
        xlApp.Visible = True
        Set xlApp = Nothing
        Screen.MousePointer = 0
        MousePointer = 0
    Else
        Prg1.Text = ""
        prgExcelGen.Text = ""
        Prg1.Value = 0
        prgExcelGen.Value = 0
        MousePointer = 0
        ShowNoRecord
        Call enableIT(True)
    End If
    Exit Sub
IVANEXEQUIELVALENCIA:
        If Err = 91 Then
            MousePointer = 0
            Prg1.Text = ""
            prgExcelGen.Text = ""
            Prg1.Value = 0
            prgExcelGen.Value = 0
        Else
            MsgBox Error
            Prg1.Text = ""
            prgExcelGen.Text = ""
            Prg1.Value = 0
            prgExcelGen.Value = 0
        End If
    Exit Sub
    
End Sub


Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    initcombo
End Sub
Sub initcombo()
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
End Sub

Function setcusname(xuscode As String) As String
    Dim rsCusmas As ADODB.Recordset
    Set rsCusmas = New ADODB.Recordset
    
    Set rsCusmas = gconDMIS.Execute("Select cusnam from ALL_CusMas where  cuscde = '" & Null2String(xuscode) & "'")
    
    If Not (rsCusmas.EOF And rsCusmas.BOF) Then
        setcusname = Null2String(rsCusmas!CUSNAM)
    Else
        setcusname = ""
    End If
    Set rsCusmas = Nothing
    Exit Function
End Function

Function setCost_of_labor(XRONO As String, Xline As String, xTECHCODE As String) As Double
    Dim rstimeconsumed                                          As ADODB.Recordset
    Dim xtekcode                                                As String
    
    Set rstimeconsumed = New ADODB.Recordset
    Set rstimeconsumed = gconDMIS.Execute("Select hrsworked,technician from CSMS_JobClock where ro_no = '" & XRONO & "' and line_no = '" & rsrepordetail!LINE_NO & "' and techcode = '" & xTECHCODE & "'")
    i = 0
    If Not (rstimeconsumed.EOF And rstimeconsumed.BOF) Then
        xtekcode = rstimeconsumed!Technician
        rstimeconsumed.MoveFirst
        Do While Not rstimeconsumed.EOF
            i = i + rstimeconsumed!hrsWorked
            rstimeconsumed.MoveNext
        Loop
    End If
    
    If i > rsrepordetail!HRSWRK Then
        setCost_of_labor = NumericVal(i * getrateperhour(xtekcode))
    Else
        setCost_of_labor = NumericVal(rsrepordetail!HRSWRK * getrateperhour(xtekcode))
    End If
    Set rstimeconsumed = Nothing
    Exit Function
End Function

Function getrateperhour(xTECHCODE As String) As Double
    Dim rsbasicrate                                     As ADODB.Recordset
    Dim rsgetteccode                                    As ADODB.Recordset
    Dim lobotmo                                         As String
    
    
    Set rsgetteccode = New ADODB.Recordset
    Set rsbasicrate = New ADODB.Recordset
    lobotmo = ""
    Set rsgetteccode = gconDMIS.Execute("Select technician from CSMS_JobClock where ro_no = '" & rsrepordetail!REP_OR & "' and line_no = '" & rsrepordetail!LINE_NO & "' and techcode = '" & Trim(xTECHCODE) & "' ")
    If Not (rsgetteccode.EOF And rsgetteccode.BOF) Then
        lobotmo = rsgetteccode!Technician
    End If
    Set rsbasicrate = gconDMIS.Execute("select salarycode from hrms_empinfo where empno = '" & xTECHCODE & "'")
    If Not (rsbasicrate.EOF And rsbasicrate.BOF) Then
        getrateperhour = ((((gconDMIS.Execute("Select salary from HRMS_SALARYGRADE where code = '" & Null2String(rsbasicrate!salarycode) & "'").Fields(0).Value) * 12) / 314) / 8)
    Else
        getrateperhour = 0
    End If
    Set rsbasicrate = Nothing
    Exit Function
End Function


Sub enableIT(XXX As Boolean)
    cmdPrint.Enabled = XXX
    cmdCancel.Enabled = XXX
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rsREPOR = Nothing
  Set rsrepordetail = Nothing
End Sub

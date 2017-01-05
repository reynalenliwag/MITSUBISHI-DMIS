VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCSMS_Report_SummaryofInternalSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Summary of Internal Sales"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4245
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSReport_SummaryofInternalSales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   4245
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   795
      Left            =   2100
      MouseIcon       =   "frmCSMSReport_SummaryofInternalSales.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSReport_SummaryofInternalSales.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   1410
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   795
      Left            =   1380
      MouseIcon       =   "frmCSMSReport_SummaryofInternalSales.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSReport_SummaryofInternalSales.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1410
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4035
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54984705
         CurrentDate     =   40066
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54984705
         CurrentDate     =   40066
      End
   End
   Begin wizProgBar.Prg prgExcelGen 
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   476
      Picture         =   "frmCSMSReport_SummaryofInternalSales.frx":27F3
      ForeColor       =   0
      Appearance      =   2
      BorderStyle     =   2
      BarForeColor    =   8454016
      BarPicture      =   "frmCSMSReport_SummaryofInternalSales.frx":280F
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
End
Attribute VB_Name = "frmCSMS_Report_SummaryofInternalSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    cmdPrint.Enabled = False
    Screen.MousePointer = 11
    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet                                        As Excel.Worksheet
    Dim rsIntCharge                                    As New ADODB.Recordset
    Dim rsCodeDesc                                     As New ADODB.Recordset
    Dim rslabor                                        As New ADODB.Recordset
    Dim rsParts                                        As New ADODB.Recordset
    Dim rsMaterials                                    As New ADODB.Recordset
    
    Dim rssubparts                                     As New ADODB.Recordset
    Dim rssubmat                                       As New ADODB.Recordset
    Dim rssublabor                                     As New ADODB.Recordset
    Dim rssubparts2                                    As New ADODB.Recordset
    Dim rssubmat2                                      As New ADODB.Recordset
    Dim rssublabor2                                    As New ADODB.Recordset

    
    Dim InvoiceDte                                         As Date
    Dim strRoNo                                            As String
    Dim strCustname                                        As String
    Dim intRoAmount                                        As Double
    Dim intCompnyAmt                                       As Double
    Dim intlabor                                           As Double
    Dim CompnySubTtal                                      As Double
    Dim intParts                                           As Double
    Dim intMaterials                                       As Double
    Dim strInternalCharges                                 As String
    Dim strInternalCharges2                                As String
        
    Dim TotalRoAmt                                         As Double
    Dim TotalCompAmt                                       As Double
    Dim TotalLaborAmt                                      As Double
    Dim TotalPartsAmt                                      As Double
    Dim TotalMatAmt                                        As Double
    Dim dblLabor                                           As Double
    Dim SubRoAmnt                                          As Double
    Dim dblParts                                           As Double
    Dim dblMaterials                                       As Double
    Dim VAT                                                As Double
    Dim COUNTER                                            As Integer
    Dim SQL                                                As String
    Dim SQL1                                               As String
    Dim SQL2                                               As String
    Dim SQLLABOR                                           As String
    Dim SQLPARTS                                           As String
    Dim SQLMATERIALS                                       As String
    Dim CloneRo                                            As String
    Dim rochange                                           As Boolean
    Dim SubtPrint                                          As Boolean
    
'-------------------------------------------------------------------
    Dim sqlsubparts                                        As String
    Dim sqlsubmat                                          As String
    Dim sqlsublabor                                        As String
    Dim totsublabor                                        As Double
    Dim totsubmat                                          As Double
    Dim totsubparts                                        As Double
    Dim totsublaboramt                                     As Double
    Dim totsubmatamt                                       As Double
    Dim totsubpartsamt                                     As Double
    Dim totsubdbllaboramt                                  As Double
    Dim totsubdblmatamt                                    As Double
    Dim totsubdblpartsamt                                  As Double
    Dim totsubcomp_laboramt                                As Double
    Dim totsubcomp_partsamt                                As Double
    Dim totsubcomp_matamt                                  As Double
    Dim sqlsubparts2                                       As String
    Dim sqlsubmat2                                         As String
    Dim sqlsublabor2                                       As String
    Dim totsublabor2                                       As Double
    Dim totsubmat2                                         As Double
    Dim totsubparts2                                       As Double
    Dim totsublaboramt2                                    As Double
    Dim totsubmatamt2                                      As Double
    Dim totsubpartsamt2                                    As Double
    Dim totsubdbllaboramt2                                 As Double
    Dim totsubdblmatamt2                                   As Double
    Dim totsubdblpartsamt2                                 As Double
    Dim totsubcomp_laboramt2                               As Double
    Dim totsubcomp_partsamt2                               As Double
    Dim totsubcomp_matamt2                                 As Double
 '------------------------------------------------------------------
 
    totsublabor = 0
    totsubmat = 0
    totsubparts = 0
    totsublaboramt = 0
    totsubmatamt = 0
    totsubpartsamt = 0
    totsubdbllaboramt = 0
    totsubdblmatamt = 0
    totsubdblpartsamt = 0
    totsubcomp_laboramt = 0
    totsubcomp_partsamt = 0
    totsubcomp_matamt = 0
    totsublabor2 = 0
    totsubmat2 = 0
    totsubparts2 = 0
    totsublaboramt2 = 0
    totsubmatamt2 = 0
    totsubpartsamt2 = 0
    totsubdbllaboramt2 = 0
    totsubdblmatamt2 = 0
    totsubdblpartsamt2 = 0
    totsubcomp_laboramt2 = 0
    totsubcomp_partsamt2 = 0
    totsubcomp_matamt2 = 0
    
    prgExcelGen.Text = ""

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "Internal Charges Report.XLT")
    Set xlSheet = xlBook.Worksheets(1)

    xlSheet.Cells(1, "A") = COMPANY_NAME
    xlSheet.Cells(2, "A") = COMPANY_ADDRESS

    VAT = 1.12
    COUNTER = 9
    
    
    xlSheet.Cells(4, "A") = "Internal Charges"
    xlSheet.Cells(5, "A") = "For the Month of " & UCase(MonthName(Month(DTPicker1))) & ""
    
    SQL = ("SELECT  CSMS_Repor.invoice,CSMS_Repor.dte_comp,CSMS_Repor.REP_OR, " & _
        "CSMS_Repor.NIYM,CSMS_Repor.RO_AMOUNT ,CSMS_Ro_Det.DET_AMT , CSMS_Repor.labor, " & _
        "CSMS_Repor.parts, CSMS_Repor.material, CSMS_Ro_Det.Code, CSMS_Ro_Det.livil, CSMS_Ro_Det.LINE_NO " & _
        "From " & _
        "CSMS_Repor INNER JOIN  CSMS_Ro_Det ON CSMS_Repor.REP_OR = CSMS_Ro_Det.REP_OR " & _
        "where " & _
        "(CSMS_Repor.TRANSTYPE = 'R') AND (CSMS_Repor.DTE_COMP IS NOT NULL) AND " & _
        "(CSMS_Ro_Det.TRANSTYPE = 'R') AND  (CSMS_Ro_Det.WCODE IN ('S', 'C'))  AND " & _
        "(CSMS_Repor.DTE_REL) BETWEEN '" & DTPicker1 & "' AND '" & DTPicker2 & "' " & _
        "order by CSMS_Repor.dte_comp,CSMS_Repor.rep_or,CSMS_Ro_Det.livil,CSMS_Ro_Det.line_no")
        
    rsIntCharge.Open (SQL), gconDMIS
    Dim cnt                 As Integer
    cnt = 1
    
    If Not (rsIntCharge.BOF And rsIntCharge.EOF) Then
        prgExcelGen.Max = rsIntCharge.RecordCount
        
        Do While Not rsIntCharge.EOF
            prgExcelGen.Value = cnt
            prgExcelGen.Text = Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %"

            DoEvents
        
      
   
            
         SubtPrint = False
         CloneRo = strRoNo
                             
                             
         InvoiceDte = Null2String(Trim(rsIntCharge!dte_comp))
         strRoNo = Null2String(Trim(rsIntCharge!REP_OR))
         strCustname = Null2String(Trim(rsIntCharge!NIYM))
         intRoAmount = Null2String(Trim(rsIntCharge!ro_amount))
         intCompnyAmt = Null2String(Trim(rsIntCharge!DET_AMT))
         strInternalCharges = Null2String(Trim(rsIntCharge!Code))
          
          
         SQL2 = ("select description from amis_chartaccount where acctCode in (select chartcodes from cmis_sbook where code = '" & strInternalCharges & "' )")
         SQLLABOR = ("select  isnull(sum(det_amt),0) as labor from csms_ro_det where WCODE IN ('S', 'C') and rep_or = '" & strRoNo & "' and livil = 1")
         SQLPARTS = ("select  isnull(sum(det_amt),0) as parts from csms_ro_det where WCODE IN ('S', 'C') and rep_or = '" & strRoNo & "' and livil = 2")
         SQLMATERIALS = ("select  isnull(sum(det_amt),0) as materials from csms_ro_det where WCODE IN ('S', 'C') and  rep_or = '" & strRoNo & "' and livil = 3")
'updated by:    IEBV 08162010_0136pm
'description:   Query for getting the sublet amount of the RO
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         sqlsublabor = ("select isnull(sum(det_amt),0) as totalsublabor from csms_po_dt where wcode in('S','C') and rep_or = '" & strRoNo & "' and livil = 1 and status = 'P' and jobtype = 'GJ'")
         sqlsubparts = ("select isnull(sum(det_amt),0) as totalsubparts from csms_po_dt where wcode in('S','C') and rep_or = '" & strRoNo & "' and livil = 2 and status = 'P' and jobtype = 'GJ'")
         sqlsubmat = ("select isnull(sum(det_amt),0) as totalsubmat from csms_po_dt where wcode in('S','C') and rep_or = '" & strRoNo & "' and livil = 3 and status = 'P' and jobtype = 'GJ'")
         
         sqlsublabor2 = ("select isnull(sum(det_amt),0) as totalsublabor2 from csms_po_dt where wcode in('S','C') and rep_or = '" & strRoNo & "' and livil = 1 and status = 'P' and jobtype = 'BP'")
         sqlsubparts2 = ("select isnull(sum(det_amt),0) as totalsubparts2 from csms_po_dt where wcode in('S','C') and rep_or = '" & strRoNo & "' and livil = 2 and status = 'P' and jobtype = 'BP'")
         sqlsubmat2 = ("select isnull(sum(det_amt),0) as totalsubmat2 from csms_po_dt where wcode in('S','C') and rep_or = '" & strRoNo & "' and livil = 3 and status = 'P' and jobtype = 'BP'")
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
         rslabor.Open (SQLLABOR), gconDMIS
         rsParts.Open (SQLPARTS), gconDMIS
         rsMaterials.Open (SQLMATERIALS), gconDMIS
         rsCodeDesc.Open (SQL2), gconDMIS
'updated by:    IEBV 08162010_0136pm
'---------------------------------------------------
         rssublabor.Open (sqlsublabor), gconDMIS
         rssubparts.Open (sqlsubparts), gconDMIS
         rssubmat.Open (sqlsubmat), gconDMIS
         
         rssublabor2.Open (sqlsublabor2), gconDMIS
         rssubparts2.Open (sqlsubparts2), gconDMIS
         rssubmat2.Open (sqlsubmat2), gconDMIS
'---------------------------------------------------
        
        If Not ((rsCodeDesc.BOF And rsCodeDesc.EOF)) Or Not ((rslabor.BOF And rslabor.EOF)) Or Not ((rsParts.BOF And rsParts.EOF)) Or Not ((rsMaterials.BOF And rsMaterials.EOF)) Then
        'GET INTERNAL CHARGE DESCRIPTION
            On Error Resume Next
            intlabor = Null2String(rslabor!labor)
            intParts = Null2String(rsParts!parts)
            intMaterials = Null2String(rsMaterials!MATERIALs)
            strInternalCharges2 = rsCodeDesc!Description

'----------------------------------------------------------------------
            totsublabor = Null2String(rssublabor!totalsublabor)
            totsubparts = Null2String(rssubparts!totalsubparts)
            totsubmat = Null2String(rssubmat!totalsubmat)
            
            totsublabor2 = Null2String(rssublabor2!totalsublabor2)
            totsubparts2 = Null2String(rssubparts2!totalsubparts2)
            totsubmat2 = Null2String(rssubmat2!totalsubmat2)
'----------------------------------------------------------------------

        End If
          If rochange And CloneRo <> strRoNo Then
            'subtotal
            xlSheet.Cells(COUNTER, "D") = "SUBTOTAL"
            xlSheet.Cells(COUNTER, "E") = SubRoAmnt
            xlSheet.Cells(COUNTER, "F") = dblLabor
            xlSheet.Cells(COUNTER, "G") = dblParts
            xlSheet.Cells(COUNTER, "H") = dblMaterials
'updated by:    IEBV 08162010_0200pm
'description:   To display the sub total sublet amount of the RO
'-------------------------------------------------------------------
            xlSheet.Cells(COUNTER, "I") = totsubdbllaboramt
            xlSheet.Cells(COUNTER, "J") = totsubdblpartsamt
            xlSheet.Cells(COUNTER, "K") = totsubdblmatamt

            xlSheet.Cells(COUNTER, "L") = totsubdbllaboramt2
            xlSheet.Cells(COUNTER, "M") = totsubdblpartsamt2
            xlSheet.Cells(COUNTER, "N") = totsubdblmatamt2
'-------------------------------------------------------------------
            
            TotalRoAmt = intRoAmount + TotalRoAmt
            TotalCompAmt = SubRoAmnt + TotalCompAmt
            TotalLaborAmt = dblLabor + TotalLaborAmt
            TotalPartsAmt = dblParts + TotalPartsAmt
            TotalMatAmt = dblMaterials + TotalMatAmt
           
'updated by:    IEBV 08162010_0200pm
'description:   To compute the total sublet amount of the RO
'--------------------------------------------------------------------------
            totsubcomp_laboramt = totsubdbllaboramt + totsubcomp_laboramt
            totsubcomp_partsamt = totsubdblpartsamt + totsubcomp_partsamt
            totsubcomp_matamt = totsubdblmatamt + totsubdblmatamt
            
            totsubcomp_laboramt2 = totsubdbllaboramt2 + totsubcomp_laboramt2
            totsubcomp_partsamt2 = totsubdblpartsamt2 + totsubcomp_partsamt2
            totsubcomp_matamt2 = totsubdblmatamt2 + totsubdblmatamt2
'--------------------------------------------------------------------------

            COUNTER = COUNTER + 2
            SubRoAmnt = 0
          ElseIf CloneRo = "" Then
            COUNTER = COUNTER + 1
            xlSheet.Cells(COUNTER, "F") = intlabor
            xlSheet.Cells(COUNTER, "G") = intParts
            xlSheet.Cells(COUNTER, "H") = intMaterials
'updated by:    IEBV 08162010_0200pm
'description:   To display the sublet amount of the RO
'--------------------------------------------------------------------------
            xlSheet.Cells(COUNTER, "I") = totsublabor
            xlSheet.Cells(COUNTER, "J") = totsubparts
            xlSheet.Cells(COUNTER, "K") = totsubmat
            
            xlSheet.Cells(COUNTER, "L") = totsublabor2
            xlSheet.Cells(COUNTER, "M") = totsubparts2
            xlSheet.Cells(COUNTER, "N") = totsubmat2
'--------------------------------------------------------------------------
            COUNTER = COUNTER - 1
          Else
            'if ro has multiple compny amount
            'xlSheet.Cells(counter, "D") = intRoAmount
            xlSheet.Cells(COUNTER, "E") = intCompnyAmt
            SubtPrint = True
          End If

            
            If SubtPrint = False Then
                 'print every R.O. thas has no multiple company amount details
                 xlSheet.Cells(COUNTER, "A") = InvoiceDte
                 xlSheet.Cells(COUNTER, "B") = strRoNo
                 xlSheet.Cells(COUNTER, "C") = strCustname
                 xlSheet.Cells(COUNTER, "D") = intRoAmount
                 xlSheet.Cells(COUNTER, "E") = intCompnyAmt
                 xlSheet.Cells(COUNTER, "F") = intlabor
                 xlSheet.Cells(COUNTER, "G") = intParts
                 xlSheet.Cells(COUNTER, "H") = intMaterials
'updated by:    IEBV 08162010_0210pm
'description:   To display the sublet amount of the RO
'----------------------------------------------------------------------------
                 xlSheet.Cells(COUNTER, "I") = totsublabor
                 xlSheet.Cells(COUNTER, "J") = totsubparts
                 xlSheet.Cells(COUNTER, "K") = totsubmat
                 
                 xlSheet.Cells(COUNTER, "L") = totsublabor2
                 xlSheet.Cells(COUNTER, "M") = totsubparts2
                 xlSheet.Cells(COUNTER, "N") = totsubmat2
'---------------------------------------------------------------------------
                 xlSheet.Cells(COUNTER, "O") = strInternalCharges2
                
                
            End If
            'details for subtotal
            dblLabor = intlabor
            dblParts = intParts
            dblMaterials = intMaterials
'updated by:    IEBV 08162010_0210pm
'description:   For the subtotal sublet amount of the RO
'----------------------------------------------------------------------------
            totsubdbllaboramt = totsublabor
            totsubdblpartsamt = totsubparts
            totsubdblmatamt = totsubmat
            
            totsubdbllaboramt2 = totsublabor2
            totsubdblpartsamt2 = totsubparts2
            totsubdblmatamt2 = totsubmat2
'----------------------------------------------------------------------------
                                   
                 
          If CloneRo <> strRoNo Then
            rochange = True
          End If
                  
            cnt = cnt + 1
            COUNTER = COUNTER + 1
            CompnySubTtal = intCompnyAmt + CompnySubTtal
            SubRoAmnt = intCompnyAmt + SubRoAmnt
            rsIntCharge.MoveNext


            
            Set rsCodeDesc = Nothing
            Set rslabor = Nothing
            Set rsParts = Nothing
            Set rsMaterials = Nothing
'--------------------------------------------------------------------------
            Set rssublabor = Nothing
            Set rssubparts = Nothing
            Set rssubmat = Nothing
            
            Set rssublabor2 = Nothing
            Set rssubparts2 = Nothing
            Set rssubmat2 = Nothing
'--------------------------------------------------------------------------

        Loop
        
            'LAST DETAILS
            xlSheet.Cells(COUNTER, "D") = "SUBTOTAL"
            xlSheet.Cells(COUNTER, "E") = SubRoAmnt
            xlSheet.Cells(COUNTER, "F") = dblLabor
            xlSheet.Cells(COUNTER, "G") = dblParts
            xlSheet.Cells(COUNTER, "H") = dblMaterials
'updated by:    IEBV 08162010_0210pm
'description:   To display the total sublet amount of all the RO
'----------------------------------------------------------------------
            xlSheet.Cells(COUNTER, "I") = totsubdbllaboramt
            xlSheet.Cells(COUNTER, "J") = totsubdblpartsamt
            xlSheet.Cells(COUNTER, "K") = totsubdblmatamt
            
            xlSheet.Cells(COUNTER, "L") = totsubdbllaboramt2
            xlSheet.Cells(COUNTER, "M") = totsubdblpartsamt2
            xlSheet.Cells(COUNTER, "N") = totsubdblmatamt2
'------------------------------------------------------------------------
            'TotalRoAmt = intRoAmount + TotalRoAmt
            TotalCompAmt = SubRoAmnt + TotalCompAmt
            TotalLaborAmt = dblLabor + TotalLaborAmt
            TotalPartsAmt = dblParts + TotalPartsAmt
            TotalMatAmt = dblMaterials + TotalMatAmt
'updated by:    IEBV 08162010_0210pm
'description:   To compute the total sublet amount of all the RO
'---------------------------------------------------------------------------
            totsubcomp_laboramt = totsubdbllaboramt + totsubcomp_laboramt
            totsubcomp_partsamt = totsubdblpartsamt + totsubcomp_partsamt
            totsubcomp_matamt = totsubdblmatamt + totsubcomp_matamt
            
            totsubcomp_laboramt2 = totsubdbllaboramt2 + totsubcomp_laboramt2
            totsubcomp_partsamt2 = totsubdblpartsamt2 + totsubcomp_partsamt2
            totsubcomp_matamt2 = totsubdblmatamt2 + totsubcomp_matamt2
'---------------------------------------------------------------------------

             COUNTER = COUNTER + 2
                
             xlSheet.Cells(COUNTER, "D") = "TOTAL"
             'xlSheet.Cells(counter, "D") = TotalRoAmt
             xlSheet.Cells(COUNTER, "E") = TotalCompAmt
             xlSheet.Cells(COUNTER, "F") = TotalLaborAmt
             xlSheet.Cells(COUNTER, "G") = TotalPartsAmt
             xlSheet.Cells(COUNTER, "H") = TotalMatAmt
'updated by:    IEBV 08162010_0215pm
'description:   To display the  computed sublet amount of all the RO
'---------------------------------------------------------------------------
             xlSheet.Cells(COUNTER, "I") = totsubcomp_laboramt
             xlSheet.Cells(COUNTER, "J") = totsubcomp_partsamt
             xlSheet.Cells(COUNTER, "K") = totsubcomp_matamt
             
             xlSheet.Cells(COUNTER, "L") = totsubcomp_laboramt2
             xlSheet.Cells(COUNTER, "M") = totsubcomp_partsamt2
             xlSheet.Cells(COUNTER, "N") = totsubcomp_matamt2
'--------------------------------------------------------------------------

            
            xlSheet.Range("D" & COUNTER & ":" & "N" & COUNTER).Font.Bold = True
            xlSheet.Range("D" & COUNTER & ":" & "N" & COUNTER).Borders(xlBottom).LineStyle = xlDouble
            xlSheet.Range("D" & COUNTER & ":" & "N" & COUNTER).Borders(xlTop).LineStyle = xlDouble
        
        prgExcelGen.Text = "Generation (100% Completed)"
        xlApp.Visible = True
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "SUMMARY OF INTERNAL SALES", "", "", "", "SUMMARY - " & DTPicker1 & " TO " & DTPicker2, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    Else
        Call ShowNoRecord
    End If

    Set xlApp = Nothing
    cmdPrint.Enabled = True
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
 
    Call CenterMe(frmMain, Me, 1)
    DTPicker1.Value = firstDay(Date)
    DTPicker2.Value = Date
    
    Screen.MousePointer = 0
End Sub










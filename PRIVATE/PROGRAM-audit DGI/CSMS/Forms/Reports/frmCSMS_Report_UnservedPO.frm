VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMS_Report_UnservedPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unserved Sublet PO"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3900
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_Report_UnservedPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   3900
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1943
      MouseIcon       =   "frmCSMS_Report_UnservedPO.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_Report_UnservedPO.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   780
      Width           =   735
   End
   Begin Crystal.CrystalReport RPT 
      Left            =   3120
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker TXTFrom 
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   40126
   End
   Begin MSComCtl2.DTPicker txtTO 
      Height          =   345
      Left            =   2010
      TabIndex        =   2
      Top             =   360
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   40126
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1230
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_Report_UnservedPO.frx":161F
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   780
      Width           =   735
   End
   Begin VB.Label labCap 
      AutoSize        =   -1  'True
      Caption         =   "Date Range ( From ~ To )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   953
      TabIndex        =   4
      Top             =   90
      Width           =   1995
   End
End
Attribute VB_Name = "frmCSMS_Report_UnservedPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim RSTMP                                               As New ADODB.Recordset
    Dim xlApp                                               As Excel.Application
    Dim xlBook                                              As Excel.Workbook
    Dim xlSheet                                             As Excel.Worksheet
    Dim cnt                                                 As Integer
    Dim REC_CNT                                             As Long
    
    cnt = 2
    Set RSTMP = gconDMIS.Execute("SELECT PO_NO, RO_NO, PO_DATE, CUST_NAME, CONTRACTOR_NAME, SUBLET_TOTAL_AMT, SUBLET_TOTAL_VAT, SUBLET_TOTAL_NET_AMT FROM CSMS_PO_HD WHERE PO_NO NOT IN " & _
        " ( " & _
        " SELECT PO_NO FROM CSMS_PO_RC_HD WHERE STATUS <> 'C'" & _
        " ) " & _
        " AND PO_DATE BETWEEN " & N2Str2Null(TXTFrom) & " AND " & N2Str2Null(txtTO) & _
        " AND STATUS = 'P'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        On Error GoTo ERROR_MSG
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "Unserved Sublet Report.xlt")
        Set xlSheet = xlBook.Worksheets(1)
        
        REC_CNT = RSTMP.RecordCount
        xlSheet.Cells(1, "A") = COMPANY_NAME
        xlSheet.Cells(2, "A") = COMPANY_ADDRESS
        xlSheet.Cells(5, "A") = "Date Range " & TXTFrom & " - " & txtTO
        xlSheet.Cells(7, 1).CopyFromRecordset RSTMP
        Set RSTMP = Nothing
        
        Set RSTMP = gconDMIS.Execute("SELECT SUM(ISNULL(SUBLET_TOTAL_AMT,0)) FROM CSMS_PO_HD WHERE PO_NO NOT IN  (  SELECT PO_NO FROM CSMS_PO_RC_HD WHERE STATUS <> 'C' )  AND PO_DATE BETWEEN " & N2Str2Null(TXTFrom) & " AND " & N2Str2Null(txtTO) & " AND STATUS = 'P'")
        xlSheet.Cells(7 + REC_CNT, "F") = NumericVal(RSTMP.Fields(0))
        Set RSTMP = gconDMIS.Execute("SELECT SUM(ISNULL(SUBLET_TOTAL_VAT,0)) FROM CSMS_PO_HD WHERE PO_NO NOT IN  (  SELECT PO_NO FROM CSMS_PO_RC_HD WHERE STATUS <> 'C' )  AND PO_DATE BETWEEN " & N2Str2Null(TXTFrom) & " AND " & N2Str2Null(txtTO) & " AND STATUS = 'P'")
        xlSheet.Cells(7 + REC_CNT, "G") = NumericVal(RSTMP.Fields(0))
        Set RSTMP = gconDMIS.Execute("SELECT SUM(ISNULL(SUBLET_TOTAL_NET_AMT,0)) FROM CSMS_PO_HD WHERE PO_NO NOT IN  (  SELECT PO_NO FROM CSMS_PO_RC_HD WHERE STATUS <> 'C' )  AND PO_DATE BETWEEN " & N2Str2Null(TXTFrom) & " AND " & N2Str2Null(txtTO) & " AND STATUS = 'P'")
        xlSheet.Cells(7 + REC_CNT, "H") = NumericVal(RSTMP.Fields(0))
        
        xlSheet.Range("F" & 7 + REC_CNT & ":" & "H" & 7 + REC_CNT).Font.Bold = True
        xlSheet.Range("F" & 7 + REC_CNT & ":" & "H" & 7 + REC_CNT).Borders(xlBottom).LineStyle = xlDouble
        xlSheet.Range("F" & 7 + REC_CNT & ":" & "H" & 7 + REC_CNT).Borders(xlTop).LineStyle = xlDouble
        xlApp.Windows.ITEM(1).Caption = "Unserved Sublet Report"
        xlApp.Visible = True
        
        Set xlApp = Nothing
        Set xlSheet = Nothing
        Set xlBook = Nothing
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "UNSERVED SUBLET PO REPORTS", "", "", "", "SUMMARY - " & TXTFrom & " TO " & txtTO, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    Else
        Call ShowNoRecord
    End If
    Set RSTMP = Nothing
    
    Exit Sub
ERROR_MSG:
    'MsgBox Err.Number & " - " & Err.Description
    If Err.Number = 1004 Then
        MsgBox "Report File not found", vbCritical, "Error"
        'MsgBox Err.Number & " - " & Err.Description
    Else
        MsgBox "Unknown Error occured", vbCritical, "Error"
    End If
    
    Set xlApp = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    
    TXTFrom.Value = firstDay(Date)
    txtTO.Value = Date
    
    Screen.MousePointer = 0
End Sub

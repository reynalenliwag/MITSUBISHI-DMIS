VERSION 5.00
Begin VB.Form frmHRMS_Reports_IndividualPayrollSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Individual Payroll Summary"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   Icon            =   "IndividualPayrollSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
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
      Left            =   2250
      MouseIcon       =   "IndividualPayrollSummary.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "IndividualPayrollSummary.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   2040
      Width           =   1215
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
      Left            =   1050
      MouseIcon       =   "IndividualPayrollSummary.frx":0B27
      MousePointer    =   99  'Custom
      Picture         =   "IndividualPayrollSummary.frx":0C79
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2025
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   4125
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "IndividualPayrollSummary.frx":1118
         Left            =   90
         List            =   "IndividualPayrollSummary.frx":111A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1260
         Width           =   3735
      End
      Begin VB.ComboBox cboReport 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "IndividualPayrollSummary.frx":111C
         Left            =   90
         List            =   "IndividualPayrollSummary.frx":111E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Select Employee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select Employee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmHRMS_Reports_IndividualPayrollSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApplication As Excel.Application
Dim xlBook    As Excel.Workbook
Dim xlSheet   As Excel.Worksheet
Private Sub Form_Load()
 CenterMe frmMain, Me, 1
 InitCombo
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Sub InitCombo()
 With cboReport
    .AddItem "Daily Retail Sales Report", 0
    .AddItem "Montly Retail Sales Report", 1
    .AddItem "Retail Sales Agent Performance Report", 2
    .AddItem "Network Retail Sales-Yearly", 3
    .AddItem "Network Retail Sales-Monthly", 4
    .ListIndex = 0
 End With
End Sub
Private Sub cmdPrint_Click()
    Select Case cboReport.ListIndex
        Case 0
            Call DailyRetailReport
        Case 1
            Call MonthlyRetailReport
        Case 2
        Case 3
            Call NetworkRetailSalesYearly
        Case 4
            Call NetworkRetailSalesMonthly
    End Select
End Sub


Sub NetworkRetailSalesYearly()
                Dim rsDealer      As ADODB.Recordset
                Dim rsModel       As ADODB.Recordset
                Dim rsCount       As ADODB.Recordset
                Dim xcol As Integer
                Dim jrow As Integer
                Dim intcol As Integer
                Dim introw As Integer
                Set xlApplication = CreateObject("Excel.Application")
                Set xlBook = xlApplication.Workbooks.Open(SMIS_REPORT_PATH & "EXCEL\NEWORK RETAIL SALES-YEARLY.xlt")
                Set xlSheet = xlApplication.Worksheets(1)
                
                Set rsDealer = New ADODB.Recordset
                Set rsDealer = gconDMIS.Execute("SELECT DEALER_CODE,* FROM ALL_DEALERS ORDER BY PROVINCIAL")
                Set rsModel = New ADODB.Recordset
                Set rsModel = gconDMIS.Execute("SELECT DISTINCT(MODEL) FROM ALL_MODEL ORDER BY MODEL")
                xcol = 3
                jrow = 5
                While Not rsDealer.EOF
                        xlSheet.Cells(4, xcol) = Null2String(rsDealer!DEALER_CODE)
                        While Not rsModel.EOF
                                xlSheet.Cells(jrow, "B") = Null2String(rsModel!MODEL)
                                        Set rsCount = gconDMIS.Execute("SELECT COUNT(*) As Tcount FROM SMIS_RETAILSALES WHERE MODEL_CODE='" & Null2String(rsModel!MODEL) & "' AND DEALER_CODE='" & Null2String(rsDealer!DEALER_CODE) & "'")
                                        xlSheet.Cells(introw + 5, intcol + 3) = Null2String(rsCount!TCOUNT)
                                introw = introw + 1
                                jrow = jrow + 1
                                rsModel.MoveNext
                        
                        Wend
                xcol = xcol + 1
                intcol = intcol + 1
                rsDealer.MoveNext
                Wend
                xlApplication.Visible = True
                Set xlApplication = Nothing
                Set xlBook = Nothing
                Set xlSheet = Nothing
End Sub
Sub DailyRetailReport()
    Dim xlApplication As Excel.Application
    Dim xlBook        As Excel.Workbook
    Dim xlSheet       As Excel.Worksheet
    Dim rsDealer      As ADODB.Recordset
    Dim rsModel       As ADODB.Recordset
    Dim rsVariant     As ADODB.Recordset
    Dim rsCount       As ADODB.Recordset
    Dim I             As Integer
    Dim j             As Integer
    Dim L             As Integer
    Dim X             As Integer
    Dim n             As Integer
    Set rsDealer = New ADODB.Recordset
    Set rsDealer = gconDMIS.Execute("SELECT DEALER_CODE FROM ALL_DEALERS ORDER BY PROVINCIAL")

    Set rsModel = New ADODB.Recordset
    Set rsModel = gconDMIS.Execute("SELECT DISTINCT MODEL FROM ALL_MODEL ORDER BY MODEL")
    
    Set rsVariant = New ADODB.Recordset
    Set rsCount = New ADODB.Recordset
    
    Set xlApplication = CreateObject("Excel.Application")
    Set xlBook = xlApplication.Workbooks.Open(SMIS_REPORT_PATH & "EXCEL\Daily Retail Report.xlt")
    Set xlSheet = xlApplication.Worksheets(1)
        
        While Not rsDealer.EOF
                xlSheet.Cells(I + 6, "A") = Null2String(rsDealer!DEALER_CODE)
                  While Not rsModel.EOF
                        xlSheet.Cells(4, j + 2) = Null2String(rsModel!MODEL)
                        Set rsVariant = gconDMIS.Execute("SELECT DISTINCT DESCRIPT FROM ALL_MODEL  WHERE MODEL='" & Null2String(rsModel!MODEL) & "'")
                                Set rsCount = gconDMIS.Execute("SELECT COUNT(*) AS TCOUNT FROM SMIS_MRRINV_TABLE WHERE DESCRIPT='" & Null2String(rsVariant!DESCRIPT) & "' AND RELEASED=1")
                                While Not rsVariant.EOF
                                    xlSheet.Cells(5, j + 2) = Null2String(rsVariant!DESCRIPT)
                                    'xlSheet.Cells(x + 6, j + 2) = Null2String(rsCount!TCOUNT)
                                    xlSheet.Cells(5, j + 2 + 1) = "SUB TOTAL"
                                    j = j + 1
                                    n = n + 1
                                    X = X + 1
                                    rsVariant.MoveNext
                                Wend
                        j = j + 1
                        rsModel.MoveNext
                  Wend
                I = I + 1
                rsDealer.MoveNext
        Wend
    xlApplication.Visible = True
    Set xlApplication = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
End Sub
Sub MonthlyRetailReport()
    Dim xlApplication As Excel.Application
    Dim xlBook        As Excel.Workbook
    Dim xlSheet       As Excel.Worksheet
    Dim rsDealer      As ADODB.Recordset
    Dim rsModel       As ADODB.Recordset
    Dim rsVariant     As ADODB.Recordset
    Dim rsCount       As ADODB.Recordset
    Dim I             As Integer
    Dim j             As Integer
    Dim L             As Integer
    Dim X             As Integer
    Dim n             As Integer
    Set rsDealer = New ADODB.Recordset
    Set rsDealer = gconDMIS.Execute("SELECT DEALER_CODE FROM ALL_DEALERS ORDER BY PROVINCIAL")

    Set rsModel = New ADODB.Recordset
    Set rsModel = gconDMIS.Execute("SELECT DISTINCT MODEL FROM ALL_MODEL ORDER BY MODEL")
    
    Set rsVariant = New ADODB.Recordset
    Set rsCount = New ADODB.Recordset
    
    Set xlApplication = CreateObject("Excel.Application")
    Set xlBook = xlApplication.Workbooks.Open(SMIS_REPORT_PATH & "EXCEL\Monthly Retail Report.xlt")
    Set xlSheet = xlApplication.Worksheets(1)
        
        While Not rsDealer.EOF
                xlSheet.Cells(I + 6, "A") = Null2String(rsDealer!DEALER_CODE)
                  While Not rsModel.EOF
                        xlSheet.Cells(4, j + 2) = Null2String(rsModel!MODEL)
                        Set rsVariant = gconDMIS.Execute("SELECT DISTINCT DESCRIPT FROM ALL_MODEL  WHERE MODEL='" & Null2String(rsModel!MODEL) & "'")
                                Set rsCount = gconDMIS.Execute("SELECT COUNT(*) AS TCOUNT FROM SMIS_MRRINV_TABLE WHERE DESCRIPT='" & Null2String(rsVariant!DESCRIPT) & "' AND RELEASED=1")
                                While Not rsVariant.EOF
                                    xlSheet.Cells(5, j + 2) = Null2String(rsVariant!DESCRIPT)
                                    'xlSheet.Cells(x + 6, j + 2) = Null2String(rsCount!TCOUNT)
                                    xlSheet.Cells(5, j + 2 + 1) = "SUB TOTAL"
                                    j = j + 1
                                    n = n + 1
                                    X = X + 1
                                    rsVariant.MoveNext
                                Wend
                        j = j + 1
                        rsModel.MoveNext
                  Wend
                I = I + 1
                rsDealer.MoveNext
        Wend
    xlApplication.Visible = True
    Set xlApplication = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
End Sub
Sub NetworkRetailSalesMonthly()
                Dim rsDealer      As ADODB.Recordset
                Dim rsModel       As ADODB.Recordset
                Dim rsCount       As ADODB.Recordset
                Dim xcol As Integer
                Dim jrow As Integer
                Dim intcol As Integer
                Dim introw As Integer
                Set xlApplication = CreateObject("Excel.Application")
                Set xlBook = xlApplication.Workbooks.Open(SMIS_REPORT_PATH & "EXCEL\NEWORK RETAIL SALES-Monthly.xlt")
                Set xlSheet = xlApplication.Worksheets(1)
                
                Set rsDealer = New ADODB.Recordset
                Set rsDealer = gconDMIS.Execute("SELECT DEALER_CODE,* FROM ALL_DEALERS ORDER BY PROVINCIAL")
                Set rsModel = New ADODB.Recordset
                Set rsModel = gconDMIS.Execute("SELECT DISTINCT(MODEL) FROM ALL_MODEL ORDER BY MODEL")
                xcol = 3
                jrow = 5
                While Not rsDealer.EOF
                        xlSheet.Cells(4, xcol) = Null2String(rsDealer!DEALER_CODE)
                        While Not rsModel.EOF
                                xlSheet.Cells(jrow, "B") = Null2String(rsModel!MODEL)
                                        Set rsCount = gconDMIS.Execute("SELECT COUNT(*) As Tcount FROM SMIS_RETAILSALES WHERE MODEL_CODE='" & Null2String(rsModel!MODEL) & "' AND DEALER_CODE='" & Null2String(rsDealer!DEALER_CODE) & "'")
                                        xlSheet.Cells(introw + 5, intcol + 3) = Null2String(rsCount!TCOUNT)
                                introw = introw + 1
                                jrow = jrow + 1
                                rsModel.MoveNext
                        
                        Wend
                xcol = xcol + 1
                intcol = intcol + 1
                rsDealer.MoveNext
                Wend
                xlApplication.Visible = True
                Set xlApplication = Nothing
                Set xlBook = Nothing
                Set xlSheet = Nothing

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

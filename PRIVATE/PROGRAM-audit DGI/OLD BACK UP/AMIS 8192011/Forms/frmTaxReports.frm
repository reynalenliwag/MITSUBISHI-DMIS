VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmTaxReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAX REPORTS"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15000
   Icon            =   "frmTaxReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   15000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8460
      Top             =   3690
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7755
      Left            =   0
      ScaleHeight     =   7755
      ScaleWidth      =   15015
      TabIndex        =   0
      Top             =   30
      Width           =   15015
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Export"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   13680
         MouseIcon       =   "frmTaxReports.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmTaxReports.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Export"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   735
         Left            =   12960
         MouseIcon       =   "frmTaxReports.frx":2256
         MousePointer    =   99  'Custom
         Picture         =   "frmTaxReports.frx":23A8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print Report"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdInquire 
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12240
         MouseIcon       =   "frmTaxReports.frx":2847
         MousePointer    =   99  'Custom
         Picture         =   "frmTaxReports.frx":2999
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "View"
         Top             =   120
         Width           =   735
      End
      Begin VB.ComboBox cboATC 
         Height          =   315
         ItemData        =   "frmTaxReports.frx":2CE0
         Left            =   4950
         List            =   "frmTaxReports.frx":2CE2
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   450
         Width           =   2445
      End
      Begin VB.ComboBox cboOption 
         Height          =   315
         ItemData        =   "frmTaxReports.frx":2CE4
         Left            =   90
         List            =   "frmTaxReports.frx":2CE6
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   450
         Width           =   4815
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Left            =   5850
         ScaleHeight     =   435
         ScaleWidth      =   2535
         TabIndex        =   1
         Top             =   3600
         Visible         =   0   'False
         Width           =   2595
         Begin VB.Label lblLoading 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please wait while loading"
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
            Height          =   225
            Left            =   90
            TabIndex        =   15
            Top             =   120
            Width           =   2145
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   435
            Left            =   0
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   2535
            _Version        =   655364
            _ExtentX        =   4471
            _ExtentY        =   767
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
            SubItemCaption  =   -1  'True
            RightToLeftReading=   -1  'True
         End
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   375
         Left            =   8070
         TabIndex        =   7
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         _Version        =   393216
         Format          =   48955393
         CurrentDate     =   40136
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   375
         Left            =   10110
         TabIndex        =   8
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         _Version        =   393216
         Format          =   48955393
         CurrentDate     =   40136
      End
      Begin FlexCell.Grid Grid1 
         Height          =   6765
         Left            =   30
         TabIndex        =   9
         Top             =   960
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   11933
         Appearance      =   0
         BackColor2      =   16573135
         BackColorBkg    =   -2147483645
         Cols            =   5
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         Rows            =   1
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Report Option"
         Height          =   285
         Left            =   90
         TabIndex        =   14
         Top             =   180
         Width           =   4755
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ATC Description"
         Height          =   285
         Left            =   4950
         TabIndex        =   13
         Top             =   180
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   285
         Left            =   7500
         TabIndex        =   12
         Top             =   510
         Width           =   795
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   285
         Left            =   9780
         TabIndex        =   11
         Top             =   510
         Width           =   795
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FAF1DC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00F5D8BC&
         FillStyle       =   0  'Solid
         Height          =   885
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   30
         Width           =   14925
      End
   End
   Begin VB.Menu mnuExport 
      Caption         =   "Export"
      Visible         =   0   'False
      Begin VB.Menu mnuExcel 
         Caption         =   "Export to Excel"
      End
      Begin VB.Menu mnuPDF 
         Caption         =   "Export to PDF"
      End
      Begin VB.Menu mnuHTML 
         Caption         =   "Export to HTML"
      End
   End
End
Attribute VB_Name = "frmTaxReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPORTS                                          As ADODB.Recordset
Dim i                                                  As Integer
Dim ACCT_CODE                                          As String
Dim TAXREPORT                                          As String
Dim DESCRIPTION                                        As String
Dim CMD                                                As ADODB.Command
Dim xlsWorkSheet                                       As Excel.Worksheet

Sub INIT_CBO_ATC()
'    WITHHOLDING TAX PAYABLE - COMPENSATION
'    WITHHOLDING TAX PAYABLE - EXPANDED
'    INPUT TAX
'    OUTPUT TAX
    Dim rsTAX                                          As ADODB.Recordset
    Set rsTAX = New ADODB.Recordset
    rsTAX.Open "SELECT UPPER(DESCRIPTION) AS DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 IN ('INPUT TAX','OUTPUT TAX','EXPANDED','CREDITABLE','COMPENSATION')", gconDMIS, adOpenKeyset
    If Not rsTAX.EOF And Not rsTAX.BOF Then
        Do While Not rsTAX.EOF
            cboOption.AddItem Null2String(rsTAX!DESCRIPTION)
            rsTAX.MoveNext
        Loop
    End If
    Set rsTAX = Nothing
        cboOption.AddItem "ORB LIST OF PURCHASES"
        cboOption.AddItem "ORB LIST OF SALES"

    Dim rsGET_ATC_DESC                                 As ADODB.Recordset
    Set rsGET_ATC_DESC = New ADODB.Recordset
    rsGET_ATC_DESC.Open "SELECT NATURE FROM AMIS_ATC ", gconDMIS, adOpenKeyset
    If Not rsGET_ATC_DESC.EOF And Not rsGET_ATC_DESC.BOF Then
        cboATC.AddItem "ALL ATC"
        Do While Not rsGET_ATC_DESC.EOF
            cboATC.AddItem Null2String(rsGET_ATC_DESC!NATURE)
            rsGET_ATC_DESC.MoveNext
        Loop
    End If
    Set rsGET_ATC_DESC = Nothing
End Sub

Private Sub cboOption_Click()
    initGrid
    If GET_TRANTYPE(cboOption.Text) = "INPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "OUTPUT TAX" Then
        cboATC.Enabled = False
    ElseIf GET_TRANTYPE(cboOption.Text) = "EXPANDED" Or GET_TRANTYPE(cboOption.Text) = "CREDITABLE" Or GET_TRANTYPE(cboOption.Text) = "COMPENSATION" Then
        cboATC.Enabled = True
    Else
        cboATC.Enabled = False
    End If
End Sub

Private Sub cmdExport_Click()
    PopupMenu mnuExport
End Sub

Private Sub cmdInquire_Click()
    If cboOption.Text = "" Then
        MsgBox "Select type of tax report.", vbInformation, "Type of Report"
        cboOption.SetFocus
        Exit Sub
    ElseIf cboATC.Text = "" And GET_TRANTYPE(cboOption.Text) = "EXPANDED" Then
        MsgBox "Select ATC.", vbInformation, "ATC CODE"
        cboATC.Enabled = True
        cboATC.SetFocus
        Exit Sub
    ElseIf cboATC.Text = "" And GET_TRANTYPE(cboOption.Text) = "CREDITABLE" Then
        MsgBox "Select ATC.", vbInformation, "ATC CODE"
        cboATC.Enabled = True
        cboATC.SetFocus
        Exit Sub
    ElseIf cboATC.Text = "" And GET_TRANTYPE(cboOption.Text) = "COMPENSATION" Then
        MsgBox "Select ATC.", vbInformation, "ATC CODE"
        cboATC.Enabled = True
        cboATC.SetFocus
        Exit Sub
    End If
    If dtTO.Value < dtFrom.Value Then
        MsgBox "Please check date selected.", vbInformation, "Date Range"
        Exit Sub
    End If
    Dim rsACCT_CODE                                    As ADODB.Recordset
    Set rsACCT_CODE = New ADODB.Recordset
    rsACCT_CODE.Open "SELECT ACCTCODE,TRANTYPE1,DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & cboOption.Text & "'", gconDMIS, adOpenForwardOnly
    If Not rsACCT_CODE.EOF And Not rsACCT_CODE.BOF Then
        ACCT_CODE = Null2String(rsACCT_CODE!AcctCode)
        TAXREPORT = Null2String(rsACCT_CODE!TRANTYPE1)
        DESCRIPTION = Null2String(rsACCT_CODE!DESCRIPTION)
    End If
    Set rsACCT_CODE = Nothing
    If cboOption.Text = "ORB LIST OF PURCHASES" Then
        TAXREPORT = "SLP"
    ElseIf cboOption.Text = "ORB LIST OF SALES" Then
        TAXREPORT = "SLS"
    End If
    TAX_REPORTS
End Sub

Private Sub cmdPrint_Click()
'    On Error GoTo ErrorHandler
'    Grid1.PrintPreview 100
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbExclamation
'    Resume Next

'    On Error GoTo ErrorHandler
'    Grid1.DirectPrint
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbExclamation
'    Resume Next
'    If Grid1.ExportToExcel("") Then
'        MsgBox "OK", vbExclamation
'    End If
'SetPrinting
'    Grid1.ExportToExcel ("")
    Dim xlsApplication                                 As Excel.Application
    Dim xlsWorkbook                                    As Excel.Workbook
    Dim ellaine                                        As Integer
    Set xlsApplication = New Excel.Application
    If GET_TRANTYPE(cboOption.Text) = "INPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "OUTPUT TAX" Then
        If Len(Dir(AMIS_REPORT_PATH & "TAX.xlt")) = 0 Then
            MsgBox "Report file cannot be found.", vbInformation
            Exit Sub
        End If
        Set xlsWorkbook = xlsApplication.Workbooks.Open(AMIS_REPORT_PATH & "TAX.xlt")
        Set xlsWorkSheet = xlsWorkbook.Worksheets(1)
        EXCEL_HEADER
        INPUT_OUTPUT_TAX
        Picture1.Enabled = False
        Picture2.Visible = True
        Set rsREPORTS = CMD.Execute
        If Not rsREPORTS.EOF And Not rsREPORTS.BOF Then
            'xlsWorkSheet.Cells(7, 1).CopyFromRecordset rsREPORTS
            Do While Not rsREPORTS.EOF
                xlsWorkSheet.Cells(9 + ellaine, "A") = Null2String(rsREPORTS!JDATE)
                xlsWorkSheet.Cells(9 + ellaine, "B") = Null2String(rsREPORTS!JOURNALNO)
                xlsWorkSheet.Cells(9 + ellaine, "C") = Null2String(rsREPORTS!VendorCode)
                xlsWorkSheet.Cells(9 + ellaine, "D") = Null2String(rsREPORTS!VendorName)
                xlsWorkSheet.Cells(9 + ellaine, "E") = Null2String(rsREPORTS!Address)
                xlsWorkSheet.Cells(9 + ellaine, "F") = Null2String(rsREPORTS!CITY)
                xlsWorkSheet.Cells(9 + ellaine, "G") = Null2String(N2Str2Null(rsREPORTS!TIN))
                xlsWorkSheet.Cells(9 + ellaine, "H") = ToDoubleNumber(rsREPORTS!GROSS)
                xlsWorkSheet.Cells(9 + ellaine, "I") = ToDoubleNumber(rsREPORTS!NET)
                xlsWorkSheet.Cells(9 + ellaine, "J") = ToDoubleNumber(rsREPORTS!VAT)
                xlsWorkSheet.Cells(9 + ellaine, "K") = ToDoubleNumber(rsREPORTS!RUNNINGBALANCE)
                rsREPORTS.MoveNext
                Loading
                ellaine = ellaine + 1
            Loop
        End If
    ElseIf GET_TRANTYPE(cboOption.Text) = "EXPANDED" Or GET_TRANTYPE(cboOption.Text) = "CREDITABLE" Or GET_TRANTYPE(cboOption.Text) = "COMPENSATION" Then
        If Len(Dir(AMIS_REPORT_PATH & "WITHHOLDINGTAX.xlt")) = 0 Then
            MsgBox "Report file cannot be found.", vbInformation
            Exit Sub
        End If
        Set xlsWorkbook = xlsApplication.Workbooks.Open(AMIS_REPORT_PATH & "WITHHOLDINGTAX.xlt")
        Set xlsWorkSheet = xlsWorkbook.Worksheets(1)
        EXCEL_HEADER
        EXPANDED_CREDITABLE_TAX
        Picture1.Enabled = False
        Picture2.Visible = True
        Set rsREPORTS = CMD.Execute
        If Not rsREPORTS.EOF And Not rsREPORTS.BOF Then
            'xlsWorkSheet.Cells(7, 1).CopyFromRecordset rsREPORTS
            Do While Not rsREPORTS.EOF
                xlsWorkSheet.Cells(9 + ellaine, "A") = Null2String(rsREPORTS!JDATE)
                xlsWorkSheet.Cells(9 + ellaine, "B") = Null2String(rsREPORTS!JOURNALNO)
                xlsWorkSheet.Cells(9 + ellaine, "C") = Null2String(rsREPORTS!VendorCode)
                xlsWorkSheet.Cells(9 + ellaine, "D") = Null2String(rsREPORTS!VendorName)
                xlsWorkSheet.Cells(9 + ellaine, "E") = Null2String(rsREPORTS!Address)
                xlsWorkSheet.Cells(9 + ellaine, "F") = Null2String(rsREPORTS!TIN)
                xlsWorkSheet.Cells(9 + ellaine, "G") = Null2String(rsREPORTS!ATC)
                xlsWorkSheet.Cells(9 + ellaine, "H") = Null2String(rsREPORTS!NATURE)
                xlsWorkSheet.Cells(9 + ellaine, "I") = ToDoubleNumber(rsREPORTS!TAXBASE)
                xlsWorkSheet.Cells(9 + ellaine, "J") = ToDoubleNumber(rsREPORTS!Rate)
                xlsWorkSheet.Cells(9 + ellaine, "K") = ToDoubleNumber(rsREPORTS!TAXWITHHELD)
                xlsWorkSheet.Cells(9 + ellaine, "L") = ToDoubleNumber(rsREPORTS!RUNNINGBALANCE)
                rsREPORTS.MoveNext
                Loading
                ellaine = ellaine + 1
            Loop
        End If
    ElseIf cboOption.Text = "ORB LIST OF PURCHASES" Then
        If Len(Dir(AMIS_REPORT_PATH & "ORBListofPuchases.xlt")) = 0 Then
            MsgBox "Report file cannot be found.", vbInformation
            Exit Sub
        End If
        Set xlsWorkbook = xlsApplication.Workbooks.Open(AMIS_REPORT_PATH & "ORBListofPuchases.xlt")
        Set xlsWorkSheet = xlsWorkbook.Worksheets(1)
        ORB_HEADER
        Picture1.Enabled = False
        Picture2.Visible = True
        Set rsREPORTS = CMD.Execute
        If Not rsREPORTS.EOF And Not rsREPORTS.BOF Then
            'xlsWorkSheet.Cells(7, 1).CopyFromRecordset rsREPORTS
            Do While Not rsREPORTS.EOF
                xlsWorkSheet.Cells(15 + ellaine, "A") = Null2String(rsREPORTS!invoicedate)
                xlsWorkSheet.Cells(15 + ellaine, "B") = Null2String(rsREPORTS!INVOICENO)
                xlsWorkSheet.Cells(15 + ellaine, "C") = Null2String(rsREPORTS!NAMEOFENTITY)
                xlsWorkSheet.Cells(15 + ellaine, "D") = Null2String(rsREPORTS!Address)
                xlsWorkSheet.Cells(15 + ellaine, "E") = Null2String(rsREPORTS!Model)
                xlsWorkSheet.Cells(15 + ellaine, "F") = Null2String(rsREPORTS!NO_UNITS)
                xlsWorkSheet.Cells(15 + ellaine, "G") = Null2String(rsREPORTS!VINO)
                xlsWorkSheet.Cells(15 + ellaine, "H") = Null2String(rsREPORTS!ENGINENO)
                rsREPORTS.MoveNext
                Loading
                ellaine = ellaine + 1
            Loop
        End If
    ElseIf cboOption.Text = "ORB LIST OF SALES" Then
        If Len(Dir(AMIS_REPORT_PATH & "ORBListofSales.xlt")) = 0 Then
            MsgBox "Report file cannot be found.", vbInformation
            Exit Sub
        End If
        Set xlsWorkbook = xlsApplication.Workbooks.Open(AMIS_REPORT_PATH & "ORBListofSales.xlt")
        Set xlsWorkSheet = xlsWorkbook.Worksheets(1)
        ORB_HEADER
        Picture1.Enabled = False
        Picture2.Visible = True
        Set rsREPORTS = CMD.Execute
        If Not rsREPORTS.EOF And Not rsREPORTS.BOF Then
            'xlsWorkSheet.Cells(7, 1).CopyFromRecordset rsREPORTS
            Do While Not rsREPORTS.EOF
                xlsWorkSheet.Cells(15 + ellaine, "A") = Null2String(rsREPORTS!invoicedate)
                xlsWorkSheet.Cells(15 + ellaine, "B") = Null2String(rsREPORTS!INVOICENO)
                xlsWorkSheet.Cells(15 + ellaine, "C") = Null2String(rsREPORTS!NAMEOFENTITY)
                xlsWorkSheet.Cells(15 + ellaine, "D") = Null2String(rsREPORTS!Address)
                xlsWorkSheet.Cells(15 + ellaine, "E") = Null2String(rsREPORTS!Model)
                xlsWorkSheet.Cells(15 + ellaine, "F") = Null2String(rsREPORTS!NO_UNITS)
                xlsWorkSheet.Cells(15 + ellaine, "G") = Null2String(rsREPORTS!FRAMENO)
                xlsWorkSheet.Cells(15 + ellaine, "H") = Null2String(rsREPORTS!ENGINENO)
                xlsWorkSheet.Cells(15 + ellaine, "I") = Null2String(rsREPORTS!SELLINGPRICE)
                rsREPORTS.MoveNext
                Loading
                ellaine = ellaine + 1
            Loop
        End If
    End If
    Picture1.Enabled = True
    Picture2.Visible = False
    xlsApplication.Visible = True
    Set xlsApplication = Nothing
    Set xlsWorkbook = Nothing
    Set xlsWorkSheet = Nothing
    Set rsREPORTS = Nothing
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    INIT_CBO_ATC
    'initGrid
    Dim rsMAX_MIN_DATE                                 As ADODB.Recordset
    Set rsMAX_MIN_DATE = New ADODB.Recordset
    rsMAX_MIN_DATE.Open "SELECT * FROM (SELECT MAX(JDATE) AS MAX_JDATE, MIN(JDATE) AS MIN_JDATE FROM AMIS_JOURNAL_HD WHERE STATUS = 'P')T WHERE MAX_JDATE IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsMAX_MIN_DATE.EOF And Not rsMAX_MIN_DATE.BOF Then
        dtFrom.Value = rsMAX_MIN_DATE!MIN_JDATE
        dtTO.Value = rsMAX_MIN_DATE!MAX_JDATE
    Else
        dtFrom.Value = LOGDATE
        dtTO.Value = LOGDATE
    End If
    Set rsMAX_MIN_DATE = Nothing
    Screen.MousePointer = 0
    Grid1.Rows = 1
    'USP_TAX
End Sub

Sub initGrid()
    With Grid1
        If GET_TRANTYPE(cboOption.Text) = "INPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "OUTPUT TAX" Then
            .Cols = 12
        ElseIf GET_TRANTYPE(cboOption.Text) = "EXPANDED" Or GET_TRANTYPE(cboOption.Text) = "COMPENSATION" Or GET_TRANTYPE(cboOption.Text) = "CREDITABLE" Then
            .Cols = 13
        ElseIf cboOption.Text = "ORB LIST OF PURCHASES" Then
            .Cols = 9
        ElseIf cboOption.Text = "ORB LIST OF SALES" Then
            .Cols = 10
        End If
        .Rows = 1
        '        .FixedCols = 4

        .Cell(0, 0).Text = "L/N"
        If GET_TRANTYPE(cboOption.Text) = "INPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "OUTPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "EXPANDED" Or GET_TRANTYPE(cboOption.Text) = "COMPENSATION" Or GET_TRANTYPE(cboOption.Text) = "CREDITABLE" Then
            .Cell(0, 1).Text = "DATE"
            .Column(1).Width = 65
            .Column(1).FormatString = "mm/dd/yyyy"
    
            .Cell(0, 2).Text = "JOURNAL NO."
            .Column(2).Alignment = cellCenterCenter
            .Column(2).Width = 75
        End If
        
        If GET_TRANTYPE(cboOption.Text) = "INPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "EXPANDED" Or GET_TRANTYPE(cboOption.Text) = "COMPENSATION" Then
            .Cell(0, 3).Text = "VEN. CODE"
            .Column(3).Width = 80
            .Column(3).Alignment = cellCenterCenter

            .Cell(0, 4).Text = "VENDOR NAME"
            .Column(4).Width = 200
        ElseIf GET_TRANTYPE(cboOption.Text) = "OUTPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "CREDITABLE" Then
            .Cell(0, 3).Text = "CUST. CODE"
            .Column(3).Width = 80
            .Column(3).Alignment = cellCenterCenter

            .Cell(0, 4).Text = "CUSTOMER NAME"
            .Column(4).Width = 200
        End If

        If GET_TRANTYPE(cboOption.Text) = "INPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "OUTPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "EXPANDED" Or GET_TRANTYPE(cboOption.Text) = "COMPENSATION" Or GET_TRANTYPE(cboOption.Text) = "CREDITABLE" Then
            .Cell(0, 5).Text = "ADDRESS"
            .Column(5).Width = 200
    
            .Cell(0, 6).Text = "CITY"
            .Column(6).Width = 0
    
            .Cell(0, 7).Text = "TIN"
            .Column(7).Alignment = cellCenterCenter
            .Column(7).Width = 110
        End If
        
        If GET_TRANTYPE(cboOption.Text) = "INPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "OUTPUT TAX" Then
            .Cell(0, 8).Text = "GROSS AMOUNT"
            .Column(8).Width = 0
            .Column(8).Alignment = cellRightCenter

            .Cell(0, 9).Text = "NET AMOUNT"
            .Column(9).Width = 0
            .Column(9).Alignment = cellRightCenter

            .Cell(0, 10).Text = "AMOUNT"
            .Column(10).Width = 90
            .Column(10).Alignment = cellRightCenter

            .Cell(0, 11).Text = "RUNNING BALANCE"
            .Column(11).Width = 115
            .Column(11).Alignment = cellRightCenter
        ElseIf GET_TRANTYPE(cboOption.Text) = "EXPANDED" Or GET_TRANTYPE(cboOption.Text) = "CREDITABLE" Or GET_TRANTYPE(cboOption.Text) = "COMPENSATION" Then
            .Cell(0, 7).Text = "ATC CODE"
            .Column(7).Width = 90
            .Column(7).Alignment = cellCenterCenter

            .Cell(0, 8).Text = "NATURE OF PAYMENT"
            .Column(8).Width = 100
            .Column(8).Alignment = cellLeftCenter

            .Cell(0, 9).Text = "AMOUNT TAXBASE"
            .Column(9).Width = 90
            .Column(9).Alignment = cellRightCenter

            .Cell(0, 10).Text = "TAX RATE"
            .Column(10).Width = 90
            .Column(10).Alignment = cellRightCenter

            .Cell(0, 11).Text = "TAX WITHHELD"
            .Column(11).Width = 115
            .Column(11).Alignment = cellRightCenter

            .Cell(0, 12).Text = "RUNNING BALANCE"
            .Column(12).Width = 115
            .Column(12).Alignment = cellRightCenter
        End If
        .RowHeight(0) = 35
        If cboOption.Text = "ORB LIST OF PURCHASES" Then
            .Cell(0, 1).Text = "DATE"
            .Column(1).Width = 65
            .Column(1).FormatString = "mm/dd/yyyy"
    
            .Cell(0, 2).Text = "INVOICE NO."
            .Column(2).Alignment = cellCenterCenter
            .Column(2).Width = 75
            
            .Cell(0, 3).Text = "NAME"
            .Column(3).Width = 200
            
            .Cell(0, 4).Text = "ADRESS"
            .Column(4).Width = 200
            
            .Cell(0, 5).Text = "MODEL"
            .Column(5).Width = 110
            
            .Cell(0, 6).Text = "NO. OF UNITS"
            .Column(6).Width = 80
            .Column(6).Alignment = cellCenterCenter
            
            .Cell(0, 7).Text = "VIN"
            .Column(7).Width = 110
            
            .Cell(0, 8).Text = "ENGINE NO."
            .Column(8).Width = 110
        ElseIf cboOption.Text = "ORB LIST OF SALES" Then
            .Cell(0, 1).Text = "DATE"
            .Column(1).Width = 65
            .Column(1).FormatString = "mm/dd/yyyy"
    
            .Cell(0, 2).Text = "INVOICE NO."
            .Column(2).Alignment = cellCenterCenter
            .Column(2).Width = 75
            
            .Cell(0, 3).Text = "NAME"
            .Column(3).Width = 200
            
            .Cell(0, 4).Text = "ADRESS"
            .Column(4).Width = 200
            
            .Cell(0, 5).Text = "MODEL"
            .Column(5).Width = 110
            
            .Cell(0, 6).Text = "NO. OF UNITS"
            .Column(6).Width = 80
            .Column(6).Alignment = cellCenterCenter
            
            .Cell(0, 7).Text = "VIN"
            .Column(7).Width = 110
            
            .Cell(0, 8).Text = "ENGINE NO."
            .Column(8).Width = 110
            
            .Cell(0, 9).Text = "NET OF VAT"
            .Column(9).Width = 110
            .Column(9).Alignment = cellRightCenter
        End If
        
        If GET_TRANTYPE(cboOption.Text) = "INPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "OUTPUT TAX" Then
            .Range(0, 0, 0, 11).WrapText = True
            .Range(0, 0, 0, 11).ForeColor = vbBlue
        ElseIf GET_TRANTYPE(cboOption.Text) = "EXPANDED" Or GET_TRANTYPE(cboOption.Text) = "CREDITABLE" Or GET_TRANTYPE(cboOption.Text) = "COMPENSATION" Then
            .Range(0, 0, 0, 12).WrapText = True
            .Range(0, 0, 0, 12).ForeColor = vbBlue
        End If
        
        For i = 0 To Grid1.Cols - 1
            Grid1.Column(i).Locked = True
        Next
    End With
End Sub

Sub TAX_REPORTS()
    Set CMD = New ADODB.Command
    CMD.ActiveConnection = gconDMIS
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "USP_TAXREPORTS"

    With CMD.Parameters
        .Append CMD.CreateParameter("@ACCT_CODE", adVarChar, adParamInput, 12, ACCT_CODE)
        .Append CMD.CreateParameter("@JDATE1", adDate, adParamInput, 8, dtFrom)
        .Append CMD.CreateParameter("@JDATE2", adDate, adParamInput, 8, dtTO)
        .Append CMD.CreateParameter("@TAXREPORT", adVarChar, adParamInput, 25, TAXREPORT)
    End With

    Set rsREPORTS = CMD.Execute
    FILLREPORTS
End Sub

Sub FILLREPORTS()
    Grid1.Rows = 1
    Grid1.AutoRedraw = False
    Picture1.Enabled = False
    Picture2.Visible = True

    If Not rsREPORTS.EOF And Not rsREPORTS.BOF Then
        While Not rsREPORTS.EOF
            If TAXREPORT = "INPUT TAX" Or TAXREPORT = "OUTPUT TAX" Then
                Grid1.AddItem _
                        rsREPORTS!JDATE & Chr(9) & rsREPORTS!JOURNALNO & Chr(9) & _
                                        rsREPORTS!VendorCode & Chr(9) & rsREPORTS!VendorName & Chr(9) & _
                                        Null2String(rsREPORTS!Address) & Chr(9) & Null2String(rsREPORTS!CITY) & Chr(9) & rsREPORTS!TIN & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!GROSS) & Chr(9) & ToDoubleNumber(rsREPORTS!NET) & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!VAT) & Chr(9) & ToDoubleNumber(rsREPORTS!RUNNINGBALANCE)
            ElseIf TAXREPORT = "EXPANDED" Or TAXREPORT = "CREDITABLE" Or TAXREPORT = "COMPENSATION" Then
                Grid1.AddItem _
                        rsREPORTS!JDATE & Chr(9) & rsREPORTS!JOURNALNO & Chr(9) & _
                                        rsREPORTS!VendorCode & Chr(9) & rsREPORTS!VendorName & Chr(9) & _
                                        rsREPORTS!Address & Chr(9) & rsREPORTS!TIN & Chr(9) & _
                                        rsREPORTS!ATC & Chr(9) & rsREPORTS!NATURE & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!TAXBASE) & Chr(9) & N2String(rsREPORTS!Rate) & "%" & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!TAXWITHHELD) & Chr(9) & ToDoubleNumber(rsREPORTS!RUNNINGBALANCE)
            ElseIf TAXREPORT = "SLP" Then
                Grid1.AddItem _
                        rsREPORTS!invoicedate & Chr(9) & rsREPORTS!INVOICENO & Chr(9) & _
                                        rsREPORTS!NAMEOFENTITY & Chr(9) & rsREPORTS!Address & Chr(9) & _
                                        Null2String(rsREPORTS!Model) & Chr(9) & Null2String(rsREPORTS!NO_UNITS) & Chr(9) & rsREPORTS!VINO & Chr(9) & _
                                        rsREPORTS!ENGINENO
            ElseIf TAXREPORT = "SLS" Then
                Grid1.AddItem _
                        rsREPORTS!invoicedate & Chr(9) & rsREPORTS!INVOICENO & Chr(9) & _
                                        rsREPORTS!NAMEOFENTITY & Chr(9) & rsREPORTS!Address & Chr(9) & _
                                        Null2String(rsREPORTS!Model) & Chr(9) & Null2String(rsREPORTS!NO_UNITS) & Chr(9) & rsREPORTS!FRAMENO & Chr(9) & _
                                        rsREPORTS!ENGINENO & Chr(9) & ToDoubleNumber(rsREPORTS!SELLINGPRICE)
            End If
            rsREPORTS.MoveNext
            Loading
        Wend
    End If
    Grid1.AutoRedraw = True
    Grid1.Refresh
    Picture1.Enabled = True
    Picture2.Visible = False
    cmdPrint.Enabled = True
    cmdExport.Enabled = True
    Set rsREPORTS = Nothing
End Sub

Private Sub mnuExcel_Click()
    Grid1.ExportToExcel ("")
End Sub

Private Sub mnuHTML_Click()
    Grid1.ExportToHTML ("")
End Sub

Private Sub mnuPDF_Click()
    Grid1.ExportToPDF ("")
End Sub

Sub EXCEL_HEADER()
    xlsWorkSheet.Cells(1, "A") = "COMPANY NAME :"
    xlsWorkSheet.Cells(1, "A").Font.Bold = True
    xlsWorkSheet.Cells(1, "B") = COMPANY_NAME
    'xlsWorkSheet.Cells(1, "B").Font.Bold = True

    xlsWorkSheet.Cells(2, "A") = "COMPANY_ADDRESS :"
    xlsWorkSheet.Cells(2, "A").Font.Bold = True
    xlsWorkSheet.Cells(2, "B") = COMPANY_ADDRESS
    'xlsWorkSheet.Cells(2, "B").Font.Bold = True

    xlsWorkSheet.Cells(3, "A") = "RUN DATE :"
    xlsWorkSheet.Cells(3, "A").Font.Bold = True
    xlsWorkSheet.Cells(3, "B") = LOGDATE
    xlsWorkSheet.Cells(3, "B").Cells.HorizontalAlignment = xlLeft
    'xlsWorkSheet.Cells(3, "B").Font.Bold = True

    xlsWorkSheet.Cells(5, "A") = "ACCOUNT CODE :"
    xlsWorkSheet.Cells(5, "A").Font.Bold = True
    xlsWorkSheet.Cells(5, "B") = ACCT_CODE
    '    xlsWorkSheet.Cells(5, "B").Font.Bold = True

    xlsWorkSheet.Cells(6, "A") = "ACCOUNT NAME :"
    xlsWorkSheet.Cells(6, "A").Font.Bold = True
    xlsWorkSheet.Cells(6, "B") = DESCRIPTION
    '    xlsWorkSheet.Cells(6, "B").Font.Bold = True

    xlsWorkSheet.Cells(7, "A") = "DATE RANGE :"
    xlsWorkSheet.Cells(7, "A").Font.Bold = True
    xlsWorkSheet.Cells(7, "B") = "From :" & " " & dtFrom.Value & "" & " " & "To :" & " " & dtTO & ""
    'xlsWorkSheet.Cells(7, "B").Font.Bold = True

    xlsWorkSheet.Cells(8, "A") = "DATE"
    xlsWorkSheet.Cells(8, "A").BorderAround ColorIndex:=1, Weight:=xlThin
    xlsWorkSheet.Cells(8, "A").Font.Bold = True
    xlsWorkSheet.Cells(8, "A").Interior.Color = &HF5D8BC
    xlsWorkSheet.Cells(8, "B") = "JOURNAL #"
    xlsWorkSheet.Cells(8, "B").BorderAround ColorIndex:=1, Weight:=xlThin
    xlsWorkSheet.Cells(8, "B").Font.Bold = True
    xlsWorkSheet.Cells(8, "B").Interior.Color = &HF5D8BC
    xlsWorkSheet.Cells(8, "C") = "CODE"
    xlsWorkSheet.Cells(8, "C").BorderAround ColorIndex:=1, Weight:=xlThin
    xlsWorkSheet.Cells(8, "C").Font.Bold = True
    xlsWorkSheet.Cells(8, "C").Interior.Color = &HF5D8BC
End Sub

Sub ORB_HEADER()
    xlsWorkSheet.Cells(9, "C") = COMPANY_NAME
    xlsWorkSheet.Cells(10, "C") = COMPANY_ADDRESS
    xlsWorkSheet.Cells(7, "A") = "For the Month of " & Format(dtTO.Value, "mmmm yyyy")
    xlsWorkSheet.Cells(9, "H") = COMPANY_TIN
End Sub

Sub INPUT_OUTPUT_TAX()
    If GET_TRANTYPE(cboOption.Text) = "OUTPUT TAX" Then
        xlsWorkSheet.Cells(8, "D") = "CUSTOMER NAME"
        xlsWorkSheet.Cells(8, "D").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "D").Font.Bold = True
        xlsWorkSheet.Cells(8, "D").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "E") = "ADDRESS"
        xlsWorkSheet.Cells(8, "E").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "E").Font.Bold = True
        xlsWorkSheet.Cells(8, "E").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "F") = "CITY"
        xlsWorkSheet.Cells(8, "F").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "F").Font.Bold = True
        xlsWorkSheet.Cells(8, "F").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "G") = "TIN #"
        xlsWorkSheet.Cells(8, "G").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "G").Font.Bold = True
        xlsWorkSheet.Cells(8, "G").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "H") = "GROSS AMOUNT"
        xlsWorkSheet.Cells(8, "H").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "H").Font.Bold = True
        xlsWorkSheet.Cells(8, "H").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "I") = "NET OF VAT"
        xlsWorkSheet.Cells(8, "I").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "I").Font.Bold = True
        xlsWorkSheet.Cells(8, "I").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "J") = "VAT"
        xlsWorkSheet.Cells(8, "J").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "J").Font.Bold = True
        xlsWorkSheet.Cells(8, "J").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "K") = "RUNNING-BALANCE"
        xlsWorkSheet.Cells(8, "K").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "K").Font.Bold = True
        xlsWorkSheet.Cells(8, "K").Interior.Color = &HF5D8BC
    Else
        xlsWorkSheet.Cells(8, "D") = "VENDOR NAME"
        xlsWorkSheet.Cells(8, "D").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "D").Font.Bold = True
        xlsWorkSheet.Cells(8, "D").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "E") = "ADDRESS"
        xlsWorkSheet.Cells(8, "E").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "E").Font.Bold = True
        xlsWorkSheet.Cells(8, "E").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "F") = "CITY"
        xlsWorkSheet.Cells(8, "F").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "F").Font.Bold = True
        xlsWorkSheet.Cells(8, "F").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "G") = "TIN #"
        xlsWorkSheet.Cells(8, "G").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "G").Font.Bold = True
        xlsWorkSheet.Cells(8, "G").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "H") = "GROSS AMOUNT"
        xlsWorkSheet.Cells(8, "H").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "H").Font.Bold = True
        xlsWorkSheet.Cells(8, "H").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "I") = "NET OF VAT"
        xlsWorkSheet.Cells(8, "I").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "I").Font.Bold = True
        xlsWorkSheet.Cells(8, "I").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "J") = "VAT"
        xlsWorkSheet.Cells(8, "J").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "J").Font.Bold = True
        xlsWorkSheet.Cells(8, "J").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "K") = "RUNNING-BALANCE"
        xlsWorkSheet.Cells(8, "K").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "K").Font.Bold = True
        xlsWorkSheet.Cells(8, "K").Interior.Color = &HF5D8BC
    End If
End Sub

Sub EXPANDED_CREDITABLE_TAX()
    If GET_TRANTYPE(cboOption.Text) = "CREDITABLE" Then
        xlsWorkSheet.Cells(8, "D") = "CUSTOMER NAME"
        xlsWorkSheet.Cells(8, "D").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "D").Font.Bold = True
        xlsWorkSheet.Cells(8, "D").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "E") = "ADDRESS"
        xlsWorkSheet.Cells(8, "E").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "E").Font.Bold = True
        xlsWorkSheet.Cells(8, "E").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "F") = "TIN"
        xlsWorkSheet.Cells(8, "F").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "F").Font.Bold = True
        xlsWorkSheet.Cells(8, "F").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "G") = "ATC CODE"
        xlsWorkSheet.Cells(8, "G").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "G").Font.Bold = True
        xlsWorkSheet.Cells(8, "G").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "H") = "NATURE OF PAYMENT"
        xlsWorkSheet.Cells(8, "H").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "H").Font.Bold = True
        xlsWorkSheet.Cells(8, "H").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "I") = "AMOUNT TAXBASE"
        xlsWorkSheet.Cells(8, "I").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "I").Font.Bold = True
        xlsWorkSheet.Cells(8, "I").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "J") = "TAX RATE"
        xlsWorkSheet.Cells(8, "J").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "J").Font.Bold = True
        xlsWorkSheet.Cells(8, "J").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "K") = "TAX WITHHELD"
        xlsWorkSheet.Cells(8, "K").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "K").Font.Bold = True
        xlsWorkSheet.Cells(8, "K").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "L") = "RUNNING-BALANCE"
        xlsWorkSheet.Cells(8, "L").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "L").Font.Bold = True
        xlsWorkSheet.Cells(8, "L").Interior.Color = &HF5D8BC
    Else
        xlsWorkSheet.Cells(8, "D") = "VENDOR NAME"
        xlsWorkSheet.Cells(8, "D").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "D").Font.Bold = True
        xlsWorkSheet.Cells(8, "D").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "E") = "ADDRESS"
        xlsWorkSheet.Cells(8, "E").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "E").Font.Bold = True
        xlsWorkSheet.Cells(8, "E").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "F") = "TIN"
        xlsWorkSheet.Cells(8, "F").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "F").Font.Bold = True
        xlsWorkSheet.Cells(8, "F").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "G") = "ATC CODE"
        xlsWorkSheet.Cells(8, "G").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "G").Font.Bold = True
        xlsWorkSheet.Cells(8, "G").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "H") = "NATURE OF PAYMENT"
        xlsWorkSheet.Cells(8, "H").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "H").Font.Bold = True
        xlsWorkSheet.Cells(8, "H").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "I") = "AMOUNT TAXBASE"
        xlsWorkSheet.Cells(8, "I").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "I").Font.Bold = True
        xlsWorkSheet.Cells(8, "I").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "J") = "TAX RATE"
        xlsWorkSheet.Cells(8, "J").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "J").Font.Bold = True
        xlsWorkSheet.Cells(8, "J").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "K") = "TAX WITHHELD"
        xlsWorkSheet.Cells(8, "K").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "K").Font.Bold = True
        xlsWorkSheet.Cells(8, "K").Interior.Color = &HF5D8BC

        xlsWorkSheet.Cells(8, "L") = "RUNNING-BALANCE"
        xlsWorkSheet.Cells(8, "L").BorderAround ColorIndex:=1, Weight:=xlThin
        xlsWorkSheet.Cells(8, "L").Font.Bold = True
        xlsWorkSheet.Cells(8, "L").Interior.Color = &HF5D8BC
    End If
End Sub

Sub Loading()
    If lblLoading.Caption = "Please wait while loading" Then
        lblLoading.Caption = "Please wait while loading."
    ElseIf lblLoading.Caption = "Please wait while loading." Then
        lblLoading.Caption = "Please wait while loading.."
    ElseIf lblLoading.Caption = "Please wait while loading.." Then
        lblLoading.Caption = "Please wait while loading..."
    ElseIf lblLoading.Caption = "Please wait while loading..." Then
        lblLoading.Caption = "Please wait while loading...."
    ElseIf lblLoading.Caption = "Please wait while loading...." Then
        lblLoading.Caption = "Please wait while loading....."
    ElseIf lblLoading.Caption = "Please wait while loading....." Then
        lblLoading.Caption = "Please wait while loading."
    End If
End Sub

Sub SetPrinting()
    With Grid1
        If GET_TRANTYPE(cboOption.Text) = "INPUT TAX" Or GET_TRANTYPE(cboOption.Text) = "OUTPUT TAX" Then
            .Cell(0, 6).Text = "CITY"
            .Column(6).Width = 100

            .Cell(0, 8).Text = "GROSS AMOUNT"
            .Column(8).Width = 115
            .Column(8).Alignment = cellRightCenter

            .Cell(0, 9).Text = "NET AMOUNT"
            .Column(9).Width = 115
            .Column(9).Alignment = cellRightCenter

            .Cell(0, 10).Text = "AMOUNT"
            .Column(10).Width = 115
            .Column(10).Alignment = cellRightCenter

            .Cell(0, 11).Text = "RUNNING BALANCE"
            .Column(11).Width = 115
            .Column(11).Alignment = cellRightCenter
        ElseIf GET_TRANTYPE(cboOption.Text) = "EXPANDED" Or GET_TRANTYPE(cboOption.Text) = "CREDITABLE" Then
            .Cell(0, 7).Text = "ATC CODE"
            .Column(7).Width = 90
            .Column(7).Alignment = cellCenterCenter

            .Cell(0, 8).Text = "NATURE OF PAYMENT"
            .Column(8).Width = 100
            .Column(8).Alignment = cellLeftCenter

            .Cell(0, 9).Text = "AMOUNT TAXBASE"
            .Column(9).Width = 90
            .Column(9).Alignment = cellRightCenter

            .Cell(0, 10).Text = "TAX RATE"
            .Column(10).Width = 90
            .Column(10).Alignment = cellRightCenter

            .Cell(0, 11).Text = "TAX WITHHELD"
            .Column(11).Width = 115
            .Column(11).Alignment = cellRightCenter

            .Cell(0, 12).Text = "RUNNING BALANCE"
            .Column(12).Width = 115
            .Column(12).Alignment = cellRightCenter
        End If
    End With
End Sub

Function GET_TRANTYPE(XXX As String) As String
    Dim rsTranType                                     As ADODB.Recordset
    Set rsTranType = New ADODB.Recordset
    rsTranType.Open "SELECT * FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsTranType.EOF And Not rsTranType.BOF Then
        GET_TRANTYPE = Null2String(rsTranType!TRANTYPE1)
    End If
    Set rsTranType = Nothing
End Function

Sub USP_TAX()
    Dim SQL                                            As String
    SQL = ""
    SQL = "IF EXISTS(SELECT * FROM SYS.OBJECTS WHERE TYPE='P' AND NAME='USP_TAXREPORTS')" & vbCrLf
    SQL = SQL & "DROP PROCEDURE USP_TAXREPORTS"
    gconDMIS.Execute SQL
    SQL = ""
    SQL = "CREATE PROCEDURE [USP_TAXREPORTS]" & vbCrLf
    SQL = SQL & "@ACCT_CODE      NVARCHAR(12)," & vbCrLf
    SQL = SQL & "@JDATE1         SMALLDATETIME," & vbCrLf
    SQL = SQL & "@JDATE2         SMALLDATETIME," & vbCrLf
    SQL = SQL & "@TAXREPORT      NVARCHAR(25)" & vbCrLf
    SQL = SQL & "AS" & vbCrLf
    SQL = SQL & "BEGIN" & vbCrLf
    SQL = SQL & "SET NOCOUNT ON" & vbCrLf
    SQL = SQL & "DECLARE @TAXREPORT_TABLE TABLE" & vbCrLf
    SQL = SQL & "(" & vbCrLf
    SQL = SQL & "JDATE           SMALLDATETIME," & vbCrLf
    SQL = SQL & "JOURNALNO       NVARCHAR(20)," & vbCrLf
    SQL = SQL & "VENDORCODE      NVARCHAR(100)," & vbCrLf
    SQL = SQL & "VENDORNAME      NVARCHAR(100)," & vbCrLf
    SQL = SQL & "ADDRESS         NVARCHAR(500)," & vbCrLf
    SQL = SQL & "CITY            NVARCHAR(300)," & vbCrLf
    SQL = SQL & "TIN             NVARCHAR(100)," & vbCrLf
    SQL = SQL & "ATC             NVARCHAR(50)," & vbCrLf
    SQL = SQL & "NATURE          NVARCHAR(100)," & vbCrLf
    SQL = SQL & "TAXBASE         DECIMAL(18,2)," & vbCrLf
    SQL = SQL & "RATE            DECIMAL(18,2)," & vbCrLf
    SQL = SQL & "TAXWITHHELD     DECIMAL(18,2)," & vbCrLf
    SQL = SQL & "GROSS           DECIMAL(18,2)," & vbCrLf
    SQL = SQL & "NET             DECIMAL(18,2)," & vbCrLf
    SQL = SQL & "VAT             DECIMAL(18,2)," & vbCrLf
    SQL = SQL & "RUNNINGBALANCE  DECIMAL(18,2)" & vbCrLf
    SQL = SQL & ")" & vbCrLf
    SQL = SQL & "DECLARE @JDATE          SMALLDATETIME" & vbCrLf
    SQL = SQL & "DECLARE @JOURNALNO      NVARCHAR(20)" & vbCrLf
    SQL = SQL & "DECLARE @VENDORCODE     NVARCHAR(20)" & vbCrLf
    SQL = SQL & "DECLARE @VENDORNAME     NVARCHAR(200)" & vbCrLf
    SQL = SQL & "DECLARE @ADDRESS        NVARCHAR(500)" & vbCrLf
    SQL = SQL & "DECLARE @ADDRESS2       NVARCHAR(500)" & vbCrLf
    SQL = SQL & "DECLARE @TIN            NVARCHAR(50)" & vbCrLf
    SQL = SQL & "DECLARE @GROSS          DECIMAL(18,2)" & vbCrLf
    SQL = SQL & "DECLARE @NET            DECIMAL(18,2)" & vbCrLf
    SQL = SQL & "DECLARE @AMOUNT         DECIMAL(18,2)" & vbCrLf
    SQL = SQL & "DECLARE @RUNNINGBALANCE DECIMAL(18,2)" & vbCrLf
    SQL = SQL & "DECLARE @ATC            NVARCHAR(100)" & vbCrLf
    SQL = SQL & "DECLARE @NATURE         NVARCHAR(100)" & vbCrLf
    SQL = SQL & "DECLARE @TAXBASE        DECIMAL(18,2)" & vbCrLf
    SQL = SQL & "DECLARE @RATE           DECIMAL(18,2)" & vbCrLf
    SQL = SQL & "SET @RUNNINGBALANCE=0" & vbCrLf
    SQL = SQL & "DECLARE @TAX_REPORTS CURSOR" & vbCrLf
    SQL = SQL & "IF @TAXREPORT = 'INPUT TAX'" & vbCrLf
    SQL = SQL & "BEGIN" & vbCrLf
    SQL = SQL & "SET @TAX_REPORTS = CURSOR FOR" & vbCrLf
    SQL = SQL & "SELECT HD.JDATE,HD.JTYPE+'-'+HD.VOUCHERNO JOURNALNO," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "HD.VendorCode" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "HD.VendorCode" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "RIGHT(ISNULL(DT.ENTITY,''),6)" & vbCrLf
    SQL = SQL & "END AS VENDORCODE," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS VENDORNAME," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS ADDRESS," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS2 FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS2 FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS2 FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS ADDRESS2," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(DT.ENTITY,6))" & vbCrLf
    SQL = SQL & "END AS TIN," & vbCrLf
    SQL = SQL & "HD.AMOUNTTOPAY AS GROSS," & vbCrLf
    SQL = SQL & "HD.AMOUNTTOPAY-DT.DEBIT AS NET," & vbCrLf
    SQL = SQL & "DT.DEBIT" & vbCrLf
    SQL = SQL & "FROM AMIS_JOURNAL_HD HD" & vbCrLf
    SQL = SQL & "INNER JOIN AMIS_JOURNAL_DET DT" & vbCrLf
    SQL = SQL & "ON HD.VOUCHERNO=DT.VOUCHERNO AND HD.JTYPE=DT.JTYPE" & vbCrLf
    SQL = SQL & "WHERE HD.STATUS='P' AND DT.ACCT_CODE=@ACCT_CODE AND DT.DEBIT > 0" & vbCrLf
    SQL = SQL & "AND HD.JDATE BETWEEN @JDATE1 AND @JDATE2" & vbCrLf
    SQL = SQL & "ORDER BY HD.JDATE" & vbCrLf
    SQL = SQL & "OPEN @TAX_REPORTS" & vbCrLf
    SQL = SQL & "FETCH NEXT FROM @TAX_REPORTS INTO @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@GROSS,@NET,@AMOUNT" & vbCrLf
    SQL = SQL & "SET @RUNNINGBALANCE=@AMOUNT" & vbCrLf
    SQL = SQL & "WHILE @@FETCH_STATUS=0" & vbCrLf
    SQL = SQL & "BEGIN" & vbCrLf
    SQL = SQL & "INSERT INTO @TAXREPORT_TABLE(JDATE,JOURNALNO,VENDORCODE,VENDORNAME,ADDRESS,CITY,TIN,GROSS,NET,VAT,RUNNINGBALANCE)" & vbCrLf
    SQL = SQL & "SELECT @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@GROSS,@NET,@AMOUNT,@RUNNINGBALANCE" & vbCrLf
    SQL = SQL & "FETCH NEXT FROM @TAX_REPORTS INTO @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@GROSS,@NET,@AMOUNT" & vbCrLf
    SQL = SQL & "SET @RUNNINGBALANCE=@RUNNINGBALANCE+@AMOUNT" & vbCrLf
    SQL = SQL & "End" & vbCrLf
    SQL = SQL & "CLOSE @TAX_REPORTS" & vbCrLf
    SQL = SQL & "DEALLOCATE @TAX_REPORTS" & vbCrLf
    SQL = SQL & "SELECT * FROM @TAXREPORT_TABLE ORDER BY JDATE" & vbCrLf
    SQL = SQL & "End" & vbCrLf
    SQL = SQL & "ELSE IF @TAXREPORT = 'OUTPUT TAX'" & vbCrLf
    SQL = SQL & "BEGIN" & vbCrLf
    SQL = SQL & "SET @TAX_REPORTS = CURSOR FOR" & vbCrLf
    SQL = SQL & "SELECT HD.JDATE,HD.JTYPE+'-'+HD.VOUCHERNO JOURNALNO," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'SJ' THEN" & vbCrLf
    SQL = SQL & "HD.CustomerCode" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CRJ' THEN" & vbCrLf
    SQL = SQL & "HD.CustomerCode" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE='GJ' THEN" & vbCrLf
    SQL = SQL & "RIGHT(DT.ENTITY,6)" & vbCrLf
    SQL = SQL & "END AS CUSTOMERCODE," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'SJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CRJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE='GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS CUSTOMERNAME," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'SJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT CUSTOMERADD FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CRJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT CUSTOMERADD FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT CUSTOMERADD FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS ADDRESS," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'SJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT CITY FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CRJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT CITY FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT CITY FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS ADDRESS2," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE='SJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE='CRJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS TIN," & vbCrLf
    SQL = SQL & "HD.INVOICEAMT AS GROSS," & vbCrLf
    SQL = SQL & "HD.INVOICEAMT - DT.CREDIT AS NET," & vbCrLf
    SQL = SQL & "DT.CREDIT" & vbCrLf
    SQL = SQL & "FROM AMIS_JOURNAL_HD HD" & vbCrLf
    SQL = SQL & "INNER JOIN AMIS_JOURNAL_DET DT" & vbCrLf
    SQL = SQL & "ON HD.VOUCHERNO=DT.VOUCHERNO AND HD.JTYPE=DT.JTYPE" & vbCrLf
    SQL = SQL & "WHERE HD.STATUS='P' AND DT.ACCT_CODE=@ACCT_CODE AND DT.CREDIT > 0" & vbCrLf
    SQL = SQL & "AND HD.JDATE BETWEEN @JDATE1 AND @JDATE2" & vbCrLf
    SQL = SQL & "ORDER BY HD.JDATE" & vbCrLf
    SQL = SQL & "OPEN @TAX_REPORTS" & vbCrLf
    SQL = SQL & "FETCH NEXT FROM @TAX_REPORTS INTO @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@GROSS,@NET,@AMOUNT" & vbCrLf
    SQL = SQL & "SET @RUNNINGBALANCE=@AMOUNT" & vbCrLf
    SQL = SQL & "WHILE @@FETCH_STATUS=0" & vbCrLf
    SQL = SQL & "BEGIN" & vbCrLf
    SQL = SQL & "INSERT INTO @TAXREPORT_TABLE(JDATE,JOURNALNO,VENDORCODE,VENDORNAME,ADDRESS,CITY,TIN,GROSS,NET,VAT,RUNNINGBALANCE)" & vbCrLf
    SQL = SQL & "SELECT @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@GROSS,@NET,@AMOUNT,@RUNNINGBALANCE" & vbCrLf
    SQL = SQL & "FETCH NEXT FROM @TAX_REPORTS INTO @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@GROSS,@NET,@AMOUNT" & vbCrLf
    SQL = SQL & "SET @RUNNINGBALANCE=@RUNNINGBALANCE+@AMOUNT" & vbCrLf
    SQL = SQL & "End" & vbCrLf
    SQL = SQL & "CLOSE @TAX_REPORTS" & vbCrLf
    SQL = SQL & "DEALLOCATE @TAX_REPORTS" & vbCrLf
    SQL = SQL & "SELECT * FROM @TAXREPORT_TABLE ORDER BY JDATE" & vbCrLf
    SQL = SQL & "End" & vbCrLf

    SQL = SQL & "ELSE IF @TAXREPORT = 'EXPANDED'" & vbCrLf
    SQL = SQL & "BEGIN" & vbCrLf
    SQL = SQL & "SET @TAX_REPORTS = CURSOR FOR" & vbCrLf
    SQL = SQL & "SELECT HD.JDATE,HD.JTYPE+'-'+HD.VOUCHERNO JOURNALNO," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "HD.VendorCode" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "HD.VendorCode" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "RIGHT(ISNULL(DT.ENTITY,''),6)" & vbCrLf
    SQL = SQL & "END AS VENDORCODE," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS VENDORNAME," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS ADDRESS," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS2 FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS2 FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS2 FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS ADDRESS2," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(DT.ENTITY,6))" & vbCrLf
    SQL = SQL & "END AS TIN," & vbCrLf
    SQL = SQL & "DT.ATC," & vbCrLf
    SQL = SQL & "(SELECT NATURE FROM AMIS_ATC WHERE ATC=DT.ATC) AS NATURE," & vbCrLf
    SQL = SQL & "DT.TAXBASE," & vbCrLf
    SQL = SQL & "DT.RATE," & vbCrLf
    SQL = SQL & "DT.CREDIT" & vbCrLf
    SQL = SQL & "FROM AMIS_JOURNAL_HD HD" & vbCrLf
    SQL = SQL & "INNER JOIN AMIS_JOURNAL_DET DT" & vbCrLf
    SQL = SQL & "ON HD.VOUCHERNO=DT.VOUCHERNO AND HD.JTYPE=DT.JTYPE" & vbCrLf
    SQL = SQL & "WHERE HD.STATUS='P' AND DT.ACCT_CODE=@ACCT_CODE AND DT.CREDIT > 0" & vbCrLf
    SQL = SQL & "AND HD.JDATE BETWEEN @JDATE1 AND @JDATE2" & vbCrLf
    SQL = SQL & "ORDER BY HD.JDATE" & vbCrLf
    SQL = SQL & "OPEN @TAX_REPORTS" & vbCrLf
    SQL = SQL & "FETCH NEXT FROM @TAX_REPORTS INTO @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@ATC,@NATURE,@TAXBASE,@RATE,@AMOUNT" & vbCrLf
    SQL = SQL & "SET @RUNNINGBALANCE=@AMOUNT" & vbCrLf
    SQL = SQL & "WHILE @@FETCH_STATUS=0" & vbCrLf
    SQL = SQL & "BEGIN" & vbCrLf
    SQL = SQL & "INSERT INTO @TAXREPORT_TABLE(JDATE,JOURNALNO,VENDORCODE,VENDORNAME,ADDRESS,CITY,TIN,ATC,NATURE,TAXBASE,RATE,TAXWITHHELD,RUNNINGBALANCE)" & vbCrLf
    SQL = SQL & "SELECT @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@ATC,@NATURE,@TAXBASE,@RATE,@AMOUNT,@RUNNINGBALANCE" & vbCrLf
    SQL = SQL & "FETCH NEXT FROM @TAX_REPORTS INTO @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@ATC,@NATURE,@TAXBASE,@RATE,@AMOUNT" & vbCrLf
    SQL = SQL & "SET @RUNNINGBALANCE=@RUNNINGBALANCE+@AMOUNT" & vbCrLf
    SQL = SQL & "End" & vbCrLf
    SQL = SQL & "CLOSE @TAX_REPORTS" & vbCrLf
    SQL = SQL & "DEALLOCATE @TAX_REPORTS" & vbCrLf
    SQL = SQL & "SELECT * FROM @TAXREPORT_TABLE ORDER BY JDATE" & vbCrLf
    SQL = SQL & "End" & vbCrLf

    SQL = SQL & "ELSE IF @TAXREPORT = 'COMPENSATION'" & vbCrLf
    SQL = SQL & "BEGIN" & vbCrLf
    SQL = SQL & "SET @TAX_REPORTS = CURSOR FOR" & vbCrLf
    SQL = SQL & "SELECT HD.JDATE,HD.JTYPE+'-'+HD.VOUCHERNO JOURNALNO," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "HD.VendorCode" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "HD.VendorCode" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "RIGHT(ISNULL(DT.ENTITY,''),6)" & vbCrLf
    SQL = SQL & "END AS VENDORCODE," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS VENDORNAME," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS ADDRESS," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS2 FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS2 FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ADDRESS2 FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS ADDRESS2," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'APJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CDJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_VENDOR_TABLE WHERE CODE=HD.VENDORCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_VENDOR_TABLE WHERE CODE=RIGHT(DT.ENTITY,6))" & vbCrLf
    SQL = SQL & "END AS TIN," & vbCrLf
    SQL = SQL & "DT.ATC," & vbCrLf
    SQL = SQL & "(SELECT NATURE FROM AMIS_ATC WHERE ATC=DT.ATC) AS NATURE," & vbCrLf
    SQL = SQL & "DT.TAXBASE," & vbCrLf
    SQL = SQL & "DT.RATE," & vbCrLf
    SQL = SQL & "DT.CREDIT" & vbCrLf
    SQL = SQL & "FROM AMIS_JOURNAL_HD HD" & vbCrLf
    SQL = SQL & "INNER JOIN AMIS_JOURNAL_DET DT" & vbCrLf
    SQL = SQL & "ON HD.VOUCHERNO=DT.VOUCHERNO AND HD.JTYPE=DT.JTYPE" & vbCrLf
    SQL = SQL & "WHERE HD.STATUS='P' AND DT.ACCT_CODE=@ACCT_CODE AND DT.CREDIT > 0" & vbCrLf
    SQL = SQL & "AND HD.JDATE BETWEEN @JDATE1 AND @JDATE2" & vbCrLf
    SQL = SQL & "ORDER BY HD.JDATE" & vbCrLf
    SQL = SQL & "OPEN @TAX_REPORTS" & vbCrLf
    SQL = SQL & "FETCH NEXT FROM @TAX_REPORTS INTO @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@ATC,@NATURE,@TAXBASE,@RATE,@AMOUNT" & vbCrLf
    SQL = SQL & "SET @RUNNINGBALANCE=@AMOUNT" & vbCrLf
    SQL = SQL & "WHILE @@FETCH_STATUS=0" & vbCrLf
    SQL = SQL & "BEGIN" & vbCrLf
    SQL = SQL & "INSERT INTO @TAXREPORT_TABLE(JDATE,JOURNALNO,VENDORCODE,VENDORNAME,ADDRESS,CITY,TIN,ATC,NATURE,TAXBASE,RATE,TAXWITHHELD,RUNNINGBALANCE)" & vbCrLf
    SQL = SQL & "SELECT @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@ATC,@NATURE,@TAXBASE,@RATE,@AMOUNT,@RUNNINGBALANCE" & vbCrLf
    SQL = SQL & "FETCH NEXT FROM @TAX_REPORTS INTO @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@ADDRESS2,@TIN,@ATC,@NATURE,@TAXBASE,@RATE,@AMOUNT" & vbCrLf
    SQL = SQL & "SET @RUNNINGBALANCE=@RUNNINGBALANCE+@AMOUNT" & vbCrLf
    SQL = SQL & "End" & vbCrLf
    SQL = SQL & "CLOSE @TAX_REPORTS" & vbCrLf
    SQL = SQL & "DEALLOCATE @TAX_REPORTS" & vbCrLf
    SQL = SQL & "SELECT * FROM @TAXREPORT_TABLE ORDER BY JDATE" & vbCrLf
    SQL = SQL & "End" & vbCrLf

    SQL = SQL & "ELSE IF @TAXREPORT = 'CREDITABLE'" & vbCrLf
    SQL = SQL & "BEGIN" & vbCrLf
    SQL = SQL & "SET @TAX_REPORTS = CURSOR FOR" & vbCrLf
    SQL = SQL & "SELECT HD.JDATE,HD.JTYPE+'-'+HD.VOUCHERNO JOURNALNO," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'SJ' THEN" & vbCrLf
    SQL = SQL & "HD.CustomerCode" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CRJ' THEN" & vbCrLf
    SQL = SQL & "HD.CustomerCode" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE='GJ' THEN" & vbCrLf
    SQL = SQL & "RIGHT(DT.ENTITY,6)" & vbCrLf
    SQL = SQL & "END AS CUSTOMERCODE," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'SJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CRJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE='GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS CUSTOMERNAME," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE = 'SJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT CUSTOMERADD FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'CRJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT CUSTOMERADD FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT CUSTOMERADD FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS ADDRESS," & vbCrLf
    SQL = SQL & "CASE WHEN HD.JTYPE='SJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE='CRJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=HD.CUSTOMERCODE)" & vbCrLf
    SQL = SQL & "WHEN HD.JTYPE = 'GJ' THEN" & vbCrLf
    SQL = SQL & "(SELECT TIN FROM ALL_CUSTOMER_TABLE WHERE CUSCDE=RIGHT(ISNULL(DT.ENTITY,''),6))" & vbCrLf
    SQL = SQL & "END AS TIN," & vbCrLf
    SQL = SQL & "DT.ATC," & vbCrLf
    SQL = SQL & "(SELECT NATURE FROM AMIS_ATC WHERE ATC=DT.ATC) AS NATURE," & vbCrLf
    SQL = SQL & "ISNULL(DT.TAXBASE,0)," & vbCrLf
    SQL = SQL & "ISNULL(DT.RATE,0)," & vbCrLf
    SQL = SQL & "DT.DEBIT" & vbCrLf
    SQL = SQL & "FROM AMIS_JOURNAL_HD HD" & vbCrLf
    SQL = SQL & "INNER JOIN AMIS_JOURNAL_DET DT" & vbCrLf
    SQL = SQL & "ON HD.VOUCHERNO=DT.VOUCHERNO AND HD.JTYPE=DT.JTYPE" & vbCrLf
    SQL = SQL & "WHERE HD.STATUS='P' AND DT.ACCT_CODE=@ACCT_CODE AND DT.DEBIT > 0" & vbCrLf
    SQL = SQL & "AND HD.JDATE BETWEEN @JDATE1 AND @JDATE2 AND HD.JTYPE IN ('SJ','CRJ','GJ','OPB')" & vbCrLf
    SQL = SQL & "ORDER BY HD.JDATE" & vbCrLf
    SQL = SQL & "OPEN @TAX_REPORTS" & vbCrLf
    SQL = SQL & "FETCH NEXT FROM @TAX_REPORTS INTO @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@TIN,@ATC,@NATURE,@TAXBASE,@RATE,@AMOUNT" & vbCrLf
    SQL = SQL & "SET @RUNNINGBALANCE=@AMOUNT" & vbCrLf
    SQL = SQL & "WHILE @@FETCH_STATUS=0" & vbCrLf
    SQL = SQL & "BEGIN" & vbCrLf
    SQL = SQL & "INSERT INTO @TAXREPORT_TABLE(JDATE,JOURNALNO,VENDORCODE,VENDORNAME,ADDRESS,TIN,ATC,NATURE,TAXBASE,RATE,TAXWITHHELD,RUNNINGBALANCE)" & vbCrLf
    SQL = SQL & "SELECT @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@TIN,@ATC,@NATURE,@TAXBASE,@RATE,@AMOUNT,@RUNNINGBALANCE" & vbCrLf
    SQL = SQL & "FETCH NEXT FROM @TAX_REPORTS INTO @JDATE,@JOURNALNO,@VENDORCODE,@VENDORNAME,@ADDRESS,@TIN,@ATC,@NATURE,@TAXBASE,@RATE,@AMOUNT" & vbCrLf
    SQL = SQL & "SET @RUNNINGBALANCE=@RUNNINGBALANCE+@AMOUNT" & vbCrLf
    SQL = SQL & "End" & vbCrLf
    SQL = SQL & "CLOSE @TAX_REPORTS" & vbCrLf
    SQL = SQL & "DEALLOCATE @TAX_REPORTS" & vbCrLf
    SQL = SQL & "SELECT * FROM @TAXREPORT_TABLE ORDER BY JDATE" & vbCrLf
    SQL = SQL & "End" & vbCrLf
    SQL = SQL & "End"
    gconDMIS.Execute SQL
End Sub


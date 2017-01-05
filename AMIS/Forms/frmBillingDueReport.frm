VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmBillingDueReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BILLING DUE REPORT"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15000
   Icon            =   "frmBillingDueReport.frx":0000
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
         MouseIcon       =   "frmBillingDueReport.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmBillingDueReport.frx":11D4
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
         MouseIcon       =   "frmBillingDueReport.frx":2256
         MousePointer    =   99  'Custom
         Picture         =   "frmBillingDueReport.frx":23A8
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
         MouseIcon       =   "frmBillingDueReport.frx":2847
         MousePointer    =   99  'Custom
         Picture         =   "frmBillingDueReport.frx":2999
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "View"
         Top             =   120
         Width           =   735
      End
      Begin VB.ComboBox cboOption 
         Height          =   315
         ItemData        =   "frmBillingDueReport.frx":2CE0
         Left            =   90
         List            =   "frmBillingDueReport.frx":2CE2
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   450
         Visible         =   0   'False
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
            TabIndex        =   14
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
         TabIndex        =   6
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55377921
         CurrentDate     =   40136
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   375
         Left            =   10110
         TabIndex        =   7
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55377921
         CurrentDate     =   40136
      End
      Begin FlexCell.Grid Grid1 
         Height          =   6765
         Left            =   30
         TabIndex        =   8
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
         TabIndex        =   13
         Top             =   180
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.Label lblBillingType 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   4950
         TabIndex        =   12
         Top             =   180
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   285
         Left            =   7500
         TabIndex        =   11
         Top             =   510
         Width           =   795
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   285
         Left            =   9780
         TabIndex        =   10
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
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   945
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   1125
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
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
Attribute VB_Name = "frmBillingDueReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPORTS                                               As ADODB.Recordset
Dim i                                                       As Integer
Dim ACCT_CODE                                               As String
Dim DESCRIPTION                                             As String
Dim REPORTTYPE                                              As String
Dim CMD                                                     As ADODB.Command
Dim BILLING_TYPE                                            As String
Dim xlsWorkSheet                                            As Excel.Worksheet
Dim iAMOUNTTOPAY                                            As Double

Private Sub cmdExport_Click()
    PopupMenu mnuExport
End Sub

Sub Report_Type(REPORT As String)
    BILLING_TYPE = REPORT
End Sub

Private Sub cmdInquire_Click()
    If dtTo.Value < dtFrom.Value Then
        MsgBox "Please check date selected.", vbInformation, "Date Range"
        Exit Sub
    End If
    Dim rsACCT_CODE                                         As ADODB.Recordset
    Set rsACCT_CODE = New ADODB.Recordset
    rsACCT_CODE.Open "SELECT ACCTCODE,TRANTYPE1,DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & cboOption.Text & "'", gconDMIS, adOpenForwardOnly
    If Not rsACCT_CODE.EOF And Not rsACCT_CODE.BOF Then
        ACCT_CODE = Null2String(rsACCT_CODE!AcctCode)
    End If
    Set rsACCT_CODE = Nothing

    If BILLING_TYPE = "AP DUE" Then
        REPORTTYPE = "AP"
    ElseIf BILLING_TYPE = "AR DUE" Then
        REPORTTYPE = "AR"
    End If
    BILLING_DUE_REPORT
End Sub

Private Sub cmdPrint_Click()
    CenterMe frmMain, Me, 1
    Dim rptApp                                              As CRAXDRT.Application
    Dim rptRep                                              As REPORT
    Dim crSections                                          As CRAXDRT.Sections
    Dim crSection                                           As CRAXDRT.Section
    Dim crRepObjs                                           As CRAXDRT.ReportObjects
    Dim crSubRepObj                                         As CRAXDRT.SubreportObject
    Dim crSubReport                                         As CRAXDRT.REPORT
    Dim j As Integer, k                                     As Integer
    Dim ellaine                                             As Integer

   Me.BorderStyle = vbSizable
    Me.WindowState = vbMaximized
    CRViewer1.Height = Me.Height - 700
    CRViewer1.Width = Me.Width
    CRViewer1.ZOrder 0

    Set rptApp = New CRAXDRT.Application
    If BILLING_TYPE = "AP DUE" Then
        Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "\BILLING2.rpt", 1)
    Else
        Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "\COLLECTION.rpt", 1)
    End If
    rptRep.DiscardSavedData
    
    'rptRep.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    'rptRep.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
'ELDAN 8-9-14
    rptRep.ParameterFields.GetItemByName("CompanyAddress").AddCurrentValue COMPANY_ADDRESS
    rptRep.ParameterFields.GetItemByName("CompanyName").AddCurrentValue COMPANY_NAME
    If BILLING_TYPE = "AP DUE" Then
        rptRep.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "BILLING DUE REPORT"
        Me.Caption = "BILLING DUE REPORT"
    ElseIf BILLING_TYPE = "AR DUE" Then
        rptRep.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "COLLECTION FORECAST REPORT"
        Me.Caption = "COLLECTION FORECAST REPORT"
    End If
    Call rptRep.ParameterFields(4).AddCurrentValue(CDate(dtFrom.Value))
    Call rptRep.ParameterFields(5).AddCurrentValue(CDate(dtTo.Value))
    If BILLING_TYPE = "AR DUE" Then
         Call rptRep.ParameterFields(6).AddCurrentValue("AR")
    End If
    Set crSections = rptRep.Sections
    For ellaine = 1 To crSections.Count
        Set crSection = crSections.Item(ellaine)
        Set crRepObjs = crSection.ReportObjects
        For j = 1 To crRepObjs.Count
            If crRepObjs.Item(j).Kind = crSubreportObject Then
'                Set crSubReport = rptRep.OpenSubreport(crRepObjs.Item(j).SubreportName)
'                Call crSubReport.ParameterFields(1).ClearCurrentValueAndRange
'                Call crSubReport.ParameterFields(1).AddCurrentValue(CDate(frmAMISFinancialStatements.dtpFrom.Value))
'                Call crSubReport.ParameterFields(2).ClearCurrentValueAndRange
'                Call crSubReport.ParameterFields(2).AddCurrentValue(CDate(frmAMISFinancialStatements.dtpTo.Value))
            End If
        Next
    Next
    With CRViewer1
        .ReportSource = rptRep
        .DisplayGroupTree = False
        .DisplayTabs = False
        .DisplayToolbar = True
        .ViewReport
    End With
    Set rptApp = Nothing
    Set rptRep = Nothing
 
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    initGrid
    Dim rsMAX_MIN_DATE                                      As ADODB.Recordset
    Set rsMAX_MIN_DATE = New ADODB.Recordset
    If BILLING_TYPE = "AP DUE" Then
        rsMAX_MIN_DATE.Open "SELECT * FROM (SELECT MAX(JDATE) AS MAX_JDATE, MIN(JDATE) AS MIN_JDATE FROM AMIS_AP WHERE STATUS = 'P')T WHERE MAX_JDATE IS NOT NULL", gconDMIS, adOpenKeyset
    Else
        rsMAX_MIN_DATE.Open "SELECT * FROM (SELECT MAX(JDATE) AS MAX_JDATE, MIN(JDATE) AS MIN_JDATE FROM AMIS_AR WHERE STATUS = 'P')T WHERE MAX_JDATE IS NOT NULL", gconDMIS, adOpenKeyset
    End If
    If Not rsMAX_MIN_DATE.EOF And Not rsMAX_MIN_DATE.BOF Then
        dtFrom.Value = rsMAX_MIN_DATE!MIN_JDATE
        dtTo.Value = rsMAX_MIN_DATE!MAX_JDATE
    Else
        dtFrom.Value = LOGDATE
        dtTo.Value = LOGDATE
    End If
    Set rsMAX_MIN_DATE = Nothing
    Screen.MousePointer = 0
    Grid1.Rows = 1
End Sub

Sub initGrid()
    With Grid1
        .Cols = 9
        .Rows = 1
        '        .FixedCols = 4
        .Cell(0, 0).Text = "L/N"

        If BILLING_TYPE = "AP DUE" Then
            .Cell(0, 1).Text = "VENDOR NAME"
        Else
            .Cell(0, 1).Text = "CUSTOMER NAME"
        End If
        .Column(1).Width = 200
        .Column(1).FormatString = "mm/dd/yyyy"

        .Cell(0, 2).Text = "VOUCHER NO"
        .Column(2).Alignment = cellCenterCenter
        .Column(2).Width = 75

        .Cell(0, 3).Text = "INVOICE NO"
        .Column(3).Alignment = cellCenterCenter
        .Column(3).Width = 75

        .Cell(0, 4).Text = "INVOICE DATE"
        .Column(4).Width = 80

        .Cell(0, 5).Text = "DUEDATE"
        .Column(5).Width = 80

        .Cell(0, 6).Text = "INVOICE AMOUNT"
        .Column(6).Alignment = cellRightCenter
        .Column(6).Width = 110

        .Cell(0, 7).Text = "ACCOUNT CODE"
        .Column(7).Width = 150
        .Column(7).Alignment = cellCenterCenter

        .Cell(0, 8).Text = "DESCRIPTION"
        .Column(8).Width = 200
    End With
End Sub

Sub BILLING_DUE_REPORT()
    Set CMD = New ADODB.Command
    CMD.ActiveConnection = gconDMIS
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "USP_BILLING_DATERANGE"
    CMD.CommandTimeout = 500

    With CMD.Parameters
        '.Append CMD.CreateParameter("@ACCT_CODE", adVarChar, adParamInput, 12, ACCT_CODE)
        .Append CMD.CreateParameter("@JDATE1", adDate, adParamInput, 8, dtFrom)
        .Append CMD.CreateParameter("@JDATE2", adDate, adParamInput, 8, dtTo)
        .Append CMD.CreateParameter("@REPORTTYPE", adVarChar, adParamInput, 8, REPORTTYPE)
    End With
    Set rsREPORTS = CMD.Execute
    FILLREPORTS
End Sub

Sub FILLREPORTS()
    Grid1.Rows = 1
    Grid1.AutoRedraw = False
    Picture1.Enabled = False
    Picture2.Visible = True
    iAMOUNTTOPAY = 0
    If Not rsREPORTS.EOF And Not rsREPORTS.BOF Then
        While Not rsREPORTS.EOF
            If REPORTTYPE = "AP" Then
                Grid1.AddItem _
                        rsREPORTS!VENDOR_NAME & Chr(9) & rsREPORTS!VOUCHERNO & Chr(9) & _
                                              rsREPORTS!INVOICENO & Chr(9) & rsREPORTS!invoicedate & Chr(9) & _
                                              Null2String(rsREPORTS!DUEDATE) & Chr(9) & ToDoubleNumber(rsREPORTS!AMOUNT2PAY) & Chr(9) & rsREPORTS!ACCT_CODE & Chr(9) & _
                                              rsREPORTS!DESCRIPTION
            Else
                Grid1.AddItem _
                        rsREPORTS!CUSTOMERNAME & Chr(9) & rsREPORTS!SJVoucherno & Chr(9) & _
                                               rsREPORTS!INVOICENO & Chr(9) & rsREPORTS!invoicedate & Chr(9) & _
                                               Null2String(rsREPORTS!DUEDATE) & Chr(9) & ToDoubleNumber(rsREPORTS!AR_TOPAY) & Chr(9) & rsREPORTS!ACCT_CODE & Chr(9) & _
                                               rsREPORTS!DESCRIPTION
            End If
            If REPORTTYPE = "AP" Then
                iAMOUNTTOPAY = iAMOUNTTOPAY + NumericVal(rsREPORTS!AMOUNT2PAY)
            Else
                iAMOUNTTOPAY = iAMOUNTTOPAY + NumericVal(rsREPORTS!AR_TOPAY)
            End If
            rsREPORTS.MoveNext
            Loading
        Wend
            Grid1.AddItem _
                        "" & Chr(9) & "" & Chr(9) & _
                        "" & Chr(9) & "" & Chr(9) & _
                        "TOTAL: " & Chr(9) & ToDoubleNumber(iAMOUNTTOPAY)
    End If
    Grid1.AutoRedraw = True
    Grid1.Refresh
    Picture1.Enabled = True
    Picture2.Visible = False
    cmdPrint.Enabled = True
    cmdExport.Enabled = True
    Set rsREPORTS = Nothing
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

Private Sub mnuExcel_Click()
    Grid1.ExportToExcel ("")
End Sub

Private Sub mnuHTML_Click()
    Grid1.ExportToHTML ("")
End Sub

Private Sub mnuPDF_Click()
    Grid1.ExportToPDF ("")
End Sub



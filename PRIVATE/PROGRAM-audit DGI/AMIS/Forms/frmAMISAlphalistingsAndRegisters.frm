VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAMISAlphalistingsAndRegisters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alpha Listings And Registers"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   ForeColor       =   &H8000000F&
   Icon            =   "frmAMISAlphalistingsAndRegisters.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4590
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
      Height          =   705
      Left            =   2190
      MouseIcon       =   "frmAMISAlphalistingsAndRegisters.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "frmAMISAlphalistingsAndRegisters.frx":0E1C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Close Window"
      Top             =   4530
      Width           =   795
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
      Height          =   705
      Left            =   1410
      MouseIcon       =   "frmAMISAlphalistingsAndRegisters.frx":1267
      MousePointer    =   99  'Custom
      Picture         =   "frmAMISAlphalistingsAndRegisters.frx":13B9
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Print Report"
      Top             =   4530
      Width           =   795
   End
   Begin VB.OptionButton optOfficialReceiptRegister 
      Caption         =   "Official Receipt Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   405
      TabIndex        =   14
      Top             =   3780
      Width           =   3870
   End
   Begin VB.OptionButton optPurchaseOrderRegister 
      Caption         =   "Purchase Order Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   405
      TabIndex        =   13
      Top             =   4080
      Width           =   3930
   End
   Begin VB.OptionButton optMaterialsIssuanceSlipRegister 
      Caption         =   "Materials Issuance Slip Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   405
      TabIndex        =   12
      Top             =   3480
      Width           =   3405
   End
   Begin VB.OptionButton optPartsIssuanceSlipRegister 
      Caption         =   "Parts Issuance Slip Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   420
      TabIndex        =   11
      Top             =   2895
      Width           =   3870
   End
   Begin VB.OptionButton optAccessoriesIssuanceSlipRegister 
      Caption         =   "Accessories Issuance Slip Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   420
      TabIndex        =   10
      Top             =   3195
      Width           =   3930
   End
   Begin VB.OptionButton optDeliveryReceiptPartsAndAccessoriesRegister 
      Caption         =   "Delivery Receipt - Parts And Accessories Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   420
      Left            =   405
      TabIndex        =   9
      Top             =   2115
      Width           =   3930
   End
   Begin VB.OptionButton optServiceInvoiceRegister 
      Caption         =   "Service Invoice Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   405
      TabIndex        =   8
      Top             =   2595
      Width           =   3405
   End
   Begin VB.OptionButton optDeliveryReceiptCBURegister 
      Caption         =   "Delivery Receipt - CBU Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   405
      TabIndex        =   7
      Top             =   1305
      Width           =   3405
   End
   Begin VB.OptionButton optSalesInvoicePartsAndAccessoriesRegister 
      Caption         =   "Sales Invoice - Parts And Accessories Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   510
      Left            =   390
      TabIndex        =   6
      Top             =   1560
      Width           =   3855
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
      Left            =   2955
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select year from the list"
      Top             =   225
      Width           =   1470
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
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select month from the list"
      Top             =   225
      Width           =   1470
   End
   Begin VB.OptionButton optVehicleSalesOrderRegister 
      Caption         =   "Vehicle Sales Order (VSO) Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   405
      TabIndex        =   1
      Top             =   705
      Value           =   -1  'True
      Width           =   3870
   End
   Begin VB.OptionButton optSalesInvoiceCBURegister 
      Caption         =   "Sales Invoice - CBU Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   405
      TabIndex        =   0
      Top             =   1005
      Width           =   3405
   End
   Begin Crystal.CrystalReport rptAlphaListingsAndRegisters 
      Left            =   300
      Top             =   4650
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Alpha Listings And Registers"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   2355
      TabIndex        =   5
      Top             =   300
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   270
      Width           =   675
   End
End
Attribute VB_Name = "frmAMISAlphalistingsAndRegisters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode
    Dim rsAlphaListing                                      As ADODB.Recordset
    Set rsAlphaListing = New ADODB.Recordset
    Set rsAlphaListing = gconDMIS.Execute("Select * from SMIS_SalesOrder")
    If Not rsAlphaListing.EOF And Not rsAlphaListing.BOF Then

        'rptAMISDueReport.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        'rptAMISDueReport.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAlphaListingsAndRegisters.Formulas(3) = "MonthReport = '" & cboMonth.Text & "'"
        rptAlphaListingsAndRegisters.Formulas(4) = "YearReport = '" & cboYear.Text & "'"

        'Vehicle Sales Order
        If optVehicleSalesOrderRegister.Value = True Then
            Dim rsVehicleOrder                              As ADODB.Recordset
            Set rsVehicleOrder = New ADODB.Recordset
            Set rsVehicleOrder = gconDMIS.Execute("Select * from SMIS_SalesOrder where Month(Deyt) = '" & What_month(cboMonth.Text) & "' AND Year(Deyt) = '" & cboYear.Text & "'")
            If Not rsVehicleOrder.EOF And Not rsVehicleOrder.BOF Then
                PrintSQLReport rptAlphaListingsAndRegisters, AMIS_REPORT_PATH & "AlphaListingsAndRegisters\VehicleSalesOrderRegister.rpt", "Month({SMIS_SalesOrder.Deyt}) = " & What_month(cboMonth.Text) & " AND Year({SMIS_SalesOrder.Deyt}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                LogAudit "V", "VEHICLE SALES ORDER REGISTER", cboMonth & "-" & cboYear
            Else
                ShowNoRecord
                Exit Sub
            End If
            'Sales Invoice - CBU Register
        ElseIf optSalesInvoiceCBURegister.Value = True Then
            Dim rsSalesInvoiceRegister                      As ADODB.Recordset
            Set rsSalesInvoiceRegister = New ADODB.Recordset
            Set rsSalesInvoiceRegister = gconDMIS.Execute("Select * from SMIS_SalesOrder where Month(InvoicedDate) = '" & What_month(cboMonth.Text) & "' AND Year(InvoicedDate) = '" & cboYear.Text & "'")
            If Not rsSalesInvoiceRegister.EOF And Not rsSalesInvoiceRegister.BOF Then
                PrintSQLReport rptAlphaListingsAndRegisters, AMIS_REPORT_PATH & "AlphaListingsAndRegisters\SalesInvoiceCBURegister.rpt", "Month({SMIS_SalesOrder.InvoicedDate}) = " & What_month(cboMonth.Text) & " AND Year({SMIS_SalesOrder.InvoicedDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                LogAudit "V", "SALES INVOICE - CBU REGISTER", cboMonth & "-" & cboYear
            Else
                ShowNoRecord
                Exit Sub
            End If
            'Delivery Receipt - CBU Register
        ElseIf optDeliveryReceiptCBURegister.Value = True Then
            Dim rsDeliveryReceiptRegister                   As ADODB.Recordset
            Set rsDeliveryReceiptRegister = New ADODB.Recordset
            Set rsDeliveryReceiptRegister = gconDMIS.Execute("Select * from SMIS_MrrInv where Month(DateReceived) = '" & What_month(cboMonth.Text) & "' AND Year(DateReceived) = '" & cboYear.Text & "'")
            If Not rsDeliveryReceiptRegister.EOF And Not rsDeliveryReceiptRegister.BOF Then
                PrintSQLReport rptAlphaListingsAndRegisters, AMIS_REPORT_PATH & "AlphaListingsAndRegisters\DeliveryReceiptCBURegister.rpt", "Month({SMIS_MrrInv.DateReceived}) = " & What_month(cboMonth.Text) & " AND Year({SMIS_MrrInv.DateReceived}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                LogAudit "V", "DELIVERY RECEIPT - CBU REGISTER", cboMonth & "-" & cboYear
            Else
                ShowNoRecord
                Exit Sub
            End If
            'Service Invoice Register
        ElseIf optServiceInvoiceRegister.Value = True Then
            Dim rsServiceInvoiceRegister                    As ADODB.Recordset
            Set rsServiceInvoiceRegister = New ADODB.Recordset
            Set rsServiceInvoiceRegister = gconDMIS.Execute("Select * from CSMS_vw_Repor where Month(DTE_RECD) = '" & What_month(cboMonth.Text) & "' AND Year(DTE_RECD) = '" & cboYear.Text & "'")
            If Not rsServiceInvoiceRegister.EOF And Not rsServiceInvoiceRegister.BOF Then
                PrintSQLReport rptAlphaListingsAndRegisters, AMIS_REPORT_PATH & "AlphaListingsAndRegisters\ServiceInvoiceRegister.rpt", "Month({CSMS_vw_Repor.DTE_RECD}) = " & What_month(cboMonth.Text) & " AND Year({CSMS_vw_Repor.DTE_RECD}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                LogAudit "V", "SERVICE INVOICE REGISTER", cboMonth & "-" & cboYear
            Else
                ShowNoRecord
                Exit Sub
            End If
            'Official Receipt Register
        ElseIf optOfficialReceiptRegister.Value = True Then
            Dim rsOfficialReceiptRegister                   As ADODB.Recordset
            Set rsOfficialReceiptRegister = New ADODB.Recordset
            Set rsOfficialReceiptRegister = gconDMIS.Execute("Select * from CMIS_Off_Hd where Month(OR_DATE) = '" & What_month(cboMonth.Text) & "' AND Year(OR_DATE) = '" & cboYear.Text & "'")
            If Not rsOfficialReceiptRegister.EOF And Not rsOfficialReceiptRegister.BOF Then
                PrintSQLReport rptAlphaListingsAndRegisters, AMIS_REPORT_PATH & "AlphaListingsAndRegisters\OfficialReceiptRegister.rpt", "Month({CMIS_Off_Hd.OR_DATE}) = " & What_month(cboMonth.Text) & " AND Year({CMIS_Off_Hd.OR_DATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                LogAudit "V", "OFFICIAL RECEIPT REGISTER", cboMonth & "-" & cboYear
            Else
                ShowNoRecord
                Exit Sub
            End If
            'Purchase Order Register
        ElseIf optPurchaseOrderRegister.Value = True Then
            Dim rsPurchaseOrderRegister                     As ADODB.Recordset
            Set rsPurchaseOrderRegister = New ADODB.Recordset
            Set rsPurchaseOrderRegister = gconDMIS.Execute("Select * from SMIS_PO where Month(DateOrdered) = '" & What_month(cboMonth.Text) & "' AND Year(DateOrdered) = '" & cboYear.Text & "'")
            If Not rsPurchaseOrderRegister.EOF And Not rsPurchaseOrderRegister.BOF Then
                PrintSQLReport rptAlphaListingsAndRegisters, AMIS_REPORT_PATH & "AlphaListingsAndRegisters\PurchaseOrderRegister.rpt", "Month({SMIS_PO.DateOrdered}) = " & What_month(cboMonth.Text) & " AND Year({SMIS_PO.DateOrdered}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                LogAudit "V", "PURCHASE ORDER REGISTER", cboMonth & "-" & cboYear
            Else
                ShowNoRecord
                Exit Sub
            End If
            'Sales Invoice - Parts And Accessories Register
        ElseIf optSalesInvoicePartsAndAccessoriesRegister.Value = True Then
            Dim rsSalesInvoicePARegister                    As ADODB.Recordset
            Set rsSalesInvoicePARegister = New ADODB.Recordset
            Set rsSalesInvoicePARegister = gconDMIS.Execute("Select * from PMIS_Ord_Hd where Month(TRANDATE) = '" & What_month(cboMonth.Text) & "' AND Year(TRANDATE) = '" & cboYear.Text & "'")
            If Not rsSalesInvoicePARegister.EOF And Not rsSalesInvoicePARegister.BOF Then
                PrintSQLReport rptAlphaListingsAndRegisters, AMIS_REPORT_PATH & "AlphaListingsAndRegisters\SalesInvoicePartsAccessoriesRegister.rpt", "Month({PMIS_Ord_Hd.TRANDATE}) = " & What_month(cboMonth.Text) & " AND Year({PMIS_Ord_Hd.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                LogAudit "V", "SALES INVOICE - PARTS AND ACCESSORIES REGISTER", cboMonth & "-" & cboYear
            Else
                ShowNoRecord
                Exit Sub
            End If
            'Delivery Receipt - Parts And Accessories Register
        ElseIf optDeliveryReceiptPartsAndAccessoriesRegister.Value = True Then
            Dim rsDelReceiptPARegister                      As ADODB.Recordset
            Set rsDelReceiptPARegister = New ADODB.Recordset
            Set rsDelReceiptPARegister = gconDMIS.Execute("Select * from PMIS_RR_Hd where Month(RRDATE) = '" & What_month(cboMonth.Text) & "' AND Year(RRDATE) = '" & cboYear.Text & "'")
            If Not rsDelReceiptPARegister.EOF And Not rsDelReceiptPARegister.BOF Then
                PrintSQLReport rptAlphaListingsAndRegisters, AMIS_REPORT_PATH & "AlphaListingsAndRegisters\DeliveryReceiptPartsAccessoriesRegister.rpt", "Month({PMIS_RR_Hd.RRDATE}) = " & What_month(cboMonth.Text) & " AND Year({PMIS_RR_Hd.RRDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                LogAudit "V", "DELIVERY RECEIPT - PARTS AND ACCESSORIES REGISTER", cboMonth & "-" & cboYear
            Else
                ShowNoRecord
                Exit Sub
            End If
            'Parts Issuance Slip Register
        ElseIf optPartsIssuanceSlipRegister.Value = True Then
            Dim rsPartsIssuanceRegister                     As ADODB.Recordset
            Set rsPartsIssuanceRegister = New ADODB.Recordset
            Set rsPartsIssuanceRegister = gconDMIS.Execute("Select * from PMIS_Ord_Hd where Month(TRANDATE) = '" & What_month(cboMonth.Text) & "' AND Year(TRANDATE) = '" & cboYear.Text & "'")
            If Not rsPartsIssuanceRegister.EOF And Not rsPartsIssuanceRegister.BOF Then
                PrintSQLReport rptAlphaListingsAndRegisters, AMIS_REPORT_PATH & "AlphaListingsAndRegisters\PartsIssuanceSlipRegister.rpt", "Month({PMIS_Ord_Hd.TRANDATE}) = " & What_month(cboMonth.Text) & " AND Year({PMIS_Ord_Hd.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                LogAudit "V", "PARTS ISSUANCE SLIP REGISTER", cboMonth & "-" & cboYear
            Else
                ShowNoRecord
                Exit Sub
            End If
            'Accessories Issuance Slip Register
        ElseIf optAccessoriesIssuanceSlipRegister.Value = True Then
            Dim rsAccessoriesIssuanceRegister               As ADODB.Recordset
            Set rsAccessoriesIssuanceRegister = New ADODB.Recordset
            Set rsAccessoriesIssuanceRegister = gconDMIS.Execute("Select * from PMIS_Ord_Hd where Month(TRANDATE) = '" & What_month(cboMonth.Text) & "' AND Year(TRANDATE) = '" & cboYear.Text & "'")
            If Not rsAccessoriesIssuanceRegister.EOF And Not rsAccessoriesIssuanceRegister.BOF Then
                PrintSQLReport rptAlphaListingsAndRegisters, AMIS_REPORT_PATH & "AlphaListingsAndRegisters\AccessoriesIssuanceSlipRegister.rpt", "Month({PMIS_Ord_Hd.TRANDATE}) = " & What_month(cboMonth.Text) & " AND Year({PMIS_Ord_Hd.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                LogAudit "V", "ACCESSORIES ISSUANCE SLIP REGISTER", cboMonth & "-" & cboYear
            Else
                ShowNoRecord
                Exit Sub
            End If
            'Materials Issuance Slip Register
        ElseIf optMaterialsIssuanceSlipRegister.Value = True Then
            Dim rsMaterialsIssuanceRegister                 As ADODB.Recordset
            Set rsMaterialsIssuanceRegister = New ADODB.Recordset
            Set rsMaterialsIssuanceRegister = gconDMIS.Execute("Select * from PMIS_Ord_Hd where Month(TRANDATE) = '" & What_month(cboMonth.Text) & "' AND Year(TRANDATE) = '" & cboYear.Text & "'")
            If Not rsMaterialsIssuanceRegister.EOF And Not rsMaterialsIssuanceRegister.BOF Then
                PrintSQLReport rptAlphaListingsAndRegisters, AMIS_REPORT_PATH & "AlphaListingsAndRegisters\MaterialsIssuanceSlipRegister.rpt", "Month({PMIS_Ord_Hd.TRANDATE}) = " & What_month(cboMonth.Text) & " AND Year({PMIS_Ord_Hd.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                LogAudit "V", "MATERIALS ISSUANCE SLIP REGISTER", cboMonth & "-" & cboYear
            Else
                ShowNoRecord
                Exit Sub
            End If
        End If
    Else
        ShowNoRecord
    End If
    Exit Sub
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    fillcbomonth cboMonth
    FillcboYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
End Sub


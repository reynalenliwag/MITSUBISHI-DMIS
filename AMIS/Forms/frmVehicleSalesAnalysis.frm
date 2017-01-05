VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmVehicleSalesAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VEHICLE SALES ANALYSIS"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15000
   Icon            =   "frmVehicleSalesAnalysis.frx":0000
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
         MouseIcon       =   "frmVehicleSalesAnalysis.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesAnalysis.frx":11D4
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
         MouseIcon       =   "frmVehicleSalesAnalysis.frx":2256
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesAnalysis.frx":23A8
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
         MouseIcon       =   "frmVehicleSalesAnalysis.frx":2847
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesAnalysis.frx":2999
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "View"
         Top             =   120
         Width           =   735
      End
      Begin VB.ComboBox cboOption 
         Height          =   315
         ItemData        =   "frmVehicleSalesAnalysis.frx":2CE0
         Left            =   90
         List            =   "frmVehicleSalesAnalysis.frx":2CE2
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
         Format          =   55508993
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
         Format          =   55508993
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
Attribute VB_Name = "frmVehicleSalesAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPORTS                                               As ADODB.Recordset
Dim i                                                       As Integer
Dim ACCT_CODE                                               As String
Dim DESCRIPTION                                             As String
Dim CMD                                                     As ADODB.Command
Dim BILLING_TYPE                                            As String
Dim xlsWorkSheet                                            As Excel.Worksheet
Dim QTY                                                     As Long
Dim UNITCOST                                                As Double
Dim TOTALACCESS                                             As Double
Dim TOTALCOST                                               As Double
Dim TOTALDISC                                               As Double
Dim SRP                                                     As Double
Dim SRPNETDISC                                              As Double
Dim SRPNETVAT                                               As Double
Dim OUTPUT                                                  As Double
Dim GROSSPROFIT                                             As Double
Dim GP                                                      As Double
Dim ADDTLDISC                                               As Double
Dim NETSALES                                                As Double
Dim FREEINSURANCE As Double
Dim FREELTO As Double
Dim FREECHATTEL As Double
Dim FREEBIES As Double
Dim TOTALFREE As Double
Dim COMMISSION As Double
Dim NETPROFIT As Double

Private Sub cmdExport_Click()
    PopupMenu mnuExport
End Sub

Sub Report_Type(REPORT As String)
    BILLING_TYPE = REPORT
End Sub

Private Sub cmdInquire_Click()
    initGrid
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
    SALES_ANALYSIS
End Sub

Private Sub cmdPrint_Click()
    Dim xlsApplication                                      As Excel.Application
    Dim xlsWorkbook                                         As Excel.Workbook
    Dim ellaine                                             As Integer
    QTY = 0
    UNITCOST = 0
    TOTALACCESS = 0
    TOTALCOST = 0
    TOTALDISC = 0
    ADDTLDISC = 0
    NETSALES = 0
    SRP = 0
    SRPNETDISC = 0
    SRPNETVAT = 0
    OUTPUT = 0
    GROSSPROFIT = 0
    GP = 0
    FREEINSURANCE = 0
    FREELTO = 0
    FREECHATTEL = 0
    FREEBIES = 0
    TOTALFREE = 0
    COMMISSION = 0
    NETPROFIT = 0
    Set xlsApplication = New Excel.Application
    If Len(Dir(AMIS_REPORT_PATH & "VehicleSalesAnalysis.xlt")) = 0 Then
        MsgBox "Report file cannot be found.", vbInformation
        Exit Sub
    End If
    Set xlsWorkbook = xlsApplication.Workbooks.Open(AMIS_REPORT_PATH & "VehicleSalesAnalysis.xlt")
    Set xlsWorkSheet = xlsWorkbook.Worksheets(1)
    xlsWorkSheet.Cells(1, "B") = COMPANY_NAME
    xlsWorkSheet.Cells(2, "B") = "VEHICLE SALES ANALYSIS"
    xlsWorkSheet.Cells(3, "B") = "FOR THE MONTH OF " & UCase(Format(LOGDATE, "MMMM"))
    Set rsREPORTS = CMD.Execute
    If Not rsREPORTS.EOF And Not rsREPORTS.BOF Then
        'xlsWorkSheet.Cells(7, 1).CopyFromRecordset rsREPORTS
        Do While Not rsREPORTS.EOF
            xlsWorkSheet.Cells(9 + ellaine, "A") = Null2String(rsREPORTS!SONO)
            xlsWorkSheet.Cells(9 + ellaine, "B") = Null2String(rsREPORTS!CustomerCode)
            xlsWorkSheet.Cells(9 + ellaine, "C") = Null2String(rsREPORTS!Customer)
            xlsWorkSheet.Cells(9 + ellaine, "D") = Null2String(rsREPORTS!Make)
            xlsWorkSheet.Cells(9 + ellaine, "E") = Null2String(rsREPORTS!VINO)
            xlsWorkSheet.Cells(9 + ellaine, "F") = Null2String(rsREPORTS!prodno)
            xlsWorkSheet.Cells(9 + ellaine, "G") = Null2String(rsREPORTS!invoicedate)
            xlsWorkSheet.Cells(9 + ellaine, "H") = Null2String(rsREPORTS!DATERELEASED)
            xlsWorkSheet.Cells(9 + ellaine, "I") = Null2String(rsREPORTS!BANKTERM)
            xlsWorkSheet.Cells(9 + ellaine, "J") = Null2String(rsREPORTS!Bank)
            xlsWorkSheet.Cells(9 + ellaine, "K") = NumericVal(rsREPORTS!QTY)
            
            xlsWorkSheet.Cells(9 + ellaine, "L") = ToDoubleNumber(rsREPORTS!SRP)
            xlsWorkSheet.Cells(9 + ellaine, "M") = ToDoubleNumber(rsREPORTS!DISCOUNT)
            xlsWorkSheet.Cells(9 + ellaine, "N") = ToDoubleNumber(rsREPORTS!SRPNETDISC)
            xlsWorkSheet.Cells(9 + ellaine, "O") = ToDoubleNumber(rsREPORTS!OUTPUT)
            xlsWorkSheet.Cells(9 + ellaine, "P") = ToDoubleNumber(rsREPORTS!SRPNETVAT)
            xlsWorkSheet.Cells(9 + ellaine, "Q") = Null2String(rsREPORTS!CMNO)
            xlsWorkSheet.Cells(9 + ellaine, "R") = ToDoubleNumber(rsREPORTS!ADDTLDISC)
            xlsWorkSheet.Cells(9 + ellaine, "S") = ToDoubleNumber(rsREPORTS!NETSALES)
            xlsWorkSheet.Cells(9 + ellaine, "T") = ToDoubleNumber(rsREPORTS!UNITCOST)
            xlsWorkSheet.Cells(9 + ellaine, "U") = ToDoubleNumber(rsREPORTS!TOTALACCESS)
            xlsWorkSheet.Cells(9 + ellaine, "V") = ToDoubleNumber(rsREPORTS!TOTALCOST)
            xlsWorkSheet.Cells(9 + ellaine, "W") = ToDoubleNumber(rsREPORTS!GROSSPROFIT)
            xlsWorkSheet.Cells(9 + ellaine, "X") = rsREPORTS!GP
            
            xlsWorkSheet.Cells(9 + ellaine, "Y") = ToDoubleNumber(rsREPORTS!FREEINSURANCE)
            xlsWorkSheet.Cells(9 + ellaine, "Z") = ToDoubleNumber(rsREPORTS!FREELTO)
            xlsWorkSheet.Cells(9 + ellaine, "AA") = ToDoubleNumber(rsREPORTS!FREECHATTEL)
            xlsWorkSheet.Cells(9 + ellaine, "AB") = ToDoubleNumber(rsREPORTS!FREEBIES)
            xlsWorkSheet.Cells(9 + ellaine, "AC") = ToDoubleNumber(rsREPORTS!TOTALFREE)
            xlsWorkSheet.Cells(9 + ellaine, "AD") = ToDoubleNumber(rsREPORTS!COMMISSION)
            xlsWorkSheet.Cells(9 + ellaine, "AE") = ToDoubleNumber(rsREPORTS!NETPROFIT)

            QTY = QTY + NumericVal(rsREPORTS!QTY)
            UNITCOST = UNITCOST + rsREPORTS!UNITCOST
            TOTALACCESS = TOTALACCESS + rsREPORTS!TOTALACCESS
            TOTALCOST = TOTALCOST + rsREPORTS!TOTALCOST
            TOTALDISC = TOTALDISC + rsREPORTS!DISCOUNT
            ADDTLDISC = ADDTLDISC + rsREPORTS!ADDTLDISC
            NETSALES = NETSALES + rsREPORTS!NETSALES
            SRP = SRP + rsREPORTS!SRP
            SRPNETDISC = SRPNETDISC + rsREPORTS!SRPNETDISC
            SRPNETVAT = SRPNETVAT + rsREPORTS!SRPNETVAT
            OUTPUT = OUTPUT + rsREPORTS!OUTPUT
            GROSSPROFIT = GROSSPROFIT + rsREPORTS!GROSSPROFIT
            GP = GP + ((rsREPORTS!NETSALES - rsREPORTS!TOTALCOST) / rsREPORTS!TOTALCOST) * 100
            FREEINSURANCE = FREEINSURANCE + rsREPORTS!FREEINSURANCE
            FREELTO = FREELTO + rsREPORTS!FREELTO
            FREECHATTEL = FREECHATTEL + rsREPORTS!FREECHATTEL
            FREEBIES = FREEBIES + rsREPORTS!FREEBIES
            TOTALFREE = TOTALFREE + rsREPORTS!TOTALFREE
            COMMISSION = COMMISSION + rsREPORTS!COMMISSION
            NETPROFIT = NETPROFIT + rsREPORTS!NETPROFIT
            rsREPORTS.MoveNext
            Loading
            ellaine = ellaine + 1
        Loop
    End If
    xlsWorkSheet.Cells(5, "A") = NumericVal(QTY)
    xlsWorkSheet.Cells(9 + ellaine, "L") = ToDoubleNumber(SRP)
    xlsWorkSheet.Cells(9 + ellaine, "M") = ToDoubleNumber(TOTALDISC)
    xlsWorkSheet.Cells(9 + ellaine, "N") = ToDoubleNumber(SRPNETDISC)
    xlsWorkSheet.Cells(9 + ellaine, "O") = ToDoubleNumber(OUTPUT)
    xlsWorkSheet.Cells(9 + ellaine, "P") = ToDoubleNumber(SRPNETVAT)
    '
    xlsWorkSheet.Cells(9 + ellaine, "R") = ToDoubleNumber(ADDTLDISC)
    xlsWorkSheet.Cells(9 + ellaine, "S") = ToDoubleNumber(NETSALES)
    xlsWorkSheet.Cells(9 + ellaine, "T") = ToDoubleNumber(UNITCOST)
    xlsWorkSheet.Cells(9 + ellaine, "U") = ToDoubleNumber(TOTALACCESS)
    xlsWorkSheet.Cells(9 + ellaine, "V") = ToDoubleNumber(TOTALCOST)
    xlsWorkSheet.Cells(9 + ellaine, "W") = ToDoubleNumber(GROSSPROFIT)
    xlsWorkSheet.Cells(9 + ellaine, "X") = N2String(ToDoubleNumber(GP)) & "%"
    
    xlsWorkSheet.Cells(9 + ellaine, "Y") = ToDoubleNumber(FREEINSURANCE)
    xlsWorkSheet.Cells(9 + ellaine, "Z") = ToDoubleNumber(FREELTO)
    xlsWorkSheet.Cells(9 + ellaine, "AA") = ToDoubleNumber(FREECHATTEL)
    xlsWorkSheet.Cells(9 + ellaine, "AB") = ToDoubleNumber(FREEBIES)
    xlsWorkSheet.Cells(9 + ellaine, "AC") = ToDoubleNumber(TOTALFREE)
    xlsWorkSheet.Cells(9 + ellaine, "AD") = ToDoubleNumber(COMMISSION)
    xlsWorkSheet.Cells(9 + ellaine, "AE") = ToDoubleNumber(NETPROFIT)
    xlsApplication.Visible = True
    Set xlsApplication = Nothing
    Set xlsWorkbook = Nothing
    Set xlsWorkSheet = Nothing
    Set rsREPORTS = Nothing
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    initGrid
    dtFrom.Value = firstDay(LOGDATE)
    dtTo.Value = lastDay(LOGDATE)
    Screen.MousePointer = 0
    Grid1.Rows = 1
End Sub

Sub initGrid()
    With Grid1
        .Cols = 32
        .Rows = 1
        '        .FixedCols = 4
        .Cell(0, 0).Text = "L/N"
        .RowHeight(0) = 40
        
        .Cell(0, 1).Text = "SO NO"
        .Column(1).Width = 75
        
        .Cell(0, 2).Text = "CUSTOMER CODE"
        .Column(2).Alignment = cellCenterCenter
        .Column(2).Width = 100
        
        .Cell(0, 3).Text = "CUSTOMER NAME"
        .Column(3).Width = 200
        '.Column(1).FormatString = "mm/dd/yyyy"

        .Cell(0, 4).Text = "MAKE"
        .Column(4).Alignment = cellLeftCenter
        .Column(4).Width = 250

        .Cell(0, 5).Text = "INVOICE NO"
        .Column(5).Alignment = cellCenterCenter
        .Column(5).Width = 75

        .Cell(0, 6).Text = "PRODUCT NO"
        .Column(6).Alignment = cellCenterCenter
        .Column(6).Width = 80

        .Cell(0, 7).Text = "INVOICE DATE"
        .Column(7).Alignment = cellCenterCenter
        .Column(7).Width = 80

        .Cell(0, 8).Text = "DATE RELEASE"
        .Column(8).Alignment = cellCenterCenter
        .Column(8).Width = 80

        .Cell(0, 9).Text = "BANK TERM"
        .Column(9).Width = 80
        .Column(9).Alignment = cellCenterCenter

        .Cell(0, 10).Text = "BANK"
        .Column(10).Width = 200

        .Cell(0, 11).Text = "QTY"
        .Column(11).Width = 80
        .Column(1).Alignment = cellCenterCenter

        .Cell(0, 12).Text = "SRP"
        .Column(12).Width = 80
        .Column(12).Alignment = cellRightCenter
        
        .Cell(0, 13).Text = "TOTAL DISC"
        .Column(13).Width = 80
        .Column(13).Alignment = cellRightCenter
        
        .Cell(0, 14).Text = "SRP NET OF DISC"
        .Column(14).Width = 110
        .Column(14).Alignment = cellRightCenter
        
        .Cell(0, 15).Text = "OUTPUT"
        .Column(15).Width = 80
        .Column(15).Alignment = cellRightCenter
        
        .Cell(0, 16).Text = "SRP NET OF VAT"
        .Column(16).Width = 110
        .Column(16).Alignment = cellRightCenter
        
        .Cell(0, 17).Text = "CM NO."
        .Column(17).Width = 80
        .Column(17).Alignment = cellRightCenter
        
        .Cell(0, 18).Text = "ADDITIONAL DISC"
        .Column(18).Width = 80
        .Column(18).Alignment = cellRightCenter
        
        .Cell(0, 19).Text = "NET SALES"
        .Column(19).Width = 80
        .Column(19).Alignment = cellRightCenter
        
        .Cell(0, 20).Text = "UNIT COST"
        .Column(20).Width = 80
        .Column(20).Alignment = cellRightCenter

        .Cell(0, 21).Text = "TOTAL ACCESS"
        .Column(21).Width = 80
        .Column(21).Alignment = cellRightCenter

        .Cell(0, 22).Text = "TOTAL COST"
        .Column(22).Width = 80
        .Column(22).Alignment = cellRightCenter

        .Cell(0, 23).Text = "GROSS PROFIT"
        .Column(23).Width = 80
        .Column(23).Alignment = cellRightCenter

        .Cell(0, 24).Text = "GP%"
        .Column(24).Width = 80
        .Column(24).Alignment = cellRightCenter
        
        .Cell(0, 25).Text = "FREE INSURANCE"
        .Column(25).Width = 100
        .Column(25).Alignment = cellRightCenter
        
        .Cell(0, 26).Text = "FREE LTO"
        .Column(26).Width = 80
        .Column(26).Alignment = cellRightCenter
        
        .Cell(0, 27).Text = "FREE CHATTEL"
        .Column(27).Width = 80
        .Column(27).Alignment = cellRightCenter
        
        .Cell(0, 28).Text = "FREEBIES"
        .Column(28).Width = 80
        .Column(28).Alignment = cellRightCenter
        
        .Cell(0, 29).Text = "TOTAL FREE ITEMS"
        .Column(29).Width = 100
        .Column(29).Alignment = cellRightCenter
        
        .Cell(0, 30).Text = "COMMISSION"
        .Column(30).Width = 100
        .Column(30).Alignment = cellRightCenter
        
        .Cell(0, 31).Text = "NET PROFIT"
        .Column(31).Width = 80
        .Column(31).Alignment = cellRightCenter
    End With
End Sub

Sub SALES_ANALYSIS()
    Set CMD = New ADODB.Command
    CMD.ActiveConnection = gconDMIS
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "USP_SALES_ANALYSIS"
    With CMD.Parameters
        '.Append CMD.CreateParameter("@ACCT_CODE", adVarChar, adParamInput, 12, ACCT_CODE)
        .Append CMD.CreateParameter("@JDATE1", adDate, adParamInput, 8, dtFrom)
        .Append CMD.CreateParameter("@JDATE2", adDate, adParamInput, 8, dtTo)
    End With
    Set rsREPORTS = CMD.Execute
    FILLREPORTS
End Sub

Sub FILLREPORTS()
    Grid1.Rows = 1
    Grid1.AutoRedraw = False
    Picture1.Enabled = False
    Picture2.Visible = True
    QTY = 0
    UNITCOST = 0
    TOTALACCESS = 0
    TOTALCOST = 0
    TOTALDISC = 0
    ADDTLDISC = 0
    NETSALES = 0
    SRP = 0
    SRPNETDISC = 0
    SRPNETVAT = 0
    OUTPUT = 0
    GROSSPROFIT = 0
    GP = 0
    FREEINSURANCE = 0
    FREELTO = 0
    FREECHATTEL = 0
    FREEBIES = 0
    TOTALFREE = 0
    COMMISSION = 0
    NETPROFIT = 0
    If Not rsREPORTS.EOF And Not rsREPORTS.BOF Then
        While Not rsREPORTS.EOF
            Grid1.AddItem _
                    rsREPORTS!SONO & Chr(9) & rsREPORTS!CustomerCode & Chr(9) & _
                    rsREPORTS!Customer & Chr(9) & rsREPORTS!Make & Chr(9) & _
                                        rsREPORTS!VINO & Chr(9) & rsREPORTS!prodno & Chr(9) & _
                                        Null2String(rsREPORTS!invoicedate) & Chr(9) & Null2String(rsREPORTS!DATERELEASED) & Chr(9) & rsREPORTS!BANKTERM & Chr(9) & _
                                        rsREPORTS!Bank & Chr(9) & NumericVal(rsREPORTS!QTY) & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!SRP) & Chr(9) & ToDoubleNumber(rsREPORTS!DISCOUNT) & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!SRPNETDISC) & Chr(9) & ToDoubleNumber(rsREPORTS!OUTPUT) & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!SRPNETVAT) & Chr(9) & Null2String(rsREPORTS!CMNO) & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!ADDTLDISC) & Chr(9) & ToDoubleNumber(rsREPORTS!NETSALES) & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!UNITCOST) & Chr(9) & ToDoubleNumber(rsREPORTS!TOTALACCESS) & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!TOTALCOST) & Chr(9) & ToDoubleNumber(rsREPORTS!GROSSPROFIT) & Chr(9) & _
                                        rsREPORTS!GP & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!FREEINSURANCE) & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!FREELTO) & Chr(9) & ToDoubleNumber(rsREPORTS!FREECHATTEL) & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!FREEBIES) & Chr(9) & ToDoubleNumber(rsREPORTS!TOTALFREE) & Chr(9) & _
                                        ToDoubleNumber(rsREPORTS!COMMISSION) & Chr(9) & ToDoubleNumber(rsREPORTS!NETPROFIT)
            QTY = QTY + NumericVal(rsREPORTS!QTY)
            UNITCOST = UNITCOST + rsREPORTS!UNITCOST
            TOTALACCESS = TOTALACCESS + rsREPORTS!TOTALACCESS
            TOTALCOST = TOTALCOST + rsREPORTS!TOTALCOST
            TOTALDISC = TOTALDISC + rsREPORTS!DISCOUNT
            ADDTLDISC = ADDTLDISC + rsREPORTS!ADDTLDISC
            NETSALES = NETSALES + rsREPORTS!NETSALES
            SRP = SRP + rsREPORTS!SRP
            SRPNETDISC = SRPNETDISC + rsREPORTS!SRPNETDISC
            SRPNETVAT = SRPNETVAT + rsREPORTS!SRPNETVAT
            OUTPUT = OUTPUT + rsREPORTS!OUTPUT
            GROSSPROFIT = GROSSPROFIT + rsREPORTS!GROSSPROFIT
            GP = GP + ((rsREPORTS!NETSALES - rsREPORTS!TOTALCOST) / rsREPORTS!TOTALCOST) * 100
            FREEINSURANCE = FREEINSURANCE + rsREPORTS!FREEINSURANCE
            FREELTO = FREELTO + rsREPORTS!FREELTO
            FREECHATTEL = FREECHATTEL + rsREPORTS!FREECHATTEL
            FREEBIES = FREEBIES + rsREPORTS!FREEBIES
            TOTALFREE = TOTALFREE + rsREPORTS!TOTALFREE
            COMMISSION = COMMISSION + rsREPORTS!COMMISSION
            NETPROFIT = NETPROFIT + rsREPORTS!NETPROFIT
            rsREPORTS.MoveNext
            Loading
        Wend
        Grid1.AddItem _
                "" & Chr(9) & "" & Chr(9) & _
                   "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & _
                   "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & _
                   "" & Chr(9) & "" & Chr(9) & _
                   NumericVal(QTY) & Chr(9) & ToDoubleNumber(SRP) & Chr(9) & _
                   ToDoubleNumber(TOTALDISC) & Chr(9) & ToDoubleNumber(SRPNETDISC) & Chr(9) & _
                   ToDoubleNumber(OUTPUT) & Chr(9) & ToDoubleNumber(SRPNETVAT) & Chr(9) & _
                   "" & Chr(9) & _
                   ToDoubleNumber(ADDTLDISC) & Chr(9) & ToDoubleNumber(NETSALES) & Chr(9) & _
                   ToDoubleNumber(UNITCOST) & Chr(9) & ToDoubleNumber(TOTALACCESS) & Chr(9) & _
                   ToDoubleNumber(TOTALCOST) & Chr(9) & _
                   ToDoubleNumber(GROSSPROFIT) & Chr(9) & N2String(ToDoubleNumber(GP)) & "%" & Chr(9) & _
                   ToDoubleNumber(FREEINSURANCE) & Chr(9) & ToDoubleNumber(FREELTO) & Chr(9) & _
                   ToDoubleNumber(FREECHATTEL) & Chr(9) & ToDoubleNumber(FREEBIES) & Chr(9) & _
                   ToDoubleNumber(TOTALFREE) & Chr(9) & ToDoubleNumber(COMMISSION) & Chr(9) & _
                   ToDoubleNumber(NETPROFIT)
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

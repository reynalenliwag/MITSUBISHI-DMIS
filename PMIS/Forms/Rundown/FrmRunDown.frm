VERSION 5.00
Begin VB.Form FrmPMISRunDown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Show Parts Rundown Report in Excel File"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   ForeColor       =   &H00DEDFDE&
   Icon            =   "FrmRunDown.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4110
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   450
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   870
      Width           =   1185
   End
   Begin VB.CommandButton cmdShow 
      Height          =   1425
      Left            =   2070
      Picture         =   "FrmRunDown.frx":0BC2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "View Parts Rundown Report in Excel"
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   630
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "FrmPMISRunDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp                                              As Excel.Application
Dim xlBook                                             As Excel.Workbook
Dim xlSheet                                            As Excel.Worksheet
Dim WithEvents xlBook2                                 As Excel.Workbook
Attribute xlBook2.VB_VarHelpID = -1
Dim MonCol                                             As Long
Dim RETAIL_SALES_HARI_PARTS_GJ, RETAIL_SALES_HARI_PARTS_BP, RETAIL_SALES_HARI_PARTS_COUNTER, RETAIL_SALES_HARI_PARTS_JOBBER, RETAIL_SALES_HARI_PARTS_ACCESSORY, RETAIL_SALES_HARI_PARTS_WARRANTY, RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID, RETAIL_SALES_NON_HARI_PARTS_GJ, RETAIL_SALES_NON_HARI_PARTS_BP, RETAIL_SALES_NON_HARI_PARTS_COUNTER, RETAIL_SALES_NON_HARI_PARTS_JOBBER, RETAIL_SALES_NON_HARI_PARTS_ACCESSORY, RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ, RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP, RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER, RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER, RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY, RETAIL_SALES_OTHER_BRANDS, COST_OF_SALES_HARI_PARTS_GJ, COST_OF_SALES_HARI_PARTS_BP, COST_OF_SALES_HARI_PARTS_COUNTER, COST_OF_SALES_HARI_PARTS_JOBBER, COST_OF_SALES_HARI_PARTS_ACCESSORY, COST_OF_SALES_NON_HARI_PARTS_GJ, COST_OF_SALES_NON_HARI_PARTS_BP As Double
Attribute RETAIL_SALES_HARI_PARTS_BP.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_HARI_PARTS_COUNTER.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_HARI_PARTS_JOBBER.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_HARI_PARTS_ACCESSORY.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_HARI_PARTS_WARRANTY.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_NON_HARI_PARTS_GJ.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_NON_HARI_PARTS_BP.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_NON_HARI_PARTS_COUNTER.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_NON_HARI_PARTS_JOBBER.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_NON_HARI_PARTS_ACCESSORY.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY.VB_VarUserMemId = 1073938436
Attribute RETAIL_SALES_OTHER_BRANDS.VB_VarUserMemId = 1073938436
Attribute COST_OF_SALES_HARI_PARTS_GJ.VB_VarUserMemId = 1073938436
Attribute COST_OF_SALES_HARI_PARTS_BP.VB_VarUserMemId = 1073938436
Attribute COST_OF_SALES_HARI_PARTS_COUNTER.VB_VarUserMemId = 1073938436
Attribute COST_OF_SALES_HARI_PARTS_JOBBER.VB_VarUserMemId = 1073938436
Attribute COST_OF_SALES_HARI_PARTS_ACCESSORY.VB_VarUserMemId = 1073938436
Attribute COST_OF_SALES_NON_HARI_PARTS_GJ.VB_VarUserMemId = 1073938436
Attribute COST_OF_SALES_NON_HARI_PARTS_BP.VB_VarUserMemId = 1073938436
Dim COST_OF_SALES_NON_HARI_PARTS_COUNTER, COST_OF_SALES_NON_HARI_PARTS_JOBBER, COST_OF_SALES_NON_HARI_PARTS_ACCESSORY, COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ, COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP, COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER, COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER, COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY, COST_OF_SALES_OTHER_BRANDS, BI_HARI_GJ, BI_HARI_BP, BI_HARI_ACCESSORY, BI_NON_HARI_GJ, BI_NON_HARI_BP, BI_NON_HARI_ACCESSORY, BI_OTHER_BRANDS, PURCHASES_HARI_GJ, PURCHASES_HARI_BP, PURCHASES_HARI_ACCESSORY, PURCHASES_NON_HARI_GJ, PURCHASES_NON_HARI_BP, PURCHASES_NON_HARI_ACCESSORY, PURCHASES_OTHER_BRANDS, ADJUSTMENTS, EI_HARI_GJ, EI_HARI_BP, EI_HARI_ACCESSORY As Double
Attribute COST_OF_SALES_NON_HARI_PARTS_COUNTER.VB_VarUserMemId = 1073938461
Attribute COST_OF_SALES_NON_HARI_PARTS_JOBBER.VB_VarUserMemId = 1073938461
Attribute COST_OF_SALES_NON_HARI_PARTS_ACCESSORY.VB_VarUserMemId = 1073938461
Attribute COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ.VB_VarUserMemId = 1073938461
Attribute COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP.VB_VarUserMemId = 1073938461
Attribute COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER.VB_VarUserMemId = 1073938461
Attribute COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER.VB_VarUserMemId = 1073938461
Attribute COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY.VB_VarUserMemId = 1073938461
Attribute COST_OF_SALES_OTHER_BRANDS.VB_VarUserMemId = 1073938461
Attribute BI_HARI_GJ.VB_VarUserMemId = 1073938461
Attribute BI_HARI_BP.VB_VarUserMemId = 1073938461
Attribute BI_HARI_ACCESSORY.VB_VarUserMemId = 1073938461
Attribute BI_NON_HARI_GJ.VB_VarUserMemId = 1073938461
Attribute BI_NON_HARI_BP.VB_VarUserMemId = 1073938461
Attribute BI_NON_HARI_ACCESSORY.VB_VarUserMemId = 1073938461
Attribute BI_OTHER_BRANDS.VB_VarUserMemId = 1073938461
Attribute PURCHASES_HARI_GJ.VB_VarUserMemId = 1073938461
Attribute PURCHASES_HARI_BP.VB_VarUserMemId = 1073938461
Attribute PURCHASES_HARI_ACCESSORY.VB_VarUserMemId = 1073938461
Attribute PURCHASES_NON_HARI_GJ.VB_VarUserMemId = 1073938461
Attribute PURCHASES_NON_HARI_BP.VB_VarUserMemId = 1073938461
Attribute PURCHASES_NON_HARI_ACCESSORY.VB_VarUserMemId = 1073938461
Attribute PURCHASES_OTHER_BRANDS.VB_VarUserMemId = 1073938461
Attribute ADJUSTMENTS.VB_VarUserMemId = 1073938461
Attribute EI_HARI_GJ.VB_VarUserMemId = 1073938461
Attribute EI_HARI_BP.VB_VarUserMemId = 1073938461
Attribute EI_HARI_ACCESSORY.VB_VarUserMemId = 1073938461
Dim EI_NON_HARI_GJ, EI_NON_HARI_BP, EI_NON_HARI_ACCESSORY, EI_OTHER_BRANDS, WSC_NUMBER_ORDER_SLIP_RECEIVED, WSC_COMPLETELY_SERVE_ORDER_SLIP, WSC_NUMBER_LINE_ITEM_ORDERED, WSC_COMPLETELY_SERVE_LINE_ITEM, OTC_NUMBER_ORDER_SLIP_RECEIVED, OTC_COMPLETELY_SERVE_ORDER_SLIP, OTC_NUMBER_LINE_ITEM_ORDERED, OTC_COMPLETELY_SERVE_LINE_ITEM, TOTAL_LINES_ORDERED, TOTAL_QTY_ORDERED, TOTAL_LINES_SERVED, TOTAL_QTY_SERVED, ORDERED_PARTS_FILL, ORDERED_PARTS_KILL, ORDERED_PARTS_WARRANTY, BACK_ORDER_PARTS_WARRANTY, BACK_ORDER_PARTS_REGULAR As Double
Attribute EI_NON_HARI_GJ.VB_VarUserMemId = 1073938488
Attribute EI_NON_HARI_BP.VB_VarUserMemId = 1073938488
Attribute EI_NON_HARI_ACCESSORY.VB_VarUserMemId = 1073938488
Attribute EI_OTHER_BRANDS.VB_VarUserMemId = 1073938488
Attribute WSC_NUMBER_ORDER_SLIP_RECEIVED.VB_VarUserMemId = 1073938488
Attribute WSC_COMPLETELY_SERVE_ORDER_SLIP.VB_VarUserMemId = 1073938488
Attribute WSC_NUMBER_LINE_ITEM_ORDERED.VB_VarUserMemId = 1073938488
Attribute WSC_COMPLETELY_SERVE_LINE_ITEM.VB_VarUserMemId = 1073938488
Attribute OTC_NUMBER_ORDER_SLIP_RECEIVED.VB_VarUserMemId = 1073938488
Attribute OTC_COMPLETELY_SERVE_ORDER_SLIP.VB_VarUserMemId = 1073938488
Attribute OTC_NUMBER_LINE_ITEM_ORDERED.VB_VarUserMemId = 1073938488
Attribute OTC_COMPLETELY_SERVE_LINE_ITEM.VB_VarUserMemId = 1073938488
Attribute TOTAL_LINES_ORDERED.VB_VarUserMemId = 1073938488
Attribute TOTAL_QTY_ORDERED.VB_VarUserMemId = 1073938488
Attribute TOTAL_LINES_SERVED.VB_VarUserMemId = 1073938488
Attribute TOTAL_QTY_SERVED.VB_VarUserMemId = 1073938488
Attribute ORDERED_PARTS_FILL.VB_VarUserMemId = 1073938488
Attribute ORDERED_PARTS_KILL.VB_VarUserMemId = 1073938488
Attribute ORDERED_PARTS_WARRANTY.VB_VarUserMemId = 1073938488
Attribute BACK_ORDER_PARTS_WARRANTY.VB_VarUserMemId = 1073938488
Attribute BACK_ORDER_PARTS_REGULAR.VB_VarUserMemId = 1073938488

Private Sub ShowExcel()
On Error GoTo CloseExcel
    Screen.MousePointer = 11
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\PartsRundown.xls")
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Cells(2, 4) = "Year : " & cboYear
    Dim rsPRR_VIEW_MONTH                               As ADODB.Recordset
    Dim PRR_MONTH                                      As Integer
    Dim NEWMONTH, LASTMONTH                            As Integer
    For PRR_MONTH = 1 To 12
        MonCol = PRR_MONTH + 3
        RETAIL_SALES_HARI_PARTS_GJ = 0: RETAIL_SALES_HARI_PARTS_BP = 0: RETAIL_SALES_HARI_PARTS_COUNTER = 0: RETAIL_SALES_HARI_PARTS_JOBBER = 0: RETAIL_SALES_HARI_PARTS_ACCESSORY = 0
        RETAIL_SALES_HARI_PARTS_WARRANTY = 0: RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID = 0
        RETAIL_SALES_NON_HARI_PARTS_GJ = 0: RETAIL_SALES_NON_HARI_PARTS_BP = 0: RETAIL_SALES_NON_HARI_PARTS_COUNTER = 0: RETAIL_SALES_NON_HARI_PARTS_JOBBER = 0: RETAIL_SALES_NON_HARI_PARTS_ACCESSORY = 0
        RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ = 0: RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP = 0: RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER = 0: RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER = 0: RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = 0
        RETAIL_SALES_OTHER_BRANDS = 0

        COST_OF_SALES_HARI_PARTS_GJ = 0: COST_OF_SALES_HARI_PARTS_BP = 0: COST_OF_SALES_HARI_PARTS_COUNTER = 0: COST_OF_SALES_HARI_PARTS_JOBBER = 0: COST_OF_SALES_HARI_PARTS_ACCESSORY = 0
        COST_OF_SALES_NON_HARI_PARTS_GJ = 0: COST_OF_SALES_NON_HARI_PARTS_BP = 0: COST_OF_SALES_NON_HARI_PARTS_COUNTER = 0: COST_OF_SALES_NON_HARI_PARTS_JOBBER = 0: COST_OF_SALES_NON_HARI_PARTS_ACCESSORY = 0
        COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ = 0: COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP = 0: COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER = 0: COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER = 0: COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = 0
        COST_OF_SALES_OTHER_BRANDS = 0

        BI_HARI_GJ = 0: BI_HARI_BP = 0: BI_HARI_ACCESSORY = 0: BI_NON_HARI_GJ = 0: BI_NON_HARI_BP = 0: BI_NON_HARI_ACCESSORY = 0
        BI_OTHER_BRANDS = 0
        PURCHASES_HARI_GJ = 0: PURCHASES_HARI_BP = 0: PURCHASES_HARI_ACCESSORY = 0: PURCHASES_NON_HARI_GJ = 0: PURCHASES_NON_HARI_BP = 0: PURCHASES_NON_HARI_ACCESSORY = 0
        PURCHASES_OTHER_BRANDS = 0

        ADJUSTMENTS = 0
        EI_HARI_GJ = 0: EI_HARI_BP = 0: EI_HARI_ACCESSORY = 0: EI_NON_HARI_GJ = 0: EI_NON_HARI_BP = 0: EI_NON_HARI_ACCESSORY = 0
        EI_OTHER_BRANDS = 0

        WSC_NUMBER_ORDER_SLIP_RECEIVED = 0: WSC_COMPLETELY_SERVE_ORDER_SLIP = 0: WSC_NUMBER_LINE_ITEM_ORDERED = 0: WSC_COMPLETELY_SERVE_LINE_ITEM = 0
        OTC_NUMBER_ORDER_SLIP_RECEIVED = 0: OTC_COMPLETELY_SERVE_ORDER_SLIP = 0: OTC_NUMBER_LINE_ITEM_ORDERED = 0: OTC_COMPLETELY_SERVE_LINE_ITEM = 0
        TOTAL_LINES_ORDERED = 0: TOTAL_QTY_ORDERED = 0: TOTAL_LINES_SERVED = 0: TOTAL_QTY_SERVED = 0
        ORDERED_PARTS_FILL = 0: ORDERED_PARTS_KILL = 0: ORDERED_PARTS_WARRANTY = 0
        BACK_ORDER_PARTS_WARRANTY = 0: BACK_ORDER_PARTS_REGULAR = 0
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("SELECT * FROM PMIS_vw_PRR_VIEW WHERE MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear.Text)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            rsPRR_VIEW_MONTH.MoveFirst
            Do While Not rsPRR_VIEW_MONTH.EOF
                If rsPRR_VIEW_MONTH!NON_HARI = "N" Then
                    If rsPRR_VIEW_MONTH!Type = "P" Then
                        If rsPRR_VIEW_MONTH!SALES_ORIGIN = "W" Then
                            RETAIL_SALES_HARI_PARTS_COUNTER = RETAIL_SALES_HARI_PARTS_COUNTER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                            COST_OF_SALES_HARI_PARTS_COUNTER = COST_OF_SALES_HARI_PARTS_COUNTER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                        ElseIf rsPRR_VIEW_MONTH!SALES_ORIGIN = "J" Then
                            RETAIL_SALES_HARI_PARTS_JOBBER = RETAIL_SALES_HARI_PARTS_JOBBER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                            COST_OF_SALES_HARI_PARTS_JOBBER = COST_OF_SALES_HARI_PARTS_JOBBER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                        Else
                            If rsPRR_VIEW_MONTH!SI_TYPE = "G" Then
                                RETAIL_SALES_HARI_PARTS_GJ = RETAIL_SALES_HARI_PARTS_GJ + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                COST_OF_SALES_HARI_PARTS_GJ = COST_OF_SALES_HARI_PARTS_GJ + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                            End If
                            If rsPRR_VIEW_MONTH!SI_TYPE = "B" Then
                                RETAIL_SALES_HARI_PARTS_BP = RETAIL_SALES_HARI_PARTS_BP + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                COST_OF_SALES_HARI_PARTS_BP = COST_OF_SALES_HARI_PARTS_BP + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                            End If
                        
                            'If rsPRR_VIEW_MONTH!PAY_CLASS = "W" And rsPRR_VIEW_MONTH!NON_HARI = "N" Then
                            '    RETAIL_SALES_HARI_PARTS_WARRANTY = RETAIL_SALES_HARI_PARTS_WARRANTY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                            'End If
                            'If rsPRR_VIEW_MONTH!PAY_CLASS = "C" And rsPRR_VIEW_MONTH!NON_HARI = "N" Then
                            '    RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID = RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                            'End If
            
                        End If
                    Else
                        RETAIL_SALES_HARI_PARTS_ACCESSORY = RETAIL_SALES_HARI_PARTS_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                        COST_OF_SALES_HARI_PARTS_ACCESSORY = COST_OF_SALES_HARI_PARTS_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                    End If
                    
                Else
                    If rsPRR_VIEW_MONTH!Type = "P" Then
                        If rsPRR_VIEW_MONTH!SALES_ORIGIN = "W" Then
                            RETAIL_SALES_NON_HARI_PARTS_COUNTER = RETAIL_SALES_NON_HARI_PARTS_COUNTER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                            COST_OF_SALES_NON_HARI_PARTS_COUNTER = COST_OF_SALES_NON_HARI_PARTS_COUNTER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                        ElseIf rsPRR_VIEW_MONTH!SALES_ORIGIN = "J" Then
                            RETAIL_SALES_NON_HARI_PARTS_JOBBER = RETAIL_SALES_NON_HARI_PARTS_JOBBER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                            COST_OF_SALES_NON_HARI_PARTS_JOBBER = COST_OF_SALES_NON_HARI_PARTS_JOBBER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                        Else
                            If rsPRR_VIEW_MONTH!SI_TYPE = "G" Then
                                RETAIL_SALES_NON_HARI_PARTS_GJ = RETAIL_SALES_NON_HARI_PARTS_GJ + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                COST_OF_SALES_NON_HARI_PARTS_GJ = COST_OF_SALES_NON_HARI_PARTS_GJ + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                            End If
                            If rsPRR_VIEW_MONTH!SI_TYPE = "B" Then
                                RETAIL_SALES_NON_HARI_PARTS_BP = RETAIL_SALES_NON_HARI_PARTS_BP + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                COST_OF_SALES_NON_HARI_PARTS_BP = COST_OF_SALES_NON_HARI_PARTS_BP + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                            End If
                        End If
        
'                        If rsPRR_VIEW_MONTH!SI_TYPE = "G" And rsPRR_VIEW_MONTH!NON_HARI = "C" Then
'                            RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
'                            COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
'                        End If
'                        If rsPRR_VIEW_MONTH!SI_TYPE = "B" And rsPRR_VIEW_MONTH!NON_HARI = "C" Then
'                            RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
'                            COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
'                        End If
'                        If rsPRR_VIEW_MONTH!SI_TYPE = "S" And rsPRR_VIEW_MONTH!SALES_ORIGIN = "W" And rsPRR_VIEW_MONTH!NON_HARI = "C" Then
'                            RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
'                            COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
'                        End If
'                        If rsPRR_VIEW_MONTH!SI_TYPE = "S" And rsPRR_VIEW_MONTH!SALES_ORIGIN = "J" And rsPRR_VIEW_MONTH!NON_HARI = "C" Then
'                            RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
'                            COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
'                        End If
'                        If rsPRR_VIEW_MONTH!SI_TYPE = "S" And rsPRR_VIEW_MONTH!SALES_ORIGIN = "M" And rsPRR_VIEW_MONTH!NON_HARI = "C" Then
'                            RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
'                            COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
'                        End If
'                        If rsPRR_VIEW_MONTH!NON_HARI = "O" Then
'                            RETAIL_SALES_OTHER_BRANDS = RETAIL_SALES_OTHER_BRANDS + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
'                            COST_OF_SALES_OTHER_BRANDS = COST_OF_SALES_OTHER_BRANDS + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
'                        End If
                    Else
                    MsgBox rsPRR_VIEW_MONTH!Type
                        RETAIL_SALES_NON_HARI_PARTS_ACCESSORY = RETAIL_SALES_NON_HARI_PARTS_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                        COST_OF_SALES_NON_HARI_PARTS_ACCESSORY = COST_OF_SALES_NON_HARI_PARTS_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                        Stop
                    End If
                End If
                


                rsPRR_VIEW_MONTH.MoveNext
            Loop
        End If

                xlSheet.Cells(8, MonCol) = RETAIL_SALES_HARI_PARTS_GJ
                xlSheet.Cells(9, MonCol) = RETAIL_SALES_HARI_PARTS_BP
                xlSheet.Cells(10, MonCol) = RETAIL_SALES_HARI_PARTS_COUNTER
                xlSheet.Cells(11, MonCol) = RETAIL_SALES_HARI_PARTS_JOBBER
                xlSheet.Cells(12, MonCol) = RETAIL_SALES_HARI_PARTS_ACCESSORY

                xlSheet.Cells(15, MonCol) = RETAIL_SALES_HARI_PARTS_WARRANTY
                xlSheet.Cells(16, MonCol) = RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID

                xlSheet.Cells(20, MonCol) = RETAIL_SALES_NON_HARI_PARTS_GJ
                xlSheet.Cells(21, MonCol) = RETAIL_SALES_NON_HARI_PARTS_BP
                xlSheet.Cells(22, MonCol) = RETAIL_SALES_NON_HARI_PARTS_COUNTER
                xlSheet.Cells(23, MonCol) = RETAIL_SALES_NON_HARI_PARTS_JOBBER
                xlSheet.Cells(24, MonCol) = RETAIL_SALES_NON_HARI_PARTS_ACCESSORY

                xlSheet.Cells(26, MonCol) = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ
                xlSheet.Cells(27, MonCol) = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP
                xlSheet.Cells(28, MonCol) = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER
                xlSheet.Cells(29, MonCol) = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER
                xlSheet.Cells(30, MonCol) = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY

                xlSheet.Cells(35, MonCol) = RETAIL_SALES_OTHER_BRANDS

                xlSheet.Cells(41, MonCol) = COST_OF_SALES_HARI_PARTS_GJ
                xlSheet.Cells(42, MonCol) = COST_OF_SALES_HARI_PARTS_BP
                xlSheet.Cells(43, MonCol) = COST_OF_SALES_HARI_PARTS_COUNTER
                xlSheet.Cells(44, MonCol) = COST_OF_SALES_HARI_PARTS_JOBBER
                xlSheet.Cells(45, MonCol) = COST_OF_SALES_HARI_PARTS_ACCESSORY

                xlSheet.Cells(49, MonCol) = COST_OF_SALES_NON_HARI_PARTS_GJ
                xlSheet.Cells(50, MonCol) = COST_OF_SALES_NON_HARI_PARTS_BP
                xlSheet.Cells(51, MonCol) = COST_OF_SALES_NON_HARI_PARTS_COUNTER
                xlSheet.Cells(52, MonCol) = COST_OF_SALES_NON_HARI_PARTS_JOBBER
                xlSheet.Cells(53, MonCol) = COST_OF_SALES_NON_HARI_PARTS_ACCESSORY

                xlSheet.Cells(55, MonCol) = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ
                xlSheet.Cells(56, MonCol) = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP
                xlSheet.Cells(57, MonCol) = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER
                xlSheet.Cells(58, MonCol) = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER
                xlSheet.Cells(59, MonCol) = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY
                xlSheet.Cells(64, MonCol) = COST_OF_SALES_OTHER_BRANDS

'        Set rsPRR_VIEW_MONTH = Nothing
'        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
'        If PRR_MONTH = 1 Then
'            Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,STOCK_TYPE,MONTH(DATE_GEN) AS MONTH_DATE, YEAR(DATE_GEN) AS YEAR_DATE from PMIS_vw_PRR_BEG_INVENTORY_HARI where MONTH(DATE_GEN) = 12 AND YEAR(DATE_GEN) = " & NumericVal(cboYear) - 1 & " ORDER BY DATE_GEN ASC")
'        Else
'            Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,STOCK_TYPE,MONTH(DATE_GEN) AS MONTH_DATE, YEAR(DATE_GEN) AS YEAR_DATE from PMIS_vw_PRR_BEG_INVENTORY_HARI where MONTH(DATE_GEN) + 1 = " & NumericVal(PRR_MONTH) & " AND YEAR(DATE_GEN) = " & cboYear & " ORDER BY DATE_GEN ASC")
'        End If
'        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.EOF Then
'            rsPRR_VIEW_MONTH.MoveFirst
'            Do While Not rsPRR_VIEW_MONTH.EOF
'                If Null2String(rsPRR_VIEW_MONTH!STOCK_TYPE) = "GJ" Then
'                    BI_HARI_GJ = NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                If Null2String(rsPRR_VIEW_MONTH!STOCK_TYPE) = "BP" Then
'                    BI_HARI_BP = NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                If Null2String(rsPRR_VIEW_MONTH!STOCK_TYPE) = "AC" Then
'                    BI_HARI_ACCESSORY = NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                xlSheet.Cells(70, MonCol) = BI_HARI_GJ
'                xlSheet.Cells(71, MonCol) = BI_HARI_BP
'                xlSheet.Cells(72, MonCol) = BI_HARI_ACCESSORY
'                rsPRR_VIEW_MONTH.MoveNext
'            Loop
'        End If
'        Set rsPRR_VIEW_MONTH = Nothing
'        BI_OTHER_BRANDS = NumericVal(0)
'        xlSheet.Cells(78, MonCol) = BI_OTHER_BRANDS
'
'        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
'        If PRR_MONTH = 1 Then
'            Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,STOCK_TYPE,MONTH(DATE_GEN) AS MONTH_DATE, YEAR(DATE_GEN) AS YEAR_DATE from PMIS_vw_PRR_BEG_INVENTORY_NON_HARI where MONTH(DATE_GEN) = 12 AND YEAR(DATE_GEN) = " & NumericVal(cboYear) - 1 & " ORDER BY DATE_GEN ASC")
'        Else
'            Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,STOCK_TYPE,MONTH(DATE_GEN) AS MONTH_DATE, YEAR(DATE_GEN) AS YEAR_DATE from PMIS_vw_PRR_BEG_INVENTORY_NON_HARI where MONTH(DATE_GEN) + 1 = " & NumericVal(PRR_MONTH) & " AND YEAR(DATE_GEN) = " & cboYear & " ORDER BY DATE_GEN ASC")
'        End If
'        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.EOF Then
'            rsPRR_VIEW_MONTH.MoveFirst
'            Do While Not rsPRR_VIEW_MONTH.EOF
'                If Null2String(rsPRR_VIEW_MONTH!STOCK_TYPE) = "GJ" Then
'                    BI_NON_HARI_GJ = NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                If Null2String(rsPRR_VIEW_MONTH!STOCK_TYPE) = "BP" Then
'                    BI_NON_HARI_BP = NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                If Null2String(rsPRR_VIEW_MONTH!STOCK_TYPE) = "AC" Then
'                    BI_NON_HARI_ACCESSORY = NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                xlSheet.Cells(74, MonCol) = BI_NON_HARI_GJ
'                xlSheet.Cells(75, MonCol) = BI_NON_HARI_BP
'                xlSheet.Cells(76, MonCol) = BI_NON_HARI_ACCESSORY
'                rsPRR_VIEW_MONTH.MoveNext
'            Loop
'        End If
'        Set rsPRR_VIEW_MONTH = Nothing
'        NEWMONTH = 0
'        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
'        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,LTRIM(RTRIM(STOCK_TYPE)) AS STOCKTYPE,MONTH_TRAN AS MONTH_DATE, YEAR_TRAN AS YEAR_DATE from PMIS_vw_PRR_PURCHASES_HARI where MONTH_TRAN = " & PRR_MONTH & " AND YEAR_TRAN=" & cboYear & " ORDER BY MONTH_TRAN ASC")
'        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.EOF Then
'            rsPRR_VIEW_MONTH.MoveFirst
'            Do While Not rsPRR_VIEW_MONTH.EOF
'                LASTMONTH = rsPRR_VIEW_MONTH!MONTH_DATE
'                If NEWMONTH = 0 Or LASTMONTH <> NEWMONTH Then
'                    NEWMONTH = LASTMONTH
'                    PURCHASES_HARI_GJ = NumericVal(0)
'                    PURCHASES_HARI_BP = NumericVal(0)
'                    PURCHASES_HARI_ACCESSORY = NumericVal(0)
'                End If
'                If Null2String(rsPRR_VIEW_MONTH!STOCKTYPE) = "GJ" Then
'                    PURCHASES_HARI_GJ = PURCHASES_HARI_GJ + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                If Null2String(rsPRR_VIEW_MONTH!STOCKTYPE) = "BP" Then
'                    PURCHASES_HARI_BP = PURCHASES_HARI_BP + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                If Null2String(rsPRR_VIEW_MONTH!STOCKTYPE) = "AC" Then
'                    PURCHASES_HARI_ACCESSORY = PURCHASES_HARI_ACCESSORY + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                xlSheet.Cells(83, MonCol) = PURCHASES_HARI_GJ
'                xlSheet.Cells(84, MonCol) = PURCHASES_HARI_BP
'                xlSheet.Cells(85, MonCol) = PURCHASES_HARI_ACCESSORY
'                rsPRR_VIEW_MONTH.MoveNext
'            Loop
'        End If
'        Set rsPRR_VIEW_MONTH = Nothing
'        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
'        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,LTRIM(RTRIM(STOCK_TYPE)) AS STOCKTYPE,MONTH_TRAN AS MONTH_DATE, YEAR_TRAN AS YEAR_DATE from PMIS_vw_PRR_PURCHASES_NON_HARI where MONTH_TRAN = " & PRR_MONTH & " AND YEAR_TRAN = " & cboYear & " ORDER BY MONTH_TRAN ASC")
'        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.EOF Then
'            rsPRR_VIEW_MONTH.MoveFirst
'            Do While Not rsPRR_VIEW_MONTH.EOF
'                LASTMONTH = rsPRR_VIEW_MONTH!MONTH_DATE
'                If NEWMONTH = 0 Or LASTMONTH <> NEWMONTH Then
'                    NEWMONTH = LASTMONTH
'                    PURCHASES_NON_HARI_GJ = NumericVal(0)
'                    PURCHASES_NON_HARI_BP = NumericVal(0)
'                    PURCHASES_NON_HARI_ACCESSORY = NumericVal(0)
'                End If
'                If Null2String(rsPRR_VIEW_MONTH!STOCKTYPE) = "GJ" Then
'                    PURCHASES_NON_HARI_GJ = PURCHASES_NON_HARI_GJ + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                If Null2String(rsPRR_VIEW_MONTH!STOCKTYPE) = "BP" Then
'                    PURCHASES_NON_HARI_BP = PURCHASES_NON_HARI_BP + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                If Null2String(rsPRR_VIEW_MONTH!STOCKTYPE) = "AC" Then
'                    PURCHASES_NON_HARI_ACCESSORY = PURCHASES_NON_HARI_ACCESSORY + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
'                End If
'                xlSheet.Cells(87, MonCol) = PURCHASES_NON_HARI_GJ
'                xlSheet.Cells(88, MonCol) = PURCHASES_NON_HARI_BP
'                xlSheet.Cells(89, MonCol) = PURCHASES_NON_HARI_ACCESSORY
'                rsPRR_VIEW_MONTH.MoveNext
'            Loop
'        End If
        Set rsPRR_VIEW_MONTH = Nothing
        PURCHASES_OTHER_BRANDS = NumericVal(0)
        xlSheet.Cells(91, MonCol) = PURCHASES_OTHER_BRANDS

        ADJUSTMENTS = NumericVal(0)
        xlSheet.Cells(94, MonCol) = ADJUSTMENTS

        'EI = BI + PURCHASE +/- ADJUSTMENTS - COST OF GOODS SOLD
        EI_HARI_GJ = BI_HARI_GJ + PURCHASES_HARI_GJ + ADJUSTMENTS - COST_OF_SALES_HARI_PARTS_GJ
        EI_HARI_BP = BI_HARI_BP + PURCHASES_HARI_BP + ADJUSTMENTS - COST_OF_SALES_HARI_PARTS_BP
        EI_HARI_ACCESSORY = BI_HARI_ACCESSORY + PURCHASES_HARI_ACCESSORY + ADJUSTMENTS - COST_OF_SALES_HARI_PARTS_ACCESSORY

        EI_NON_HARI_GJ = BI_NON_HARI_GJ + PURCHASES_NON_HARI_GJ + ADJUSTMENTS - COST_OF_SALES_NON_HARI_PARTS_GJ
        EI_NON_HARI_BP = BI_NON_HARI_BP + PURCHASES_NON_HARI_BP + ADJUSTMENTS - COST_OF_SALES_NON_HARI_PARTS_BP
        EI_NON_HARI_ACCESSORY = BI_NON_HARI_ACCESSORY + PURCHASES_NON_HARI_ACCESSORY + ADJUSTMENTS - COST_OF_SALES_NON_HARI_PARTS_ACCESSORY

        xlSheet.Cells(98, MonCol) = EI_HARI_GJ
        xlSheet.Cells(99, MonCol) = EI_HARI_BP
        xlSheet.Cells(100, MonCol) = EI_HARI_ACCESSORY

        xlSheet.Cells(102, MonCol) = EI_NON_HARI_GJ
        xlSheet.Cells(103, MonCol) = EI_NON_HARI_BP
        xlSheet.Cells(104, MonCol) = EI_NON_HARI_ACCESSORY

        EI_OTHER_BRANDS = NumericVal(0)
        xlSheet.Cells(106, MonCol) = EI_OTHER_BRANDS

'        Set rsPRR_VIEW_MONTH = Nothing
'        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
'        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select * from PMIS_vw_Prr_Demand where MONTH_TRAN = " & PRR_MONTH & " AND YEAR_TRAN = " & cboYear & " ORDER BY MONTH_TRAN ASC")
'        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.EOF Then
'            'PARTNO, DATE_GEN, C_REQUESTED, C_SERVED, C_UNSERVED, C_FILLRATE, D_ORDERED, D_SERVED, D_UNSERVED, D_BACKORDER,
'            'D_FILLRATE , D_EMERGENCY_PO, D_ONORDER, S_REQUESTED, S_SERVED, S_UNSERVED, S_BACKORDER, S_FILLRATE
'            rsPRR_VIEW_MONTH.MoveFirst
'            Do While Not rsPRR_VIEW_MONTH.EOF
'                LASTMONTH = rsPRR_VIEW_MONTH!MONTH_TRAN
'                If NEWMONTH = 0 Or LASTMONTH <> NEWMONTH Then
'                    NEWMONTH = LASTMONTH
'                    WSC_NUMBER_ORDER_SLIP_RECEIVED = NumericVal(0)
'                    WSC_COMPLETELY_SERVE_ORDER_SLIP = NumericVal(0)
'                    WSC_NUMBER_LINE_ITEM_ORDERED = NumericVal(0)
'                    WSC_COMPLETELY_SERVE_LINE_ITEM = NumericVal(0)
'                    OTC_NUMBER_ORDER_SLIP_RECEIVED = NumericVal(0)
'                    OTC_COMPLETELY_SERVE_ORDER_SLIP = NumericVal(0)
'                    OTC_NUMBER_LINE_ITEM_ORDERED = NumericVal(0)
'                    OTC_COMPLETELY_SERVE_LINE_ITEM = NumericVal(0)
'                End If
'                'WSC_NUMBER_ORDER_SLIP_RECEIVED = WSC_NUMBER_ORDER_SLIP_RECEIVED + N2Str2Zero(rsPRR_VIEW_MONTH!S_REQUESTED)
'                'WSC_COMPLETELY_SERVE_ORDER_SLIP = WSC_COMPLETELY_SERVE_ORDER_SLIP + N2Str2Zero(rsPRR_VIEW_MONTH!S_SERVED)
'                WSC_NUMBER_LINE_ITEM_ORDERED = WSC_NUMBER_LINE_ITEM_ORDERED + N2Str2Zero(rsPRR_VIEW_MONTH!S_REQUESTED)
'                WSC_COMPLETELY_SERVE_LINE_ITEM = WSC_COMPLETELY_SERVE_LINE_ITEM + N2Str2Zero(rsPRR_VIEW_MONTH!S_SERVED)
'                'OTC_NUMBER_ORDER_SLIP_RECEIVED = OTC_NUMBER_ORDER_SLIP_RECEIVED + N2Str2Zero(rsPRR_VIEW_MONTH!S_REQUESTED)
'                'OTC_COMPLETELY_SERVE_ORDER_SLIP = OTC_COMPLETELY_SERVE_ORDER_SLIP + N2Str2Zero(rsPRR_VIEW_MONTH!S_REQUESTED)
'                OTC_NUMBER_LINE_ITEM_ORDERED = OTC_NUMBER_LINE_ITEM_ORDERED + N2Str2Zero(rsPRR_VIEW_MONTH!C_REQUESTED)
'                OTC_COMPLETELY_SERVE_LINE_ITEM = OTC_COMPLETELY_SERVE_LINE_ITEM + N2Str2Zero(rsPRR_VIEW_MONTH!C_SERVED)
'                xlSheet.Cells(117, MonCol) = WSC_NUMBER_ORDER_SLIP_RECEIVED
'                xlSheet.Cells(118, MonCol) = WSC_COMPLETELY_SERVE_ORDER_SLIP
'                xlSheet.Cells(119, MonCol) = WSC_NUMBER_LINE_ITEM_ORDERED
'                xlSheet.Cells(120, MonCol) = WSC_COMPLETELY_SERVE_LINE_ITEM
'                xlSheet.Cells(124, MonCol) = OTC_NUMBER_ORDER_SLIP_RECEIVED
'                xlSheet.Cells(125, MonCol) = OTC_COMPLETELY_SERVE_ORDER_SLIP
'                xlSheet.Cells(126, MonCol) = OTC_NUMBER_LINE_ITEM_ORDERED
'                xlSheet.Cells(127, MonCol) = OTC_COMPLETELY_SERVE_LINE_ITEM
'                rsPRR_VIEW_MONTH.MoveNext
'            Loop
'        End If

        'WSC_NUMBER_ORDER_SLIP_RECEIVED = NumericVal(0)
        'WSC_COMPLETELY_SERVE_ORDER_SLIP = NumericVal(0)
        'WSC_NUMBER_LINE_ITEM_ORDERED = NumericVal(0)
        'WSC_COMPLETELY_SERVE_LINE_ITEM = NumericVal(0)
        'xlSheet.Cells(117, MonCol) = WSC_NUMBER_ORDER_SLIP_RECEIVED
        'xlSheet.Cells(118, MonCol) = WSC_COMPLETELY_SERVE_ORDER_SLIP
        'xlSheet.Cells(119, MonCol) = WSC_NUMBER_LINE_ITEM_ORDERED
        'xlSheet.Cells(120, MonCol) = WSC_COMPLETELY_SERVE_LINE_ITEM

        'OTC_NUMBER_ORDER_SLIP_RECEIVED = NumericVal(0)
        'OTC_COMPLETELY_SERVE_ORDER_SLIP = NumericVal(0)
        'OTC_NUMBER_LINE_ITEM_ORDERED = NumericVal(0)
        'OTC_COMPLETELY_SERVE_LINE_ITEM = NumericVal(0)
        'xlSheet.Cells(124, MonCol) = OTC_NUMBER_ORDER_SLIP_RECEIVED
        'xlSheet.Cells(125, MonCol) = OTC_COMPLETELY_SERVE_ORDER_SLIP
        'xlSheet.Cells(126, MonCol) = OTC_NUMBER_LINE_ITEM_ORDERED
        'xlSheet.Cells(127, MonCol) = OTC_COMPLETELY_SERVE_LINE_ITEM

        TOTAL_LINES_ORDERED = NumericVal(0)
        TOTAL_QTY_ORDERED = NumericVal(0)
        TOTAL_LINES_SERVED = NumericVal(0)
        TOTAL_QTY_SERVED = NumericVal(0)
        xlSheet.Cells(132, MonCol) = TOTAL_LINES_ORDERED
        xlSheet.Cells(133, MonCol) = TOTAL_QTY_ORDERED
        xlSheet.Cells(134, MonCol) = TOTAL_LINES_SERVED
        xlSheet.Cells(135, MonCol) = TOTAL_QTY_SERVED

        ORDERED_PARTS_FILL = NumericVal(0)
        ORDERED_PARTS_KILL = NumericVal(0)

        ORDERED_PARTS_WARRANTY = NumericVal(0)
        xlSheet.Cells(139, MonCol) = ORDERED_PARTS_FILL
        xlSheet.Cells(140, MonCol) = ORDERED_PARTS_KILL
        xlSheet.Cells(141, MonCol) = ORDERED_PARTS_WARRANTY

        BACK_ORDER_PARTS_WARRANTY = NumericVal(0)
        BACK_ORDER_PARTS_REGULAR = NumericVal(0)
        xlSheet.Cells(144, MonCol) = BACK_ORDER_PARTS_WARRANTY
        xlSheet.Cells(145, MonCol) = BACK_ORDER_PARTS_REGULAR
    Next
    xlApp.Visible = True
    Set xlApp = Nothing
    Screen.MousePointer = 0
CloseExcel:
    Set xlApp = Nothing
End Sub

Private Sub cmdShow_Click()
    If Function_Access(LOGID, "Acess_PRINT", "REPORTS PARTS RUNDOWN EXCEL") = False Then Exit Sub
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:
    Call ShowExcel
    LogAudit "V", "PARTS RUNDOWN", cboYear

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    FillcboYear cboYear
    cboYear.Text = Year(LOGDATE)
End Sub


VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Begin VB.Form frmPMISReports_PrintPartsRunDown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Parts Run-Down Report"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "PrintPartsRunDown.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   4305
   Begin VB.OptionButton optMAD 
      Caption         =   "Moving Average Demand"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   19
      Top             =   660
      Width           =   3795
   End
   Begin wizProgBar.Prg prgExcelGen 
      Height          =   330
      Left            =   90
      TabIndex        =   18
      Top             =   3150
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   582
      Picture         =   "PrintPartsRunDown.frx":0E42
      ForeColor       =   0
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "PrintPartsRunDown.frx":0E5E
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack 
      Height          =   3765
      Left            =   660
      TabIndex        =   17
      Top             =   6240
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   "Select month from the list"
      Top             =   4020
      Width           =   2445
   End
   Begin VB.OptionButton Option8 
      Caption         =   "Fill Rate Reports"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   5
      Top             =   1830
      Width           =   3795
   End
   Begin VB.OptionButton Option9 
      Caption         =   "Ordered Parts Report by Category"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   6
      Top             =   2130
      Width           =   3795
   End
   Begin VB.OptionButton Option10 
      Caption         =   "Parts Back-Order Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   7
      Top             =   2430
      Width           =   3795
   End
   Begin VB.OptionButton Option11 
      Caption         =   "Excel Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   8
      Top             =   2730
      Width           =   3795
   End
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2730
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4020
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -360
      TabIndex        =   14
      Top             =   3480
      Width           =   5355
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Inventory Adjustments"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   4
      Top             =   1530
      Width           =   3795
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Total Purchases Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   3
      Top             =   1230
      Width           =   3795
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Beginning Inventory Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   2
      Top             =   930
      Width           =   3795
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Total Cost of Sales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   1
      Top             =   360
      Width           =   3795
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Total Retail Sales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   0
      Top             =   60
      Value           =   -1  'True
      Width           =   3795
   End
   Begin SHDocVwCtl.WebBrowser browRank 
      Height          =   3945
      Left            =   6180
      TabIndex        =   13
      Top             =   6960
      Width           =   8685
      ExtentX         =   15319
      ExtentY         =   6959
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin Crystal.CrystalReport rptPrintRunDown 
      Left            =   3750
      Top             =   4530
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
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
      Height          =   765
      Left            =   2160
      MouseIcon       =   "PrintPartsRunDown.frx":0E7A
      MousePointer    =   99  'Custom
      Picture         =   "PrintPartsRunDown.frx":0FCC
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Close Window"
      Top             =   4560
      Width           =   735
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
      Height          =   765
      Left            =   1440
      MouseIcon       =   "PrintPartsRunDown.frx":1417
      MousePointer    =   99  'Custom
      Picture         =   "PrintPartsRunDown.frx":1569
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Print Report"
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   -390
      TabIndex        =   16
      Top             =   3690
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   3690
      Width           =   1335
   End
End
Attribute VB_Name = "frmPMISReports_PrintPartsRunDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSTKSTAT                                                         As ADODB.Recordset

Dim xlApp                                                             As Excel.Application
Dim xlBook                                                            As Excel.Workbook
Dim xlSheet                                                           As Excel.Worksheet
Dim MonCol                                                            As Long
Dim RETAIL_SALES_HARI_PARTS_GJ, RETAIL_SALES_HARI_PARTS_BP, RETAIL_SALES_HARI_PARTS_COUNTER, RETAIL_SALES_HARI_PARTS_JOBBER, RETAIL_SALES_HARI_PARTS_ACCESSORY, RETAIL_SALES_HARI_PARTS_WARRANTY, RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID, RETAIL_SALES_NON_HARI_PARTS_GJ, RETAIL_SALES_NON_HARI_PARTS_BP, RETAIL_SALES_NON_HARI_PARTS_COUNTER, RETAIL_SALES_NON_HARI_PARTS_JOBBER, RETAIL_SALES_NON_HARI_PARTS_ACCESSORY, RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ, RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP, RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER, RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER, RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY, RETAIL_SALES_OTHER_BRANDS, COST_OF_SALES_HARI_PARTS_GJ, COST_OF_SALES_HARI_PARTS_BP, COST_OF_SALES_HARI_PARTS_COUNTER, COST_OF_SALES_HARI_PARTS_JOBBER, COST_OF_SALES_HARI_PARTS_ACCESSORY, COST_OF_SALES_NON_HARI_PARTS_GJ, COST_OF_SALES_NON_HARI_PARTS_BP As Double
Attribute RETAIL_SALES_HARI_PARTS_BP.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_HARI_PARTS_COUNTER.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_HARI_PARTS_JOBBER.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_HARI_PARTS_ACCESSORY.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_HARI_PARTS_WARRANTY.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_NON_HARI_PARTS_GJ.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_NON_HARI_PARTS_BP.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_NON_HARI_PARTS_COUNTER.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_NON_HARI_PARTS_JOBBER.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_NON_HARI_PARTS_ACCESSORY.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY.VB_VarUserMemId = 1073938437
Attribute RETAIL_SALES_OTHER_BRANDS.VB_VarUserMemId = 1073938437
Attribute COST_OF_SALES_HARI_PARTS_GJ.VB_VarUserMemId = 1073938437
Attribute COST_OF_SALES_HARI_PARTS_BP.VB_VarUserMemId = 1073938437
Attribute COST_OF_SALES_HARI_PARTS_COUNTER.VB_VarUserMemId = 1073938437
Attribute COST_OF_SALES_HARI_PARTS_JOBBER.VB_VarUserMemId = 1073938437
Attribute COST_OF_SALES_HARI_PARTS_ACCESSORY.VB_VarUserMemId = 1073938437
Attribute COST_OF_SALES_NON_HARI_PARTS_GJ.VB_VarUserMemId = 1073938437
Attribute COST_OF_SALES_NON_HARI_PARTS_BP.VB_VarUserMemId = 1073938437
Dim COST_OF_SALES_NON_HARI_PARTS_COUNTER, COST_OF_SALES_NON_HARI_PARTS_JOBBER, COST_OF_SALES_NON_HARI_PARTS_ACCESSORY, COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ, COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP, COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER, COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER, COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY, COST_OF_SALES_OTHER_BRANDS, BI_HARI_GJ, BI_HARI_BP, BI_HARI_ACCESSORY, BI_NON_HARI_GJ, BI_NON_HARI_BP, BI_NON_HARI_ACCESSORY, BI_OTHER_BRANDS, PURCHASES_HARI_GJ, PURCHASES_HARI_BP, PURCHASES_HARI_ACCESSORY, PURCHASES_NON_HARI_GJ, PURCHASES_NON_HARI_BP, PURCHASES_NON_HARI_ACCESSORY, PURCHASES_OTHER_BRANDS, ADJUSTMENTS, EI_HARI_GJ, EI_HARI_BP, EI_HARI_ACCESSORY As Double
Attribute COST_OF_SALES_NON_HARI_PARTS_COUNTER.VB_VarUserMemId = 1073938465
Attribute COST_OF_SALES_NON_HARI_PARTS_JOBBER.VB_VarUserMemId = 1073938465
Attribute COST_OF_SALES_NON_HARI_PARTS_ACCESSORY.VB_VarUserMemId = 1073938465
Attribute COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ.VB_VarUserMemId = 1073938465
Attribute COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP.VB_VarUserMemId = 1073938465
Attribute COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER.VB_VarUserMemId = 1073938465
Attribute COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER.VB_VarUserMemId = 1073938465
Attribute COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY.VB_VarUserMemId = 1073938465
Attribute COST_OF_SALES_OTHER_BRANDS.VB_VarUserMemId = 1073938465
Attribute BI_HARI_GJ.VB_VarUserMemId = 1073938465
Attribute BI_HARI_BP.VB_VarUserMemId = 1073938465
Attribute BI_HARI_ACCESSORY.VB_VarUserMemId = 1073938465
Attribute BI_NON_HARI_GJ.VB_VarUserMemId = 1073938465
Attribute BI_NON_HARI_BP.VB_VarUserMemId = 1073938465
Attribute BI_NON_HARI_ACCESSORY.VB_VarUserMemId = 1073938465
Attribute BI_OTHER_BRANDS.VB_VarUserMemId = 1073938465
Attribute PURCHASES_HARI_GJ.VB_VarUserMemId = 1073938465
Attribute PURCHASES_HARI_BP.VB_VarUserMemId = 1073938465
Attribute PURCHASES_HARI_ACCESSORY.VB_VarUserMemId = 1073938465
Attribute PURCHASES_NON_HARI_GJ.VB_VarUserMemId = 1073938465
Attribute PURCHASES_NON_HARI_BP.VB_VarUserMemId = 1073938465
Attribute PURCHASES_NON_HARI_ACCESSORY.VB_VarUserMemId = 1073938465
Attribute PURCHASES_OTHER_BRANDS.VB_VarUserMemId = 1073938465
Attribute ADJUSTMENTS.VB_VarUserMemId = 1073938465
Attribute EI_HARI_GJ.VB_VarUserMemId = 1073938465
Attribute EI_HARI_BP.VB_VarUserMemId = 1073938465
Attribute EI_HARI_ACCESSORY.VB_VarUserMemId = 1073938465
Dim EI_NON_HARI_GJ, EI_NON_HARI_BP, EI_NON_HARI_ACCESSORY, EI_OTHER_BRANDS, WSC_NUMBER_ORDER_SLIP_RECEIVED, WSC_COMPLETELY_SERVE_ORDER_SLIP, WSC_NUMBER_LINE_ITEM_ORDERED, WSC_COMPLETELY_SERVE_LINE_ITEM, OTC_NUMBER_ORDER_SLIP_RECEIVED, OTC_COMPLETELY_SERVE_ORDER_SLIP, OTC_NUMBER_LINE_ITEM_ORDERED, OTC_COMPLETELY_SERVE_LINE_ITEM As Double
Attribute EI_NON_HARI_GJ.VB_VarUserMemId = 1073938492
Attribute EI_NON_HARI_BP.VB_VarUserMemId = 1073938492
Attribute EI_NON_HARI_ACCESSORY.VB_VarUserMemId = 1073938492
Attribute EI_OTHER_BRANDS.VB_VarUserMemId = 1073938492
Attribute WSC_NUMBER_ORDER_SLIP_RECEIVED.VB_VarUserMemId = 1073938492
Attribute WSC_COMPLETELY_SERVE_ORDER_SLIP.VB_VarUserMemId = 1073938492
Attribute WSC_NUMBER_LINE_ITEM_ORDERED.VB_VarUserMemId = 1073938492
Attribute WSC_COMPLETELY_SERVE_LINE_ITEM.VB_VarUserMemId = 1073938492
Attribute OTC_NUMBER_ORDER_SLIP_RECEIVED.VB_VarUserMemId = 1073938492
Attribute OTC_COMPLETELY_SERVE_ORDER_SLIP.VB_VarUserMemId = 1073938492
Attribute OTC_NUMBER_LINE_ITEM_ORDERED.VB_VarUserMemId = 1073938492
Attribute OTC_COMPLETELY_SERVE_LINE_ITEM.VB_VarUserMemId = 1073938492
Dim DD_TOTAL_LINES_ORDERED, DD_TOTAL_QTY_ORDERED, DD_TOTAL_LINES_SERVED, DD_TOTAL_QTY_SERVED, HARI_ORDERED_FILL, HARI_ORDERED_KILL, HARI_ORDERED_WARRANTY, BO_PER_ITEM_WARRANTY, BO_PER_ITEM_REGULAR As Double
Attribute DD_TOTAL_LINES_ORDERED.VB_VarUserMemId = 1073938504
Attribute DD_TOTAL_QTY_ORDERED.VB_VarUserMemId = 1073938504
Attribute DD_TOTAL_LINES_SERVED.VB_VarUserMemId = 1073938504
Attribute DD_TOTAL_QTY_SERVED.VB_VarUserMemId = 1073938504
Attribute HARI_ORDERED_FILL.VB_VarUserMemId = 1073938504
Attribute HARI_ORDERED_KILL.VB_VarUserMemId = 1073938504
Attribute HARI_ORDERED_WARRANTY.VB_VarUserMemId = 1073938504
Attribute BO_PER_ITEM_WARRANTY.VB_VarUserMemId = 1073938504
Attribute BO_PER_ITEM_REGULAR.VB_VarUserMemId = 1073938504

Sub ShowExcel()
    Screen.MousePointer = 11


    If Len(Dir(App.Path & "\PartsRundown.xlt")) <= 0 Then
        If EXTRACT_FILES(105, "PartsRundown.xlt") = False Then
            MsgBox "Please Put PartsRundown.xlt on " & vbCrLf & App.Path, vbInformation
            Exit Sub
        End If
    End If


    Set xlApp = CreateObject("Excel.Application")


    Set xlBook = xlApp.Workbooks.Open(App.Path & "\PartsRundown.xlt")
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Cells(2, 4) = "Year : " & cboYear
    Dim rsPRR_VIEW_MONTH                               As ADODB.Recordset
    Dim PRR_MONTH                                      As Integer
    Dim NEWMONTH, LASTMONTH                            As Integer

    Dim TOTAL_RETAIL_SALES, TOTAL_COST_OF_SALES, TOTAL_INVENTORY, MONTHLY_AVERAGE_DEMAND, SUM_DISCOUNT_HARI, SUM_DISCOUNT_NON_HARI As Double
    Dim vJanuary, vFebruary, vMarch, vApril, vMay, vJune, vJuly, vAugust, vSeptember, vOctober, vNovember, vDecember As Double
    Dim ADJUSTMENTS_GJ_HARI, ADJUSTMENTS_BP_HARI, ADJUSTMENTS_AC_HARI, ADJUSTMENTS_GJ_NON_HARI, ADJUSTMENTS_BP_NON_HARI, ADJUSTMENTS_AC_NON_HARI As Double

    vJanuary = 0
    vFebruary = 0
    vMarch = 0
    vApril = 0
    vMay = 0
    vJune = 0
    vJuly = 0
    vAugust = 0
    vSeptember = 0
    vOctober = 0
    vNovember = 0
    vDecember = 0

    prgExcelGen.Max = 72
    For PRR_MONTH = 1 To What_month(cboMonth)
        MonCol = PRR_MONTH + 3
        prgExcelGen.Value = 0
        RETAIL_SALES_HARI_PARTS_GJ = 0: RETAIL_SALES_HARI_PARTS_BP = 0: RETAIL_SALES_HARI_PARTS_COUNTER = 0: RETAIL_SALES_HARI_PARTS_JOBBER = 0: RETAIL_SALES_HARI_PARTS_ACCESSORY = 0
        RETAIL_SALES_HARI_PARTS_WARRANTY = 0: RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID = 0
        RETAIL_SALES_NON_HARI_PARTS_GJ = 0: RETAIL_SALES_NON_HARI_PARTS_BP = 0: RETAIL_SALES_NON_HARI_PARTS_COUNTER = 0: RETAIL_SALES_NON_HARI_PARTS_JOBBER = 0: RETAIL_SALES_NON_HARI_PARTS_ACCESSORY = 0
        RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ = 0: RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP = 0: RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER = 0: RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER = 0: RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = 0
        RETAIL_SALES_OTHER_BRANDS = 0: SUM_DISCOUNT_NON_HARI = 0: SUM_DISCOUNT_HARI = 0

        COST_OF_SALES_HARI_PARTS_GJ = 0: COST_OF_SALES_HARI_PARTS_BP = 0: COST_OF_SALES_HARI_PARTS_COUNTER = 0: COST_OF_SALES_HARI_PARTS_JOBBER = 0: COST_OF_SALES_HARI_PARTS_ACCESSORY = 0
        COST_OF_SALES_NON_HARI_PARTS_GJ = 0: COST_OF_SALES_NON_HARI_PARTS_BP = 0: COST_OF_SALES_NON_HARI_PARTS_COUNTER = 0: COST_OF_SALES_NON_HARI_PARTS_JOBBER = 0: COST_OF_SALES_NON_HARI_PARTS_ACCESSORY = 0
        COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ = 0: COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP = 0: COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER = 0: COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER = 0: COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = 0
        COST_OF_SALES_OTHER_BRANDS = 0

        BI_HARI_GJ = 0: BI_HARI_BP = 0: BI_HARI_ACCESSORY = 0: BI_NON_HARI_GJ = 0: BI_NON_HARI_BP = 0: BI_NON_HARI_ACCESSORY = 0
        BI_OTHER_BRANDS = 0
        PURCHASES_HARI_GJ = 0: PURCHASES_HARI_BP = 0: PURCHASES_HARI_ACCESSORY = 0: PURCHASES_NON_HARI_GJ = 0: PURCHASES_NON_HARI_BP = 0: PURCHASES_NON_HARI_ACCESSORY = 0
        PURCHASES_OTHER_BRANDS = 0

        ADJUSTMENTS = 0
        ADJUSTMENTS_GJ_HARI = 0: ADJUSTMENTS_BP_HARI = 0: ADJUSTMENTS_AC_HARI = 0: ADJUSTMENTS_GJ_NON_HARI = 0: ADJUSTMENTS_BP_NON_HARI = 0: ADJUSTMENTS_AC_NON_HARI = 0
        EI_HARI_GJ = 0: EI_HARI_BP = 0: EI_HARI_ACCESSORY = 0: EI_NON_HARI_GJ = 0: EI_NON_HARI_BP = 0: EI_NON_HARI_ACCESSORY = 0
        EI_OTHER_BRANDS = 0

        WSC_NUMBER_ORDER_SLIP_RECEIVED = 0: WSC_COMPLETELY_SERVE_ORDER_SLIP = 0: WSC_NUMBER_LINE_ITEM_ORDERED = 0: WSC_COMPLETELY_SERVE_LINE_ITEM = 0
        OTC_NUMBER_ORDER_SLIP_RECEIVED = 0: OTC_COMPLETELY_SERVE_ORDER_SLIP = 0: OTC_NUMBER_LINE_ITEM_ORDERED = 0: OTC_COMPLETELY_SERVE_LINE_ITEM = 0

        DD_TOTAL_LINES_ORDERED = 0
        DD_TOTAL_QTY_ORDERED = 0
        DD_TOTAL_LINES_SERVED = 0
        DD_TOTAL_QTY_SERVED = 0

        HARI_ORDERED_FILL = 0
        HARI_ORDERED_KILL = 0
        HARI_ORDERED_WARRANTY = 0

        BO_PER_ITEM_WARRANTY = 0
        BO_PER_ITEM_REGULAR = 0

        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("SELECT * FROM PMIS_vw_PRR_VIEW WHERE MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear.Text)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            rsPRR_VIEW_MONTH.MoveFirst
            Do While Not rsPRR_VIEW_MONTH.EOF
                If rsPRR_VIEW_MONTH!NON_HARI = "N" Then
                    If rsPRR_VIEW_MONTH!Type = "P" Then
                        '=========================================================================================================
                        'updating code: jaa - 07152008        - to trace all accessories with the type of "P" (HARI Accessories starts with "08")
                        If Left(Trim(rsPRR_VIEW_MONTH!STOCK_ORD), 2) <> "08" Then
                            'updating code: jaa - 10212008        - regardless of Sales Origin, if the trantype is CASH or CHARGE, Count as Counter
                            If (rsPRR_VIEW_MONTH!TranType = "CSH" Or rsPRR_VIEW_MONTH!TranType = "CHG") And rsPRR_VIEW_MONTH!SALES_ORIGIN <> "J" Then
                                If rsPRR_VIEW_MONTH!SALES_ORIGIN = "W" Or rsPRR_VIEW_MONTH!SALES_ORIGIN = "M" Or rsPRR_VIEW_MONTH!SALES_ORIGIN = "O" Then
                                    RETAIL_SALES_HARI_PARTS_COUNTER = RETAIL_SALES_HARI_PARTS_COUNTER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                    COST_OF_SALES_HARI_PARTS_COUNTER = COST_OF_SALES_HARI_PARTS_COUNTER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                                End If
                            ElseIf rsPRR_VIEW_MONTH!SALES_ORIGIN = "J" Then
                                RETAIL_SALES_HARI_PARTS_JOBBER = RETAIL_SALES_HARI_PARTS_JOBBER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                COST_OF_SALES_HARI_PARTS_JOBBER = COST_OF_SALES_HARI_PARTS_JOBBER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                            Else
                                'updating code:     jaa - 10242008
                                If rsPRR_VIEW_MONTH!SI_TYPE = "G" Then
                                    RETAIL_SALES_HARI_PARTS_GJ = RETAIL_SALES_HARI_PARTS_GJ + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                    COST_OF_SALES_HARI_PARTS_GJ = COST_OF_SALES_HARI_PARTS_GJ + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                                End If
                                If rsPRR_VIEW_MONTH!SI_TYPE = "B" Then
                                    RETAIL_SALES_HARI_PARTS_BP = RETAIL_SALES_HARI_PARTS_BP + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                    COST_OF_SALES_HARI_PARTS_BP = COST_OF_SALES_HARI_PARTS_BP + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                                End If

                            End If
                            If rsPRR_VIEW_MONTH!PAY_CLASS = "W" Then
                                RETAIL_SALES_HARI_PARTS_WARRANTY = RETAIL_SALES_HARI_PARTS_WARRANTY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                            End If
                            If rsPRR_VIEW_MONTH!PAY_CLASS = "C" Or rsPRR_VIEW_MONTH!PAY_CLASS = "I" Then
                                RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID = RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                            End If
                        Else
                            RETAIL_SALES_HARI_PARTS_ACCESSORY = RETAIL_SALES_HARI_PARTS_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                            COST_OF_SALES_HARI_PARTS_ACCESSORY = COST_OF_SALES_HARI_PARTS_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                            If rsPRR_VIEW_MONTH!PAY_CLASS = "W" Then
                                RETAIL_SALES_HARI_PARTS_WARRANTY = RETAIL_SALES_HARI_PARTS_WARRANTY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                            End If
                            If rsPRR_VIEW_MONTH!PAY_CLASS = "C" Or rsPRR_VIEW_MONTH!PAY_CLASS = "I" Then
                                RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID = RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                            End If
                        End If
                        '=========================================================================================================
                    Else
                        RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                        COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                    End If
                ElseIf rsPRR_VIEW_MONTH!NON_HARI = "O" Then
                    'updating code:     10212008    - Computer Other Brands for Parts Only
                    If rsPRR_VIEW_MONTH!Type = "P" Then
                        RETAIL_SALES_OTHER_BRANDS = RETAIL_SALES_OTHER_BRANDS + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                        COST_OF_SALES_OTHER_BRANDS = COST_OF_SALES_OTHER_BRANDS + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                    Else
                        RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                        COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                    End If
                Else
                    If rsPRR_VIEW_MONTH!Type = "P" Then
                        '=========================================================================================================
                        'updating code: jaa - 07152008        - to trace all accessories with the type of "P" and starts with "08"
                        If Left(Trim(rsPRR_VIEW_MONTH!STOCK_ORD), 2) <> "08" Then
                            'updating code: jaa - 10212008        - regardless of Sales Origin, if the trantype is CASH or CHARGE, Count as Counter
                            If (rsPRR_VIEW_MONTH!TranType = "CSH" Or rsPRR_VIEW_MONTH!TranType = "CHG") And rsPRR_VIEW_MONTH!SALES_ORIGIN <> "J" Then
                                If rsPRR_VIEW_MONTH!SALES_ORIGIN = "W" Or rsPRR_VIEW_MONTH!SALES_ORIGIN = "M" Or rsPRR_VIEW_MONTH!SALES_ORIGIN = "O" Then
                                    RETAIL_SALES_NON_HARI_PARTS_COUNTER = RETAIL_SALES_NON_HARI_PARTS_COUNTER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                    COST_OF_SALES_NON_HARI_PARTS_COUNTER = COST_OF_SALES_NON_HARI_PARTS_COUNTER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                                End If
                            ElseIf rsPRR_VIEW_MONTH!SALES_ORIGIN = "J" Then
                                RETAIL_SALES_NON_HARI_PARTS_JOBBER = RETAIL_SALES_NON_HARI_PARTS_JOBBER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                COST_OF_SALES_NON_HARI_PARTS_JOBBER = COST_OF_SALES_NON_HARI_PARTS_JOBBER + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                            Else
                                'updating code:     jaa - 10242008
                                If rsPRR_VIEW_MONTH!SI_TYPE = "G" Then
                                    RETAIL_SALES_NON_HARI_PARTS_GJ = RETAIL_SALES_NON_HARI_PARTS_GJ + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                    COST_OF_SALES_NON_HARI_PARTS_GJ = COST_OF_SALES_NON_HARI_PARTS_GJ + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                                End If
                                If rsPRR_VIEW_MONTH!SI_TYPE = "B" Then
                                    RETAIL_SALES_NON_HARI_PARTS_BP = RETAIL_SALES_NON_HARI_PARTS_BP + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                                    COST_OF_SALES_NON_HARI_PARTS_BP = COST_OF_SALES_NON_HARI_PARTS_BP + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                                End If
                            End If
                        Else
                            RETAIL_SALES_NON_HARI_PARTS_ACCESSORY = RETAIL_SALES_NON_HARI_PARTS_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                            COST_OF_SALES_NON_HARI_PARTS_ACCESSORY = COST_OF_SALES_NON_HARI_PARTS_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                        End If
                    Else
                        RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVAMT)
                        COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY + NumericVal(rsPRR_VIEW_MONTH!TOTALINVCOST)
                    End If
                    '=========================================================================================================
                End If
                rsPRR_VIEW_MONTH.MoveNext
            Loop
            'RETAIL_SALES_OTHER_BRANDS

            xlSheet.Cells(8, MonCol) = RETAIL_SALES_HARI_PARTS_GJ
            prgExcelGen.Value = 1: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(9, MonCol) = RETAIL_SALES_HARI_PARTS_BP
            prgExcelGen.Value = 2: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(10, MonCol) = RETAIL_SALES_HARI_PARTS_COUNTER
            prgExcelGen.Value = 3: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(11, MonCol) = RETAIL_SALES_HARI_PARTS_JOBBER
            prgExcelGen.Value = 4: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(12, MonCol) = RETAIL_SALES_HARI_PARTS_ACCESSORY
            prgExcelGen.Value = 5: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

            xlSheet.Cells(15, MonCol) = RETAIL_SALES_HARI_PARTS_WARRANTY
            prgExcelGen.Value = 6: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(16, MonCol) = RETAIL_SALES_HARI_PARTS_CUSTOMER_PAID
            prgExcelGen.Value = 7: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

            xlSheet.Cells(20, MonCol) = RETAIL_SALES_NON_HARI_PARTS_GJ
            prgExcelGen.Value = 8: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(21, MonCol) = RETAIL_SALES_NON_HARI_PARTS_BP
            prgExcelGen.Value = 9: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(22, MonCol) = RETAIL_SALES_NON_HARI_PARTS_COUNTER
            prgExcelGen.Value = 10: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(23, MonCol) = RETAIL_SALES_NON_HARI_PARTS_JOBBER
            prgExcelGen.Value = 11: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(24, MonCol) = RETAIL_SALES_NON_HARI_PARTS_ACCESSORY
            prgExcelGen.Value = 12: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

            xlSheet.Cells(26, MonCol) = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ
            prgExcelGen.Value = 13: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(27, MonCol) = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP
            prgExcelGen.Value = 14: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(28, MonCol) = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER
            prgExcelGen.Value = 15: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(29, MonCol) = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER
            prgExcelGen.Value = 16: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(30, MonCol) = RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY
            prgExcelGen.Value = 17: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

            xlSheet.Cells(35, MonCol) = RETAIL_SALES_OTHER_BRANDS
            prgExcelGen.Value = 18: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

            TOTAL_RETAIL_SALES = RETAIL_SALES_HARI_PARTS_GJ + RETAIL_SALES_HARI_PARTS_BP + _
                                 RETAIL_SALES_HARI_PARTS_COUNTER + RETAIL_SALES_HARI_PARTS_JOBBER + _
                                 RETAIL_SALES_HARI_PARTS_ACCESSORY + RETAIL_SALES_NON_HARI_PARTS_GJ + _
                                 RETAIL_SALES_NON_HARI_PARTS_BP + RETAIL_SALES_NON_HARI_PARTS_COUNTER + _
                                 RETAIL_SALES_NON_HARI_PARTS_JOBBER + RETAIL_SALES_NON_HARI_PARTS_ACCESSORY + _
                                 RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ + RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP + _
                                 RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER + RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER + _
                                 RETAIL_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY + RETAIL_SALES_OTHER_BRANDS
            xlSheet.Cells(41, MonCol) = COST_OF_SALES_HARI_PARTS_GJ
            prgExcelGen.Value = 19: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(42, MonCol) = COST_OF_SALES_HARI_PARTS_BP
            prgExcelGen.Value = 20: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(43, MonCol) = COST_OF_SALES_HARI_PARTS_COUNTER
            prgExcelGen.Value = 21: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(44, MonCol) = COST_OF_SALES_HARI_PARTS_JOBBER
            prgExcelGen.Value = 22: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(45, MonCol) = COST_OF_SALES_HARI_PARTS_ACCESSORY
            prgExcelGen.Value = 23: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

            xlSheet.Cells(49, MonCol) = COST_OF_SALES_NON_HARI_PARTS_GJ
            prgExcelGen.Value = 24: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(50, MonCol) = COST_OF_SALES_NON_HARI_PARTS_BP
            prgExcelGen.Value = 25: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(51, MonCol) = COST_OF_SALES_NON_HARI_PARTS_COUNTER
            prgExcelGen.Value = 26: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(52, MonCol) = COST_OF_SALES_NON_HARI_PARTS_JOBBER
            prgExcelGen.Value = 27: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(53, MonCol) = COST_OF_SALES_NON_HARI_PARTS_ACCESSORY
            prgExcelGen.Value = 28: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

            xlSheet.Cells(55, MonCol) = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ
            prgExcelGen.Value = 29: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(56, MonCol) = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP
            prgExcelGen.Value = 30: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(57, MonCol) = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER
            prgExcelGen.Value = 31: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(58, MonCol) = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER
            prgExcelGen.Value = 32: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(59, MonCol) = COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY
            prgExcelGen.Value = 33: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
            xlSheet.Cells(64, MonCol) = COST_OF_SALES_OTHER_BRANDS
            prgExcelGen.Value = 34: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        End If
        TOTAL_COST_OF_SALES = COST_OF_SALES_HARI_PARTS_GJ + COST_OF_SALES_HARI_PARTS_BP + COST_OF_SALES_HARI_PARTS_COUNTER + COST_OF_SALES_HARI_PARTS_JOBBER + COST_OF_SALES_HARI_PARTS_ACCESSORY + COST_OF_SALES_NON_HARI_PARTS_GJ + COST_OF_SALES_NON_HARI_PARTS_BP + COST_OF_SALES_NON_HARI_PARTS_COUNTER + COST_OF_SALES_NON_HARI_PARTS_JOBBER + COST_OF_SALES_NON_HARI_PARTS_ACCESSORY + COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_GJ + COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_BP + COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_COUNTER + COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_JOBBER + COST_OF_SALES_NON_HARI_PARTS_CARRIED_BY_HARI_ACCESSORY + COST_OF_SALES_OTHER_BRANDS + RETAIL_SALES_HARI_PARTS_GJ

        Set rsPRR_VIEW_MONTH = Nothing
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset

        If PRR_MONTH = 1 Then
            Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TYPE as TYP,TOTAL_COST,STOCKTYPE,MONTH(DATE_GEN) AS MONTH_DATE, YEAR(DATE_GEN) AS YEAR_DATE from PMIS_vw_PRR_BEG_INVENTORY_HARI where MONTH(DATE_GEN) = 12 AND YEAR(DATE_GEN) = " & NumericVal(cboYear) - 1 & " ORDER BY DATE_GEN ASC")
        Else
            Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TYPE as TYP,TOTAL_COST,STOCKTYPE,MONTH(DATE_GEN) AS MONTH_DATE, YEAR(DATE_GEN) AS YEAR_DATE from PMIS_vw_PRR_BEG_INVENTORY_HARI where MONTH(DATE_GEN) + 1 = " & NumericVal(PRR_MONTH) & " AND YEAR(DATE_GEN) = " & cboYear & " ORDER BY DATE_GEN ASC")
        End If
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            rsPRR_VIEW_MONTH.MoveFirst
            Do While Not rsPRR_VIEW_MONTH.EOF
                If rsPRR_VIEW_MONTH!TYP <> "M" Then
                    If Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "BP" Then
                        BI_HARI_BP = BI_HARI_BP + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    ElseIf Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "AC" Then
                        BI_HARI_ACCESSORY = BI_HARI_ACCESSORY + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    Else
                        BI_HARI_GJ = BI_HARI_GJ + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    End If
                End If
                rsPRR_VIEW_MONTH.MoveNext
            Loop
        End If


        Set rsPRR_VIEW_MONTH = Nothing
        xlSheet.Cells(70, MonCol) = BI_HARI_GJ
        prgExcelGen.Value = 35: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(71, MonCol) = BI_HARI_BP
        prgExcelGen.Value = 36: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(72, MonCol) = BI_HARI_ACCESSORY
        prgExcelGen.Value = 37: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        If PRR_MONTH = 1 Then
            Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TYPE as TYP,NON_HARI,TOTAL_COST,STOCKTYPE,MONTH(DATE_GEN) AS MONTH_DATE, YEAR(DATE_GEN) AS YEAR_DATE from PMIS_vw_PRR_BEG_INVENTORY_NON_HARI where Month(DATE_GEN) = 12 AND YEAR(DATE_GEN) = " & NumericVal(cboYear) - 1 & " ORDER BY DATE_GEN ASC")
        Else
            Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TYPE as TYP,NON_HARI,TOTAL_COST,STOCKTYPE,MONTH(DATE_GEN) AS MONTH_DATE, YEAR(DATE_GEN) AS YEAR_DATE from PMIS_vw_PRR_BEG_INVENTORY_NON_HARI where month(DATE_GEN) + 1 = " & NumericVal(PRR_MONTH) & " AND YEAR(DATE_GEN) = " & cboYear & " ORDER BY DATE_GEN ASC")
        End If
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            rsPRR_VIEW_MONTH.MoveFirst
            Do While Not rsPRR_VIEW_MONTH.EOF
                If rsPRR_VIEW_MONTH!TYP <> "M" Then
                    If Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "BP" Then
                        BI_NON_HARI_BP = BI_NON_HARI_BP + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    ElseIf Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "AC" Then
                        BI_NON_HARI_ACCESSORY = BI_NON_HARI_ACCESSORY + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    Else
                        BI_NON_HARI_GJ = BI_NON_HARI_GJ + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    End If
                End If
                rsPRR_VIEW_MONTH.MoveNext
            Loop
        End If
        '******************************************************************************************
        'UPDATING CODE:     JAA - 08062008          - TO UPDATE OTHER BRANDS IN THE COMPUTATIONS
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        If PRR_MONTH = 1 Then
            Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TYPE as TYP,NON_HARI,TOTAL_COST,MONTH(DATE_GEN) AS MONTH_DATE, YEAR(DATE_GEN) AS YEAR_DATE from PMIS_Prr_Beg_Inventory where TYPE = 'P' AND NON_HARI = 'O' AND Month(DATE_GEN) = 12 AND YEAR(DATE_GEN) = " & NumericVal(cboYear) - 1 & " ORDER BY DATE_GEN ASC")
        Else
            Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TYPE as TYP,NON_HARI,TOTAL_COST,MONTH(DATE_GEN) AS MONTH_DATE, YEAR(DATE_GEN) AS YEAR_DATE from PMIS_Prr_Beg_Inventory where TYPE = 'P' AND NON_HARI = 'O' AND month(DATE_GEN) + 1 = " & NumericVal(PRR_MONTH) & " AND YEAR(DATE_GEN) = " & cboYear & " ORDER BY DATE_GEN ASC")
        End If
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            rsPRR_VIEW_MONTH.MoveFirst
            Do While Not rsPRR_VIEW_MONTH.EOF
                BI_OTHER_BRANDS = BI_OTHER_BRANDS + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                rsPRR_VIEW_MONTH.MoveNext
            Loop
        End If
        '******************************************************************************************

        xlSheet.Cells(74, MonCol) = BI_NON_HARI_GJ
        prgExcelGen.Value = 38: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(75, MonCol) = BI_NON_HARI_BP
        prgExcelGen.Value = 39: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(76, MonCol) = BI_NON_HARI_ACCESSORY
        prgExcelGen.Value = 40: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(78, MonCol) = BI_OTHER_BRANDS
        prgExcelGen.Value = 41

        Set rsPRR_VIEW_MONTH = Nothing
        NEWMONTH = 0
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,LTRIM(RTRIM(STOCKTYPE)) AS STOCKTYPE,MONTH_TRAN AS MONTH_DATE, YEAR_TRAN AS YEAR_DATE from PMIS_vw_PRR_PURCHASES_HARI where MONTH_TRAN = " & PRR_MONTH & " AND YEAR_TRAN=" & cboYear & " ORDER BY MONTH_TRAN ASC")
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            rsPRR_VIEW_MONTH.MoveFirst
            Do While Not rsPRR_VIEW_MONTH.EOF
                LASTMONTH = rsPRR_VIEW_MONTH!MONTH_DATE
                If NEWMONTH = 0 Or LASTMONTH <> NEWMONTH Then
                    NEWMONTH = LASTMONTH
                    PURCHASES_HARI_GJ = NumericVal(0)
                    PURCHASES_HARI_BP = NumericVal(0)
                    PURCHASES_HARI_ACCESSORY = NumericVal(0)
                End If
                If Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "BP" Then
                    PURCHASES_HARI_BP = PURCHASES_HARI_BP + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                ElseIf Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "AC" Then
                    PURCHASES_HARI_ACCESSORY = PURCHASES_HARI_ACCESSORY + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                Else
                    PURCHASES_HARI_GJ = PURCHASES_HARI_GJ + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                End If
                rsPRR_VIEW_MONTH.MoveNext
            Loop
        End If
        xlSheet.Cells(83, MonCol) = PURCHASES_HARI_GJ
        prgExcelGen.Value = 42: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(84, MonCol) = PURCHASES_HARI_BP
        prgExcelGen.Value = 43: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(85, MonCol) = PURCHASES_HARI_ACCESSORY
        prgExcelGen.Value = 44: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

        Set rsPRR_VIEW_MONTH = Nothing
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,LTRIM(RTRIM(STOCKTYPE)) AS STOCKTYPE,MONTH_TRAN AS MONTH_DATE, YEAR_TRAN AS YEAR_DATE from PMIS_vw_PRR_PURCHASES_NON_HARI where MONTH_TRAN = " & PRR_MONTH & " AND YEAR_TRAN = " & cboYear & " ORDER BY MONTH_TRAN ASC")
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            rsPRR_VIEW_MONTH.MoveFirst
            Do While Not rsPRR_VIEW_MONTH.EOF
                LASTMONTH = Month(rsPRR_VIEW_MONTH!MONTH_DATE)
                If NEWMONTH = 0 Or LASTMONTH <> NEWMONTH Then
                    NEWMONTH = LASTMONTH
                    PURCHASES_NON_HARI_GJ = NumericVal(0)
                    PURCHASES_NON_HARI_BP = NumericVal(0)
                    PURCHASES_NON_HARI_ACCESSORY = NumericVal(0)
                End If

                If Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "BP" Then
                    PURCHASES_NON_HARI_BP = PURCHASES_NON_HARI_BP + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                ElseIf Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "AC" Then
                    PURCHASES_NON_HARI_ACCESSORY = PURCHASES_NON_HARI_ACCESSORY + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                Else
                    PURCHASES_NON_HARI_GJ = PURCHASES_NON_HARI_GJ + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                End If
                rsPRR_VIEW_MONTH.MoveNext
            Loop
        End If
        xlSheet.Cells(87, MonCol) = PURCHASES_NON_HARI_GJ
        prgExcelGen.Value = 45: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(88, MonCol) = PURCHASES_NON_HARI_BP
        prgExcelGen.Value = 46: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(89, MonCol) = PURCHASES_NON_HARI_ACCESSORY
        prgExcelGen.Value = 47: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        Set rsPRR_VIEW_MONTH = Nothing

        'PURCHASES_OTHER_BRANDS = NumericVal(0)
        '**********
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,LTRIM(RTRIM(STOCKTYPE)) AS STOCKTYPE,MONTH_TRAN AS MONTH_DATE, YEAR_TRAN AS YEAR_DATE from PMIS_vw_PRR_PURCHASES where NON_HARI = 'O' AND MONTH_TRAN = " & PRR_MONTH & " AND YEAR_TRAN = " & cboYear & " ORDER BY MONTH_TRAN ASC")
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            rsPRR_VIEW_MONTH.MoveFirst
            Do While Not rsPRR_VIEW_MONTH.EOF
                PURCHASES_OTHER_BRANDS = PURCHASES_OTHER_BRANDS + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                rsPRR_VIEW_MONTH.MoveNext
            Loop
        End If
        xlSheet.Cells(91, MonCol) = PURCHASES_OTHER_BRANDS
        '**********
        'ADJUSTMENTS
        '**********

        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,LTRIM(RTRIM(STOCKTYPE)) AS STOCKTYPE,NON_HARI from PMIS_Prr_Adjustments where NON_HARI <> 'O' AND TRANNO = '111111' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            rsPRR_VIEW_MONTH.MoveFirst
            Do While Not rsPRR_VIEW_MONTH.EOF
                If Trim(Null2String(rsPRR_VIEW_MONTH!NON_HARI)) = "N" Then
                    If Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "BP" Then
                        ADJUSTMENTS_BP_HARI = ADJUSTMENTS_BP_HARI + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    ElseIf Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "AC" Then
                        ADJUSTMENTS_AC_HARI = ADJUSTMENTS_AC_HARI + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    Else
                        ADJUSTMENTS_GJ_HARI = ADJUSTMENTS_GJ_HARI + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    End If
                Else
                    If Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "BP" Then
                        ADJUSTMENTS_BP_NON_HARI = ADJUSTMENTS_BP_NON_HARI + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    ElseIf Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "AC" Then
                        ADJUSTMENTS_AC_NON_HARI = ADJUSTMENTS_AC_NON_HARI + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    Else
                        ADJUSTMENTS_GJ_NON_HARI = ADJUSTMENTS_GJ_NON_HARI + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    End If
                End If
                rsPRR_VIEW_MONTH.MoveNext
            Loop
        End If

        '***********************************************************************************************
        'updating code:     jaa - 09242008      - Adjustments is already been computed after this code
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,LTRIM(RTRIM(STOCKTYPE)) AS STOCKTYPE,NON_HARI from PMIS_Prr_Adjustments where NON_HARI <> 'O' AND TRANNO = '000000' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            rsPRR_VIEW_MONTH.MoveFirst
            Do While Not rsPRR_VIEW_MONTH.EOF
                If Trim(Null2String(rsPRR_VIEW_MONTH!NON_HARI)) = "N" Then
                    If Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "BP" Then
                        ADJUSTMENTS_BP_HARI = ADJUSTMENTS_BP_HARI - NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    ElseIf Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "AC" Then
                        ADJUSTMENTS_AC_HARI = ADJUSTMENTS_AC_HARI - NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    Else
                        ADJUSTMENTS_GJ_HARI = ADJUSTMENTS_GJ_HARI - NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    End If
                Else
                    If Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "BP" Then
                        ADJUSTMENTS_BP_NON_HARI = ADJUSTMENTS_BP_NON_HARI - NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    ElseIf Trim(Null2String(rsPRR_VIEW_MONTH!StockType)) = "AC" Then
                        ADJUSTMENTS_AC_NON_HARI = ADJUSTMENTS_AC_NON_HARI - NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    Else
                        ADJUSTMENTS_GJ_NON_HARI = ADJUSTMENTS_GJ_NON_HARI - NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                    End If
                End If
                rsPRR_VIEW_MONTH.MoveNext
            Loop
        End If

        '***********************************************************************************************

        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,LTRIM(RTRIM(STOCKTYPE)) AS STOCKTYPE,TRANNO from PMIS_Prr_Adjustments where NON_HARI = 'O' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        'Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select TOTAL_COST,LTRIM(RTRIM(STOCKTYPE)) AS STOCKTYPE,TRANNO from PMIS_Prr_Adjustments where MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            rsPRR_VIEW_MONTH.MoveFirst
            Do While Not rsPRR_VIEW_MONTH.EOF
                If rsPRR_VIEW_MONTH!TRANNO = "000000" Then
                    ADJUSTMENTS = ADJUSTMENTS - NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                Else
                    ADJUSTMENTS = ADJUSTMENTS + NumericVal(N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_COST))
                End If
                rsPRR_VIEW_MONTH.MoveNext
            Loop
        End If


        'EI_OTHER_BRANDS = BI_OTHER_BRANDS + PURCHASES_OTHER_BRANDS + ADJUSTMENTS - COST_OF_SALES_OTHER_BRANDS
        'WHERE TO DEDUCT ADJUSTMENTS? WHAT IS THE BRAND OF ADJUSTMENTS?
        EI_OTHER_BRANDS = BI_OTHER_BRANDS + PURCHASES_OTHER_BRANDS + ADJUSTMENTS - COST_OF_SALES_OTHER_BRANDS
        ADJUSTMENTS = ADJUSTMENTS + ADJUSTMENTS_GJ_HARI + ADJUSTMENTS_BP_HARI + ADJUSTMENTS_AC_HARI + ADJUSTMENTS_GJ_NON_HARI + ADJUSTMENTS_BP_NON_HARI + ADJUSTMENTS_AC_NON_HARI

        xlSheet.Cells(94, MonCol) = ADJUSTMENTS

        EI_HARI_GJ = BI_HARI_GJ + PURCHASES_HARI_GJ + ADJUSTMENTS_GJ_HARI - _
                     (COST_OF_SALES_HARI_PARTS_GJ + COST_OF_SALES_HARI_PARTS_JOBBER + _
                      COST_OF_SALES_HARI_PARTS_COUNTER)

        EI_HARI_BP = BI_HARI_BP + PURCHASES_HARI_BP + ADJUSTMENTS_BP_HARI - COST_OF_SALES_HARI_PARTS_BP
        EI_HARI_ACCESSORY = BI_HARI_ACCESSORY + PURCHASES_HARI_ACCESSORY + ADJUSTMENTS_AC_HARI - COST_OF_SALES_HARI_PARTS_ACCESSORY

        EI_NON_HARI_GJ = BI_NON_HARI_GJ + PURCHASES_NON_HARI_GJ + ADJUSTMENTS_GJ_NON_HARI - (COST_OF_SALES_NON_HARI_PARTS_GJ + COST_OF_SALES_NON_HARI_PARTS_JOBBER + COST_OF_SALES_NON_HARI_PARTS_COUNTER)
        EI_NON_HARI_BP = BI_NON_HARI_BP + PURCHASES_NON_HARI_BP + ADJUSTMENTS_BP_NON_HARI - COST_OF_SALES_NON_HARI_PARTS_BP
        EI_NON_HARI_ACCESSORY = BI_NON_HARI_ACCESSORY + PURCHASES_NON_HARI_ACCESSORY + ADJUSTMENTS_AC_NON_HARI - COST_OF_SALES_NON_HARI_PARTS_ACCESSORY

        xlSheet.Cells(98, MonCol) = EI_HARI_GJ
        prgExcelGen.Value = 48: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(99, MonCol) = EI_HARI_BP
        prgExcelGen.Value = 49: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(100, MonCol) = EI_HARI_ACCESSORY
        prgExcelGen.Value = 50: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

        xlSheet.Cells(102, MonCol) = EI_NON_HARI_GJ
        prgExcelGen.Value = 51: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(103, MonCol) = EI_NON_HARI_BP
        prgExcelGen.Value = 52: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(104, MonCol) = EI_NON_HARI_ACCESSORY
        prgExcelGen.Value = 53: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(106, MonCol) = EI_OTHER_BRANDS

        TOTAL_INVENTORY = EI_HARI_GJ + EI_HARI_BP + EI_HARI_ACCESSORY + EI_NON_HARI_GJ + EI_NON_HARI_BP + EI_NON_HARI_ACCESSORY + EI_OTHER_BRANDS


        '        If PRR_MONTH = 1 Then vJanuary = TOTAL_INVENTORY
        '        If PRR_MONTH = 2 Then vFebruary = TOTAL_INVENTORY
        '        If PRR_MONTH = 3 Then vMarch = TOTAL_INVENTORY
        '        If PRR_MONTH = 4 Then vApril = TOTAL_INVENTORY
        '        If PRR_MONTH = 5 Then vMay = TOTAL_INVENTORY
        '        If PRR_MONTH = 6 Then vJune = TOTAL_INVENTORY
        '        If PRR_MONTH = 7 Then vJuly = TOTAL_INVENTORY
        '        If PRR_MONTH = 8 Then vAugust = TOTAL_INVENTORY
        '        If PRR_MONTH = 9 Then vSeptember = TOTAL_INVENTORY
        '        If PRR_MONTH = 10 Then vOctober = TOTAL_INVENTORY
        '        If PRR_MONTH = 11 Then vNovember = TOTAL_INVENTORY
        '        If PRR_MONTH = 12 Then vDecember = TOTAL_INVENTORY


        If PRR_MONTH = 1 Then vJanuary = TOTAL_RETAIL_SALES
        If PRR_MONTH = 2 Then vFebruary = TOTAL_RETAIL_SALES
        If PRR_MONTH = 3 Then vMarch = TOTAL_RETAIL_SALES
        If PRR_MONTH = 4 Then vApril = TOTAL_RETAIL_SALES
        If PRR_MONTH = 5 Then vMay = TOTAL_RETAIL_SALES
        If PRR_MONTH = 6 Then vJune = TOTAL_RETAIL_SALES
        If PRR_MONTH = 7 Then vJuly = TOTAL_RETAIL_SALES
        If PRR_MONTH = 8 Then vAugust = TOTAL_RETAIL_SALES
        If PRR_MONTH = 9 Then vSeptember = TOTAL_RETAIL_SALES
        If PRR_MONTH = 10 Then vOctober = TOTAL_RETAIL_SALES
        If PRR_MONTH = 11 Then vNovember = TOTAL_RETAIL_SALES
        If PRR_MONTH = 12 Then vDecember = TOTAL_RETAIL_SALES



        If PRR_MONTH = 1 Then MONTHLY_AVERAGE_DEMAND = (vJanuary + vDecember + vNovember) / 3
        If PRR_MONTH = 2 Then MONTHLY_AVERAGE_DEMAND = (vFebruary + vJanuary + vDecember) / 3
        If PRR_MONTH = 3 Then MONTHLY_AVERAGE_DEMAND = (vMarch + vFebruary + vJanuary) / 3
        If PRR_MONTH = 4 Then MONTHLY_AVERAGE_DEMAND = (vApril + vMarch + vFebruary) / 3
        If PRR_MONTH = 5 Then MONTHLY_AVERAGE_DEMAND = (vMay + vApril + vMarch) / 3
        If PRR_MONTH = 6 Then MONTHLY_AVERAGE_DEMAND = (vJune + vMay + vApril) / 3
        If PRR_MONTH = 7 Then MONTHLY_AVERAGE_DEMAND = (vJuly + vJune + vMay) / 3
        If PRR_MONTH = 8 Then MONTHLY_AVERAGE_DEMAND = (vAugust + vJuly + vJune) / 3
        If PRR_MONTH = 9 Then MONTHLY_AVERAGE_DEMAND = (vSeptember + vAugust + vJuly) / 3
        If PRR_MONTH = 10 Then MONTHLY_AVERAGE_DEMAND = (vOctober + vSeptember + vAugust) / 3
        If PRR_MONTH = 11 Then MONTHLY_AVERAGE_DEMAND = (vNovember + vOctober + vSeptember) / 3
        If PRR_MONTH = 12 Then MONTHLY_AVERAGE_DEMAND = (vDecember + vNovember + vOctober) / 3



        xlSheet.Cells(109, MonCol) = MONTHLY_AVERAGE_DEMAND
        prgExcelGen.Value = 53: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        If MONTHLY_AVERAGE_DEMAND > 0 Then
            xlSheet.Cells(111, MonCol) = TOTAL_INVENTORY / MONTHLY_AVERAGE_DEMAND
            prgExcelGen.Value = 54: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        Else
            xlSheet.Cells(111, MonCol) = 0
            prgExcelGen.Value = 54: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        End If
        If TOTAL_COST_OF_SALES > 0 Then
            xlSheet.Cells(113, MonCol) = TOTAL_INVENTORY / TOTAL_COST_OF_SALES
            prgExcelGen.Value = 55: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        Else
            xlSheet.Cells(113, MonCol) = 0
            prgExcelGen.Value = 55: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        End If

        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        'updating code:     jaa    - to include jobber and sales/marketing in counting for Workshop Customer Fill Rate
        'Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select COUNT(REFPISNO) as TOTAL_ORDERED from PMIS_vw_PARTS_PRS_TRANS where REFPISNO IS NOT NULL AND SALES_ORIGIN = 'S' AND STATUS = 'P' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select COUNT(REFPISNO) as TOTAL_ORDERED from PMIS_vw_PARTS_PRS_TRANS where REFPISNO IS NOT NULL AND SALES_ORIGIN <> 'W' AND STATUS = 'P' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            WSC_NUMBER_ORDER_SLIP_RECEIVED = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_ORDERED)
        End If
        'updating code:     jaa    - to include jobber and sales/marketing in counting for Workshop Customer Fill Rate and include BILLED transaction
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        'Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select COUNT(REFPRSNO) as TOTAL_SERVED from PMIS_vw_PARTS_PRS_ISSUANCE where REFPRSNO IS NOT NULL AND SALES_ORIGIN = 'S' AND STATUS = 'P' AND TRANTYPE = 'RIV' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select COUNT(REFPRSNO) as TOTAL_SERVED from PMIS_vw_PARTS_PRS_ISSUANCE where REFPRSNO IS NOT NULL AND SALES_ORIGIN <> 'W' AND (STATUS = 'P' OR STATUS = 'B') AND TRANTYPE = 'RIV' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            WSC_COMPLETELY_SERVE_ORDER_SLIP = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_SERVED)
        End If
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select DISTINCT COUNT(PARTNO) as TOTAL_ORDERED from PMIS_vw_Prr_Demand where S_REQUESTED > 0 AND MONTH(DATE_GEN) = " & PRR_MONTH & " AND YEAR(DATE_GEN) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            WSC_NUMBER_LINE_ITEM_ORDERED = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_ORDERED)
        End If
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select DISTINCT COUNT(PARTNO) as TOTAL_SERVED from PMIS_vw_Prr_Demand where S_SERVED > 0 AND MONTH(DATE_GEN) = " & PRR_MONTH & " AND YEAR(DATE_GEN) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            WSC_COMPLETELY_SERVE_LINE_ITEM = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_SERVED)
        End If

        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        'updating code:     jaa    - Count all Over the Counter only
        'Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select COUNT(REFPISNO) as TOTAL_ORDERED from PMIS_vw_PARTS_PRS_TRANS where REFPISNO IS NOT NULL AND SALES_ORIGIN <> 'S' AND STATUS = 'P' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select COUNT(REFPISNO) as TOTAL_ORDERED from PMIS_vw_PARTS_PRS_TRANS where REFPISNO IS NOT NULL AND SALES_ORIGIN = 'W' AND STATUS = 'P' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            OTC_NUMBER_ORDER_SLIP_RECEIVED = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_ORDERED)
        End If
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        'updating code:     jaa    - Count all Over the Counter only
        'Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select COUNT(REFPRSNO) as TOTAL_SERVED from PMIS_vw_PARTS_PRS_ISSUANCE where REFPRSNO IS NOT NULL AND SALES_ORIGIN <> 'S' AND STATUS = 'P' AND (TRANTYPE = 'CSH' OR TRANTYPE = 'CHG') AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select COUNT(REFPRSNO) as TOTAL_SERVED from PMIS_vw_PARTS_PRS_ISSUANCE where REFPRSNO IS NOT NULL AND SALES_ORIGIN = 'W' AND STATUS = 'P' AND (TRANTYPE = 'CSH' OR TRANTYPE = 'CHG') AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            OTC_COMPLETELY_SERVE_ORDER_SLIP = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_SERVED)
        End If
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select DISTINCT COUNT(PARTNO) as TOTAL_ORDERED from PMIS_vw_Prr_Demand where C_REQUESTED > 0 AND MONTH(DATE_GEN) = " & PRR_MONTH & " AND YEAR(DATE_GEN) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            OTC_NUMBER_LINE_ITEM_ORDERED = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_ORDERED)
        End If
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select DISTINCT COUNT(PARTNO) as TOTAL_SERVED from PMIS_vw_Prr_Demand where C_SERVED > 0 AND MONTH(DATE_GEN) = " & PRR_MONTH & " AND YEAR(DATE_GEN) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            OTC_COMPLETELY_SERVE_LINE_ITEM = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_SERVED)
        End If

        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select DISTINCT COUNT(PARTNO) as TOTAL_ORDERED from PMIS_vw_Prr_Demand where D_ORDERED > 0 AND MONTH(DATE_GEN) = " & PRR_MONTH & " AND YEAR(DATE_GEN) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            DD_TOTAL_LINES_ORDERED = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_ORDERED)
        End If
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select SUM(D_ORDERED) as TOTAL_ORDERED from PMIS_vw_Prr_Demand where D_ORDERED > 0 AND MONTH(DATE_GEN) = " & PRR_MONTH & " AND YEAR(DATE_GEN) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            DD_TOTAL_QTY_ORDERED = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_ORDERED)
        End If
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select DISTINCT COUNT(PARTNO) as TOTAL_SERVED from PMIS_vw_Prr_Demand where D_SERVED > 0 AND MONTH(DATE_GEN) = " & PRR_MONTH & " AND YEAR(DATE_GEN) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            DD_TOTAL_LINES_SERVED = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_SERVED)
        End If
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select SUM(D_SERVED) as TOTAL_SERVED from PMIS_vw_Prr_Demand where D_SERVED > 0 AND MONTH(DATE_GEN) = " & PRR_MONTH & " AND YEAR(DATE_GEN) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            DD_TOTAL_QTY_SERVED = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_SERVED)
        End If

        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select DISTINCT COUNT(PART_ORD) as TOTAL_ORDERED from PMIS_vw_PO_Details where ORDERTYPE <> 'W' AND POFILL = 1 AND STATUS = 'P' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            HARI_ORDERED_FILL = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_ORDERED)
        End If

        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select DISTINCT COUNT(PART_ORD) as TOTAL_ORDERED from PMIS_vw_PO_Details where ORDERTYPE <> 'W' AND POFILL = 0 AND STATUS = 'P' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            HARI_ORDERED_KILL = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_ORDERED)
        End If

        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select DISTINCT COUNT(PART_ORD) as TOTAL_ORDERED from PMIS_vw_PO_Details where ORDERTYPE = 'W' AND STATUS = 'P' AND MONTH(TRANDATE) = " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            HARI_ORDERED_WARRANTY = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_ORDERED)
        End If

        'FOR BACK-ORDER WARRANTY
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select DISTINCT COUNT(PART_ORD) as TOTAL_BO from PMIS_vw_PO_Details where QTY_BACKORDER > 0 AND ORDERTYPE = 'W' AND STATUS = 'P' AND MONTH(TRANDATE) >= " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            BO_PER_ITEM_WARRANTY = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_BO)
        End If

        'FOR BACK-ORDER WARRANTY
        Set rsPRR_VIEW_MONTH = New ADODB.Recordset
        Set rsPRR_VIEW_MONTH = gconDMIS.Execute("Select DISTINCT COUNT(PART_ORD) as TOTAL_BO from PMIS_vw_PO_Details where QTY_BACKORDER > 0 AND ORDERTYPE = 'R' AND STATUS = 'P' AND MONTH(TRANDATE) >= " & PRR_MONTH & " AND YEAR(TRANDATE) = " & cboYear)
        If Not rsPRR_VIEW_MONTH.EOF And Not rsPRR_VIEW_MONTH.BOF Then
            BO_PER_ITEM_REGULAR = N2Str2Zero(rsPRR_VIEW_MONTH!TOTAL_BO)
        End If

        xlSheet.Cells(117, MonCol) = WSC_NUMBER_ORDER_SLIP_RECEIVED
        prgExcelGen.Value = 56: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(118, MonCol) = WSC_COMPLETELY_SERVE_ORDER_SLIP
        prgExcelGen.Value = 57: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(119, MonCol) = WSC_NUMBER_LINE_ITEM_ORDERED
        prgExcelGen.Value = 58: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(120, MonCol) = WSC_COMPLETELY_SERVE_LINE_ITEM
        prgExcelGen.Value = 59: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

        xlSheet.Cells(124, MonCol) = OTC_NUMBER_ORDER_SLIP_RECEIVED
        prgExcelGen.Value = 60: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(125, MonCol) = OTC_COMPLETELY_SERVE_ORDER_SLIP
        prgExcelGen.Value = 61: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(126, MonCol) = OTC_NUMBER_LINE_ITEM_ORDERED
        prgExcelGen.Value = 62: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(127, MonCol) = OTC_COMPLETELY_SERVE_LINE_ITEM
        prgExcelGen.Value = 63: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

        xlSheet.Cells(132, MonCol) = DD_TOTAL_LINES_ORDERED
        prgExcelGen.Value = 64: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(133, MonCol) = DD_TOTAL_QTY_ORDERED
        prgExcelGen.Value = 65: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(134, MonCol) = DD_TOTAL_LINES_SERVED
        prgExcelGen.Value = 66: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(135, MonCol) = DD_TOTAL_QTY_SERVED
        prgExcelGen.Value = 67: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

        xlSheet.Cells(139, MonCol) = HARI_ORDERED_FILL
        prgExcelGen.Value = 68: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(140, MonCol) = HARI_ORDERED_KILL
        prgExcelGen.Value = 69: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(141, MonCol) = HARI_ORDERED_WARRANTY
        prgExcelGen.Value = 70: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

        xlSheet.Cells(144, MonCol) = BO_PER_ITEM_WARRANTY
        prgExcelGen.Value = 71: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents
        xlSheet.Cells(145, MonCol) = BO_PER_ITEM_REGULAR
        prgExcelGen.Value = 72: prgExcelGen.Text = "Generating PRR for " & Cap1st(Left(The_month(PRR_MONTH), 3)) & " " & cboYear.Text & " (" & Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %)": DoEvents

    Next
    prgExcelGen.Text = "PRR Generation (100% Completed)"
    xlApp.Visible = True
    Set xlApp = Nothing
    Screen.MousePointer = 0

CloseExcel:
    Set xlApp = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    Screen.MousePointer = 11
    If Option1.Value = True Then
        rptPrintRunDown.WindowTitle = "PARTS RUNDOWN REPORT - TOTAL RETAIL SALES"
        rptPrintRunDown.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptPrintRunDown.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptPrintRunDown, PMIS_REPORT_PATH & "Rundown\PRR_Retail_Sales.rpt", "MONTH({ORD_HD.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({ORD_HD.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        NEW_LogAudit "V", "PARTS RUNDOWN REPORT", "", "", "Parts", cboMonth & " - " & cboYear, "Total Retail Sales", ""
    End If
    If Option2.Value = True Then
        rptPrintRunDown.WindowTitle = "PARTS RUNDOWN REPORT - TOTAL COST OF SALES"
        rptPrintRunDown.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptPrintRunDown.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptPrintRunDown, PMIS_REPORT_PATH & "Rundown\PRR_Cost_Of_Sales.rpt", "MONTH({ORD_HD.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({ORD_HD.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        NEW_LogAudit "V", "PARTS RUNDOWN REPORT", "", "", "Parts", cboMonth & " - " & cboYear, "Total Cost of Sales", ""
    End If
    If Option3.Value = True Then
        rptPrintRunDown.WindowTitle = "PARTS RUNDOWN REPORT - BEGINNING INVENTORY REPORT"
        rptPrintRunDown.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptPrintRunDown.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

        Dim rsSTKSTAT                                                 As ADODB.Recordset
        Set rsSTKSTAT = New ADODB.Recordset
        If What_month(cboMonth.Text) = 1 Then
            Set rsSTKSTAT = gconDMIS.Execute("Select * from PMIS_StkStat Where MONTH(DATE_GEN) = 12 and YEAR(DATE_GEN) = " & NumericVal(cboYear.Text) - 1 & " Order by DATE_GEN DESC")
            If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
                PrintSQLReport rptPrintRunDown, PMIS_REPORT_PATH & "Rundown\PRR_Beginning_Inventory.rpt", "MONTH({PMIS_STKSTAT.DATE_GEN}) = 12 and YEAR({PMIS_STKSTAT.DATE_GEN}) = " & NumericVal(cboYear.Text) - 1, DMIS_REPORT_Connection, 1
            End If
        Else
            Set rsSTKSTAT = gconDMIS.Execute("Select * from PMIS_StkStat Where MONTH(DATE_GEN) = " & What_month(cboMonth.Text) - 1 & " and YEAR(DATE_GEN) = " & cboYear.Text & " Order by DATE_GEN DESC")
            If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
                PrintSQLReport rptPrintRunDown, PMIS_REPORT_PATH & "Rundown\PRR_Beginning_Inventory.rpt", "MONTH({PMIS_STKSTAT.DATE_GEN}) = " & What_month(cboMonth.Text) - 1 & " and YEAR({PMIS_STKSTAT.DATE_GEN}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
            End If
        End If
        NEW_LogAudit "V", "PARTS RUNDOWN REPORT", "", "", "Parts", cboMonth & " - " & cboYear, "Beginning Inventory Report", ""
    End If
    If Option4.Value = True Then
        rptPrintRunDown.WindowTitle = "PARTS RUNDOWN REPORT - TOTAL PURCHASES REPORT"
        rptPrintRunDown.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptPrintRunDown.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptPrintRunDown.Formulas(11) = "forthemonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
        PrintSQLReport rptPrintRunDown, PMIS_REPORT_PATH & "Rundown\PRR_Total_Purchases.rpt", "MONTH({PO_HD.RRDATE}) = " & What_month(cboMonth.Text) & " and YEAR({PO_HD.RRDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        NEW_LogAudit "V", "PARTS RUNDOWN REPORT", "", "", "Parts", cboMonth & " - " & cboYear, "Total Purchases Report", ""
    End If
    If Option5.Value = True Then
        rptPrintRunDown.WindowTitle = "PARTS RUNDOWN REPORT - INVENTORY ADJUSTMENTS REPORT"
        rptPrintRunDown.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptPrintRunDown.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptPrintRunDown.Formulas(11) = "forthemonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
        PrintSQLReport rptPrintRunDown, PMIS_REPORT_PATH & "Rundown\PRR_Inventory_Adjustments.rpt", "MONTH({PMIS_Adjust.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({PMIS_Adjust.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        NEW_LogAudit "V", "PARTS RUNDOWN REPORT", "", "", "Parts", cboMonth & " - " & cboYear, "Inventory Adjustment Report", ""
    End If
    If Option8.Value = True Then
        rptPrintRunDown.WindowTitle = "PARTS RUNDOWN REPORT - FILL RATE REPORT"
        rptPrintRunDown.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptPrintRunDown.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptPrintRunDown.Formulas(12) = "forthemonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
        PrintSQLReport rptPrintRunDown, PMIS_REPORT_PATH & "Rundown\PRR_Fill_Rate.rpt", "MONTH({PMIS_Demand_Monitoring.Date_Gen}) = " & What_month(cboMonth.Text) & " and YEAR({PMIS_Demand_Monitoring.Date_Gen}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        NEW_LogAudit "V", "PARTS RUNDOWN REPORT", "", "", "Parts", cboMonth & " - " & cboYear, "Fill Rate Report", ""
    End If

    If Option9.Value = True Then
        rptPrintRunDown.WindowTitle = "PARTS RUNDOWN REPORT - ORDERED PARTS BY CATEGORY"
        rptPrintRunDown.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptPrintRunDown.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptPrintRunDown.Formulas(12) = "forthemonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
        PrintSQLReport rptPrintRunDown, PMIS_REPORT_PATH & "Rundown\PRR_Ordered_Parts_By_Category.rpt", "MONTH({PMIS_AllDayTran.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({PMIS_AllDayTran.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        NEW_LogAudit "V", "PARTS RUNDOWN REPORT", "", "", "Parts", cboMonth & " - " & cboYear, "Ordered Parts by Category Report", ""
    End If
    If Option10.Value = True Then
        rptPrintRunDown.WindowTitle = "PARTS RUNDOWN REPORT - PARTS BACK ORDER REPORT"
        rptPrintRunDown.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptPrintRunDown.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptPrintRunDown.Formulas(11) = "forthemonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
        If What_month(cboMonth.Text) = Month(Now) Then
            Screen.MousePointer = vbHourglass
            PrintSQLReport rptPrintRunDown, PMIS_REPORT_PATH & "Rundown\PRR_Back_Ordered_Curr.rpt", "", DMIS_REPORT_Connection, 1
            Screen.MousePointer = vbDefault
        Else
            Screen.MousePointer = vbHourglass
            PrintSQLReport rptPrintRunDown, PMIS_REPORT_PATH & "Rundown\PRR_Back_Ordered_Prev.rpt", "", DMIS_REPORT_Connection, 1
            Screen.MousePointer = vbDefault
        End If
        NEW_LogAudit "V", "PARTS RUNDOWN REPORT", "", "", "Parts", cboMonth & " - " & cboYear, "Parts Back-Order Report", ""
    End If
    If Option11.Value = True Then
        Call ShowExcel
        NEW_LogAudit "V", "PARTS RUNDOWN REPORT", "", "", "Parts", cboMonth & " - " & cboYear, "Parts Rundown Report - Excel", ""
    End If
    If optMAD.Value = True Then
        Screen.MousePointer = vbHourglass
        rptPrintRunDown.WindowTitle = "PRR - Moving Average Demand"
        rptPrintRunDown.ReportTitle = "Moving Average Demand of Parts"
        rptPrintRunDown.Formulas(11) = "monthprint = 'For the Year " & cboYear.Text & "'"
        Dim rsMAD As ADODB.Recordset
        Set rsMAD = New ADODB.Recordset
        Set rsMAD = gconDMIS.Execute("SELECT TRANDATE FROM PMIS_ALLDAYTRAN WHERE TYPE = 'P' AND TRANTYPE IN ('CSH','CHG','DR','RIV') AND (STATUS = 'P' OR STATUS = 'B') AND YEAR(TRANDATE) = " & cboYear)
        If Not rsMAD.EOF And Not rsMAD.BOF Then
           PrintSQLReport rptPrintRunDown, PMIS_REPORT_PATH & "Rundown\PRR_Moving_Average_Demand.rpt", "YEAR({ALLDAYTRAN.TRANDATE}) = " & cboYear, DMIS_REPORT_Connection, 1
        Else
            ShowNoRecord
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Set rsMAD = Nothing
        Screen.MousePointer = vbDefault
        NEW_LogAudit "V", "PARTS RUNDOWN REPORT", "", "", "Parts", cboMonth & " - " & cboYear, "Moving Average Demand", ""
    End If
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PARTS RUNDOWN REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "PARTS RUNDOWN REPORT", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    FillcboYear cboYear
    cboYear.Text = Year(LOGDATE)
    cboMonth.Text = The_month(Month(LOGDATE))

    If PRR_BUTTON_CLICK = 1 Then Option1.Value = True    'PRR - Retail Sales
    If PRR_BUTTON_CLICK = 2 Then Option2.Value = True    'PRR - Cost of Sales
    If PRR_BUTTON_CLICK = 3 Then Option3.Value = True    'PRR - Beginning Inventory Report
    If PRR_BUTTON_CLICK = 4 Then Option4.Value = True    'PRR - Total Purchases Report
    If PRR_BUTTON_CLICK = 5 Then Option5.Value = True    'PRR - Inventory Adjustments
    'If PRR_BUTTON_CLICK = 6 Then Option6.Value = True    'PRR - Parts Moving Average Demand
    'If PRR_BUTTON_CLICK = 7 Then Option7.Value = True    'PRR - Inventory Gross Return
    If PRR_BUTTON_CLICK = 8 Then Option8.Value = True    'PRR - Fill Rate Reports
    If PRR_BUTTON_CLICK = 9 Then Option10.Value = True    'PRR - Parts Back-Order Report
    If PRR_BUTTON_CLICK = 10 Then Option11.Value = True   'PRR - Excel Report

    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_PrintRankfle = Nothing
    UnloadForm Me
End Sub


Private Sub Option1_Click()
    cboMonth.Enabled = True
End Sub

Private Sub Option10_Click()
    cboMonth.Enabled = True
End Sub

Private Sub Option11_Click()
    prgExcelGen.Text = ""
End Sub

Private Sub Option2_Click()
    cboMonth.Enabled = True
End Sub

Private Sub Option3_Click()
    cboMonth.Enabled = True
End Sub

Private Sub Option4_Click()
    cboMonth.Enabled = True
End Sub

Private Sub Option5_Click()
    cboMonth.Enabled = True
End Sub

Private Sub Option8_Click()
    cboMonth.Enabled = True
End Sub

Private Sub Option9_Click()
    cboMonth.Enabled = True
End Sub

Private Sub optMAD_Click()
    cboMonth.Enabled = False
End Sub

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_Report_VehicleSalesHyundai 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hyundai Vehicles Sales Report"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_VehicleSalesHyundai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   4230
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   3720
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Generating Reports.........."
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   480
      TabIndex        =   22
      Top             =   1260
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label lblLoading 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblGenerating 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   240
      TabIndex        =   12
      Top             =   3510
      Visible         =   0   'False
      Width           =   1185
      Begin VB.OptionButton Option1 
         Caption         =   "Source Of Sales"
         Height          =   240
         Left            =   -120
         TabIndex        =   21
         Top             =   -60
         Width           =   3825
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sales Consultant Productive Analysis-A"
         Height          =   240
         Left            =   -120
         TabIndex        =   20
         Top             =   1875
         Width           =   3825
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Sales Consultant Productive Analysis"
         Height          =   240
         Left            =   -120
         TabIndex        =   19
         Top             =   2190
         Width           =   3825
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Payment Mode"
         Height          =   240
         Left            =   -120
         TabIndex        =   18
         Top             =   255
         Width           =   3825
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Quality Sales Ration Analysis"
         Height          =   240
         Left            =   -120
         TabIndex        =   17
         Top             =   585
         Width           =   3825
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Monthly Dealer Retail Sales"
         Height          =   240
         Left            =   750
         TabIndex        =   16
         Top             =   930
         Width           =   3825
      End
      Begin VB.OptionButton Option7 
         Caption         =   "SC Consolidated Productivity"
         Height          =   240
         Left            =   -120
         TabIndex        =   15
         Top             =   2520
         Width           =   3825
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Dealer Ending Inventory"
         Height          =   240
         Left            =   -90
         TabIndex        =   14
         Top             =   1545
         Width           =   3825
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Daily Dealer Retail Sales"
         Height          =   240
         Left            =   -90
         TabIndex        =   13
         Top             =   1230
         Width           =   3825
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4170
      Top             =   5940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cboReportType 
      Height          =   360
      ItemData        =   "Report_VehicleSalesHyundai.frx":0E42
      Left            =   480
      List            =   "Report_VehicleSalesHyundai.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   870
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2280
      MouseIcon       =   "Report_VehicleSalesHyundai.frx":0E46
      MousePointer    =   99  'Custom
      Picture         =   "Report_VehicleSalesHyundai.frx":0F98
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   3570
      Width           =   825
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1470
      MouseIcon       =   "Report_VehicleSalesHyundai.frx":13E3
      MousePointer    =   99  'Custom
      Picture         =   "Report_VehicleSalesHyundai.frx":1535
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   3570
      Width           =   825
   End
   Begin VB.PictureBox picRange 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   450
      ScaleHeight     =   975
      ScaleWidth      =   4365
      TabIndex        =   3
      Top             =   2730
      Width           =   4365
      Begin MSComCtl2.DTPicker dtDay 
         Height          =   390
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   56950785
         CurrentDate     =   39203
      End
      Begin VB.Label Label1 
         Caption         =   "As of"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   30
         Width           =   885
      End
   End
   Begin VB.PictureBox picMonthly 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   420
      ScaleHeight     =   1635
      ScaleWidth      =   4365
      TabIndex        =   4
      Top             =   1200
      Width           =   4365
      Begin VB.TextBox txtMonthly_Year 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   60
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1020
         Width           =   3165
      End
      Begin VB.ComboBox cboMonthly_Month 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   7
         Top             =   360
         Width           =   3195
      End
      Begin VB.Label Label6 
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   10
         Top             =   750
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "For the Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   8
         Top             =   90
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   60
      Picture         =   "Report_VehicleSalesHyundai.frx":19D4
      Top             =   30
      Width           =   1500
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   825
      Left            =   -30
      TabIndex        =   11
      Top             =   0
      Width           =   4305
      _Version        =   655364
      _ExtentX        =   7594
      _ExtentY        =   1455
      _StockProps     =   14
      Caption         =   "Sales Report    "
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   2
      ForeColor       =   4194304
   End
End
Attribute VB_Name = "frmSMIS_Report_VehicleSalesHyundai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xKawnt                                         As Integer
Dim xCount                                         As Integer
Dim rsReleased                                     As ADODB.Recordset
Dim xlApp                                          As Excel.Application
Dim xlBook                                         As Excel.Workbook
Dim xlSheet1                                       As Excel.Worksheet
Dim HARI_Certified                                 As Integer
Dim NON_HARI_CERTIFIED                             As Integer


'Sub PRINT_DAILY_VEHICLE_RETAIL_SALES()
'    Dim SQL                                                           As String
'    Dim rsCust                                                        As ADODB.Recordset
'    If Len(Dir(SMIS_REPORT_PATH & "SMIS_EXCEL\DAILY VEHICLE RETAIL SALES.xlt")) = 0 Then
'        MsgBox "Excel Directory For Sales Managment Information Could Not be Located", vbInformation
'        Exit Sub
'    End If
'
'    ''If gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESORDER WHERE month(DateReleased)=" & What_month(cboMonthly_Month) & " AND year(DateReleased)= " & (txtMonthly_Year) & "  AND STATUS<>'C'").Fields(0).Value = 0 Then
'    ''   MsgSpeech " NO SALES RECORD FOR THE DATE "
'    ''   Exit Sub
'    ''End If
'
'
'    ''Set rsCust = gconDMIS.Execute("SELECT DATERELEASED,MODELDESCRIPTION,IGNKEY_NO,RIGHT(FRAMENO,6) AS FRAMENO,COLOR,CUSTNAME,SALESAE FROM SMIS_PURCHAGREE WHERE (DAY(DateReleased)=" & Day(dtDay.Value) & " AND month(DateReleased)= " & Month(dtDay) & " AND year(DateReleased)= " & Year(dtDay) & ") ORDER BY INVOICEDDATE")
'    Set rsCust = gconDMIS.Execute("SELECT DATERELEASED,MODELDESCRIPTION,IGNKEY_NO,FRAMENO,COLOR,CUSTNAME,VI_NO,SALESAE FROM SMIS_PURCHAGREE WHERE (DAY(DateReleased)=" & Day(dtDay.Value) & " AND month(DateReleased)= " & Month(dtDay) & " AND year(DateReleased)= " & Year(dtDay) & ") ORDER BY INVOICEDDATE")
'
'    Dim xlApp
'    Dim xlBook
'    Dim xlSheet1
'    Set xlApp = CreateObject("Excel.Application")
'    Set xlBook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "SMIS_EXCEL\DAILY VEHICLE RETAIL SALES.xlt")
'    Set xlSheet1 = xlBook.Worksheets(1)
'
'
'    Dim i                                                             As Integer
'    Dim j                                                             As Integer
'
'
'    xlSheet1.Cells(6, "C") = COMPANY_NAME
'    xlSheet1.Cells(7, "C") = Format(dtDay, "MMM") & " " & Year(dtDay)
'
'    Do While Not rsCust.EOF
'        Dim rsGetOrNum2                                               As ADODB.Recordset
'        Dim getOrNum2                                                 As String
'        Dim Rnum2                                                     As String
'
'        getOrNum2 = Null2String(rsCust!VI_NO)
'        If Not rsCust.EOF And Not rsCust.BOF Then
'
'            Set rsGetOrNum2 = gconDMIS.Execute("Select OR_NUM from CMIS_off_Dt where TRANTYPE = 'VI' and invoiceno = '" & getOrNum2 & "'")
'            If Not rsGetOrNum2.EOF And Not rsGetOrNum2.BOF Then
'                Rnum2 = Null2String(rsGetOrNum2!OR_NUM)
'            End If
'        End If
'
'
'        xlSheet1.Cells(15 + j, "A") = j + 1
'        xlSheet1.Cells(15 + j, "B") = Null2String(rsCust!DateReleased)
'        xlSheet1.Cells(15 + j, "C") = Null2String(rsCust!modeldescription)
'        xlSheet1.Cells(15 + j, "D") = Null2String(rsCust!IGNKEY_NO)
'        xlSheet1.Cells(15 + j, "E") = Null2String(rsCust!frameno)
'        xlSheet1.Cells(15 + j, "F") = Null2String(rsCust!Color)
'        xlSheet1.Cells(15 + j, "G") = Null2String(rsCust!CustName)
'        xlSheet1.Cells(15 + j, "H") = Rnum2
'        xlSheet1.Cells(15 + j, "I") = Null2String(rsCust!salesae)
'        j = j + 1
'        rsCust.MoveNext
'    Loop
'
'    Set rsCust = gconDMIS.Execute("SELECT DISTINCT  MODEL , COUNT(*) AS TCOUNT  FROM SMIS_PURCHAGREE WHERE (DAY(DateReleased)=" & Day(dtDay.Value) & " AND month(DateReleased)= " & Month(dtDay) & " AND year(DateReleased)= " & Year(dtDay) & ") GROUP BY MODEL ")
'    Dim TOTALUNIT                                                     As Integer
'    j = 0
'    Do While Not rsCust.EOF
'
'        xlSheet1.Cells(34 + j, "C") = Null2String(rsCust!Model)
'        xlSheet1.Cells(34 + j, "D") = Null2String(rsCust!TCOUNT)
'        TOTALUNIT = TOTALUNIT + Null2String(rsCust!TCOUNT)
'        j = j + 1
'        rsCust.MoveNext
'    Loop
'
'    xlSheet1.Cells(34 + j, "C") = "TOTAL"
'    xlSheet1.Cells(34 + j, "D") = TOTALUNIT
'
'    xlApp.Visible = True
'    Set xlBook = Nothing
'    Set xlSheet1 = Nothing
'    Set xlApp = Nothing
'End Sub

Sub PRINT_DEALER_ENDING_INVENTORY()
    Dim SQL                                            As String
    Dim rsCust                                         As ADODB.Recordset
    If Len(Dir(SMIS_REPORT_PATH & "SMIS_EXCEL\DEALER ENDING INVENTORY.xlt")) = 0 Then
        MsgBox "Excel Directory For Sales Managment Information Could Not be Located", vbInformation
        Exit Sub
    End If


    If gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_MRRINV where released=0 and status <> 'C'").Fields(0).Value = 0 Then
        MsgSpeech " NO RECEIVING  RECORD"
        Exit Sub
    End If



    'Set rsCust = gconDMIS.Execute("SELECT DESCRIPT , IGNKEY,RIGHT(VINO,6) AS VINNO ,COLOR,PULLOUTDATE  FROM SMIS_MRRINV_TABLE where released=0 and status <> 'C' ORDER BY DATERECEIVED")
    
    
    Set rsCust = gconDMIS.Execute("SELECT DESCRIPT , IGNKEY,RIGHT(VINO,6) AS VINNO ,COLOR,PULLOUTDATE  FROM SMIS_MRRINV_TABLE where released=0 and status <> 'C' ORDER BY DESCRIPT ASC")


    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet1                                       As Excel.Worksheet
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "SMIS_EXCEL\DEALER ENDING INVENTORY.xlt")
    Set xlSheet1 = xlBook.Worksheets(1)


    Dim i                                              As Integer
    Dim j                                              As Integer


    xlSheet1.Cells(6, "B") = COMPANY_NAME
    xlSheet1.Cells(8, "B") = cboMonthly_Month & " " & txtMonthly_Year
    Do While Not rsCust.EOF
        xlSheet1.Cells(14 + j, "A") = j + 1
        xlSheet1.Cells(14 + j, "A").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(14 + j, "B") = Null2String(rsCust!DESCRIPT)
        xlSheet1.Cells(14 + j, "B").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(14 + j, "C") = Null2String(rsCust!ignkey)
        xlSheet1.Cells(14 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(14 + j, "D") = Null2String(rsCust!VINNO)
        xlSheet1.Cells(14 + j, "D").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(14 + j, "E") = Null2String(rsCust!Color)
        xlSheet1.Cells(14 + j, "E").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(14 + j, "F") = Null2String(rsCust!PullOutDate)
        xlSheet1.Cells(14 + j, "F").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(14 + j, "G") = DateDiff("d", Null2String(rsCust!PullOutDate), Date)
        xlSheet1.Cells(14 + j, "G").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(14 + j, "H").BorderAround ColorIndex:=1, Weight:=xlThin
        j = j + 1
        rsCust.MoveNext
    Loop



    xlSheet1.Cells(15 + j, "B") = "MODEL SUMMARY"
    xlSheet1.Cells(15 + j, "B").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(15 + j, "B").Font.Bold = True
    xlSheet1.Cells(15 + j, "C") = "UNIT"
    xlSheet1.Cells(15 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(15 + j, "C").Font.Bold = True
    Set rsCust = gconDMIS.Execute("SELECT DISTINCT  MODEL , COUNT(*) AS TCOUNT  FROM SMIS_MRRINV where released=0 and Status <> 'C' GROUP BY MODEL ")
    Dim TOTALUNIT                                      As Integer
    Do While Not rsCust.EOF
        j = j + 1
        xlSheet1.Cells(15 + j, "B") = Null2String(rsCust!Model)
        xlSheet1.Cells(15 + j, "B").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(15 + j, "C") = Null2String(rsCust!TCOUNT)
        xlSheet1.Cells(15 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
        TOTALUNIT = TOTALUNIT + Null2String(rsCust!TCOUNT)
        rsCust.MoveNext
    Loop
    xlSheet1.Cells(16 + j, "B") = "TOTAL"
    xlSheet1.Cells(16 + j, "B").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(16 + j, "B").Font.Bold = True
    xlSheet1.Cells(16 + j, "C") = TOTALUNIT
    xlSheet1.Cells(16 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(16 + j, "C").Font.Bold = True

    xlSheet1.Cells(8 + j, "E") = "Prepared by:"
    xlSheet1.Cells(8 + j, "E").Font.Bold = True
    xlSheet1.Cells(9 + j, "E").Font.Underline = True
    xlSheet1.Cells(10 + j, "E") = "Approved by:"
    xlSheet1.Cells(10 + j, "E").Font.Bold = True
    xlSheet1.Cells(11 + j, "E").Font.Underline = True
    xlSheet1.Cells(12 + j, "E") = "Authorized Dealer Representative"
    xlSheet1.Cells(12 + j, "E").Font.Bold = True
    xlApp.Visible = True
    Set xlBook = Nothing
    Set xlSheet1 = Nothing
    Set xlApp = Nothing


End Sub

Sub PRINT_DEALER_SALES_CONSULTANTS_PERFORMANCE_SHEET()
    Dim rsPurch                                        As ADODB.Recordset
    Dim RsSAE                                          As ADODB.Recordset
    Dim rsCountReleased                                As ADODB.Recordset
    Dim rsTotalPerMonth                                As ADODB.Recordset
    Dim rsSumPerYear                                   As ADODB.Recordset
    Dim rsGet_CS                                       As ADODB.Recordset
    Dim xCount                                         As Integer
    Dim xCOUNT_MONTH                                   As Integer
    
    xKawnt = 13
    
    Frame2.Visible = True
    
    'Timer1.Enabled = True
    'Dim strPath                                        As String
    
'    strPath = App.Path & "\loading.gif"
'    WebBrowser1.Navigate2 (strPath)
'    WebBrowser1.Navigate ("about:html body scroll='no'bgcolor='#CCFFFF' img src= " & strPath & " >/img body /html ")
'    WebBrowser1.Refresh2
    
    Set xlApp = New Excel.Application
    
    Set xlBook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "SMIS_EXCEL\SC Consolidated Productivity.xlt")
    
    
    Set xlSheet1 = xlBook.Worksheets(1)
    
    xlSheet1.Cells(7, "D") = COMPANY_NAME
    xlSheet1.Cells(8, "D") = cboMonthly_Month & " " & txtMonthly_Year
    xlSheet1.Cells(12, "Q").Borders(xlEdgeTop).Weight = xlThick
    xlSheet1.Cells(12, "Q").Borders(xlEdgeLeft).Weight = xlThick
    xlSheet1.Cells(12, "Q").Borders(xlEdgeRight).Weight = xlThick
    
    Set rsPurch = gconDMIS.Execute("Select DISTINCT SALESAE FROM SMIS_PURCHAGREE where STATUS = 'P' ORDER BY SALESAE ASC")
    If Not rsPurch.EOF And Not rsPurch.BOF Then
        Do While Not rsPurch.EOF
            Set RsSAE = gconDMIS.Execute("Select SALESAE from SMIS_PURCHAGREE Where SALESAE = '" & Null2String(rsPurch!salesae) & "'")
                xlSheet1.Cells(xKawnt, "B") = Null2String(rsPurch!salesae)
                xlSheet1.Cells(xKawnt, "B").BorderAround ColorIndex:=1, Weight:=xlThin
                lblGenerating.Caption = UCase(Null2String(rsPurch!salesae))
             For xCount = 1 To 12
             lblLoading.Caption = "Loading...."
                Set rsReleased = gconDMIS.Execute("Select COUNT(DATERELEASED) AS MODELCOUNT from SMIS_PURCHAGREE WHERE SALESAE = '" & rsPurch!salesae & "' and MONTH(DATERELEASED) = " & xCount & " AND YEAR(DATERELEASED) = '" & txtMonthly_Year & "' and STATUS = 'P'")
                    If xCount = 1 Then
                        xlSheet1.Cells(xKawnt, "D") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "D").Cells.HorizontalAlignment = xlCenter
                    ElseIf xCount = 2 Then
                        xlSheet1.Cells(xKawnt, "E") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "E").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "E").Cells.HorizontalAlignment = xlCenter
                    ElseIf xCount = 3 Then
                        xlSheet1.Cells(xKawnt, "F") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "F").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "F").Cells.HorizontalAlignment = xlCenter
                    ElseIf xCount = 4 Then
                        xlSheet1.Cells(xKawnt, "G") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "G").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "G").Cells.HorizontalAlignment = xlCenter
                    ElseIf xCount = 5 Then
                        xlSheet1.Cells(xKawnt, "H") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "H").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "H").Cells.HorizontalAlignment = xlCenter
                    ElseIf xCount = 6 Then
                        xlSheet1.Cells(xKawnt, "I") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "I").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "I").Cells.HorizontalAlignment = xlCenter
                    ElseIf xCount = 7 Then
                        xlSheet1.Cells(xKawnt, "J") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "J").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "J").Cells.HorizontalAlignment = xlCenter
                    ElseIf xCount = 8 Then
                        xlSheet1.Cells(xKawnt, "K") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "K").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "K").Cells.HorizontalAlignment = xlCenter
                    ElseIf xCount = 9 Then
                        xlSheet1.Cells(xKawnt, "L") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "L").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "L").Cells.HorizontalAlignment = xlCenter
                    ElseIf xCount = 10 Then
                        xlSheet1.Cells(xKawnt, "M") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "M").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "M").Cells.HorizontalAlignment = xlCenter
                    ElseIf xCount = 11 Then
                        xlSheet1.Cells(xKawnt, "N") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "N").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "N").Cells.HorizontalAlignment = xlCenter
                    ElseIf xCount = 12 Then
                        xlSheet1.Cells(xKawnt, "O") = Null2String(rsReleased!MODELCOUNT)
                        xlSheet1.Cells(xKawnt, "O").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet1.Cells(xKawnt, "O").Cells.HorizontalAlignment = xlCenter
                    End If
                lblLoading.Caption = "Loading........"
                Set RsSAE = Nothing
             Next xCount
             
             Set rsCountReleased = gconDMIS.Execute("Select COUNT(DATERELEASED) AS COUNTRELEASED from SMIS_PURCHAGREE WHERE SALESAE = '" & rsPurch!salesae & "' AND YEAR(DATERELEASED) = '" & txtMonthly_Year & "' and status = 'P'")
             xlSheet1.Cells(xKawnt, "Q") = Null2String(rsCountReleased!COUNTRELEASED)
             xlSheet1.Cells(xKawnt, "Q").BorderAround ColorIndex:=1, Weight:=xlThin
             
             xlSheet1.Range("Q" & xKawnt & ":" & "Q" & xKawnt).Borders(xlEdgeLeft).Weight = xlThick
             xlSheet1.Range("Q" & xKawnt & ":" & "Q" & xKawnt).Borders(xlEdgeRight).Weight = xlThick
             xlSheet1.Cells(xKawnt, "Q").Interior.Color = &H80FFFF
             xlSheet1.Cells(xKawnt, "Q").Font.Bold = True
             xlSheet1.Cells(xKawnt, "Q").Cells.HorizontalAlignment = xlCenter
             
             xlSheet1.Range("D" & xKawnt & ":" & "O" & xKawnt).Font.Color = &HFF0000
            xKawnt = xKawnt + 1
            rsPurch.MoveNext
        Loop
    End If
    Set rsPurch = Nothing
            
    For xCOUNT_MONTH = 1 To 12
    lblLoading.Caption = "Loading...."
        Set rsTotalPerMonth = gconDMIS.Execute("SELECT COUNT(DATERELEASED) AS COUNT_PERM_MONTH FROM SMIS_PURCHAGREE WHERE MONTH(DATERELEASED) = " & xCOUNT_MONTH & " and year(DATERELEASED) = '" & txtMonthly_Year & "' and status = 'P'")
            If xCOUNT_MONTH = 1 Then
                xlSheet1.Cells(xKawnt, "D") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "D").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "D").Font.Bold = True
            ElseIf xCOUNT_MONTH = 2 Then
                xlSheet1.Cells(xKawnt, "E") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "E").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "E").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "E").Font.Bold = True
            ElseIf xCOUNT_MONTH = 3 Then
                xlSheet1.Cells(xKawnt, "F") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "F").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "F").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "F").Font.Bold = True
            ElseIf xCOUNT_MONTH = 4 Then
                xlSheet1.Cells(xKawnt, "G") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "G").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "G").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "G").Font.Bold = True
            ElseIf xCOUNT_MONTH = 5 Then
                xlSheet1.Cells(xKawnt, "H") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "H").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "H").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "H").Font.Bold = True
            ElseIf xCOUNT_MONTH = 6 Then
                xlSheet1.Cells(xKawnt, "I") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "I").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "I").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "I").Font.Bold = True
            ElseIf xCOUNT_MONTH = 7 Then
                xlSheet1.Cells(xKawnt, "J") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "J").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "J").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "J").Font.Bold = True
            ElseIf xCOUNT_MONTH = 8 Then
                xlSheet1.Cells(xKawnt, "K") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "K").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "K").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "K").Font.Bold = True
            ElseIf xCOUNT_MONTH = 9 Then
                xlSheet1.Cells(xKawnt, "L") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "L").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "L").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "L").Font.Bold = True
            ElseIf xCOUNT_MONTH = 10 Then
                xlSheet1.Cells(xKawnt, "M") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "M").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "M").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "M").Font.Bold = True
            ElseIf xCOUNT_MONTH = 11 Then
                xlSheet1.Cells(xKawnt, "N") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "N").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "N").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "N").Font.Bold = True
            ElseIf xCOUNT_MONTH = 12 Then
                xlSheet1.Cells(xKawnt, "O") = Null2String(rsTotalPerMonth!COUNT_PERM_MONTH)
                xlSheet1.Cells(xKawnt, "O").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(xKawnt, "O").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(xKawnt, "O").Font.Bold = True
            End If
        Set rsTotalPerMonth = Nothing
        lblLoading.Caption = "Loading........"
    Next xCOUNT_MONTH
    
    Set rsSumPerYear = gconDMIS.Execute("Select count(DATERELEASED) AS COUNT_PER_YEAR from SMIS_PURCHAGREE WHERE YEAR(DATERELEASED) = '" & txtMonthly_Year & "' and status = 'P'")
             xlSheet1.Cells(xKawnt, "Q") = Null2String(rsSumPerYear!COUNT_PER_YEAR)
             xlSheet1.Cells(xKawnt, "Q").BorderAround ColorIndex:=1, Weight:=xlThin
             xlSheet1.Cells(xKawnt, "Q").Borders(xlEdgeLeft).Weight = xlThick
             xlSheet1.Cells(xKawnt, "Q").Cells.HorizontalAlignment = xlCenter
    Set rsSumPerYear = Nothing
    
    xlSheet1.Cells(xKawnt, "B") = "TOTAL"
    xlSheet1.Range("B" & xKawnt & ":" & "Q" & xKawnt).BorderAround ColorIndex:=1, Weight:=xlThick
    
    xlSheet1.Range("B" & xKawnt & ":" & "Q" & xKawnt).Interior.Color = &H80FFFF
    xlSheet1.Range("B" & xKawnt & ":" & "Q" & xKawnt).Font.Bold = True
    
    xKawnt = xKawnt + 2
    xlSheet1.Cells(xKawnt, "B") = "# SC CERTIFIED    :"
    xlSheet1.Range("D" & xKawnt & ":" & "O" & xKawnt).Cells.BorderAround ColorIndex:=1, Weight:=xlThin
    
    'COUNT SC CERTIFIED
    Call COUNT_SC_CERTIFIED_PER_MONTH
    
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "B") = "# SC NON-CERTIFIED    :"
    xlSheet1.Range("D" & xKawnt & ":" & "O" & xKawnt).Cells.BorderAround ColorIndex:=1, Weight:=xlThin
    
    'COUNT SC_NON_CERTIFIED
    Call COUNT_SC_NON_CERTIFIED_PER_MONTH
    
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "B") = "# SC SALARIED    :"
    xlSheet1.Range("D" & xKawnt & ":" & "O" & xKawnt).Cells.BorderAround ColorIndex:=1, Weight:=xlThin
    
    'COUNT SALARIED SAE
    Call COUNT_SC_SALARIED
    
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "B") = "# SC COMMISSIONED    :"
    xlSheet1.Range("D" & xKawnt & ":" & "O" & xKawnt).Cells.BorderAround ColorIndex:=1, Weight:=xlThin
    
    'COUNT COMMISSION SAE
    Call COUNT_SC_COMMISSION
    
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "B") = "TOTAL SC    :"
    xlSheet1.Range("D" & xKawnt & ":" & "O" & xKawnt).Borders(xlEdgeTop).Weight = xlThick
    xlSheet1.Cells(xKawnt, "D").Borders(xlEdgeLeft).Weight = xlThick
    xlSheet1.Cells(xKawnt, "O").Borders(xlEdgeRight).Weight = xlThick
    
    'COUNT TOTAL SC FOR EVERY MONTH
    Call COUNT_TOTAL_SC_PER_MONTH
    
    xlSheet1.Cells(xKawnt, "B").Font.Bold = True
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "B") = "# SC RESIGNED    :"
    xlSheet1.Range("D" & xKawnt & ":" & "O" & xKawnt).Borders(xlBottom).Weight = xlThick
    xlSheet1.Cells(xKawnt, "D").Borders(xlEdgeLeft).Weight = xlThick
    xlSheet1.Cells(xKawnt, "O").Borders(xlEdgeRight).Weight = xlThick
    
    'COUNT TOTAL SC RESIGNED
    Call COUNT_SC_RESIGNED_THIS_MONTH
    
    'xlSheet1.Range("D" & xKawnt & ":" & "O" & xKawnt).Cells.BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(xKawnt, "B").Font.Bold = True
    
    xKawnt = xKawnt + 4
    xlSheet1.Cells(xKawnt, "B") = "Prepared by:"
    xlSheet1.Cells(xKawnt, "B").Font.Bold = True
    xlSheet1.Cells(xKawnt, "K") = "Approved by:"
    xlSheet1.Cells(xKawnt, "K").Font.Bold = True
    
    xKawnt = xKawnt + 2
    xlSheet1.Cells(xKawnt, "B") = "SALES ADMIN. SUPERVISOR"
    xlSheet1.Cells(xKawnt, "B").Borders(xlEdgeTop).Weight = xlThin
    xlSheet1.Cells(xKawnt, "B").Font.Bold = True
    
    xKawnt = xKawnt + 2
    xlSheet1.Cells(xKawnt, "A") = "NOTE:"
    xlSheet1.Cells(xKawnt, "A").Font.Bold = True
    xlSheet1.Cells(xKawnt, "A").Font.Color = &HFF&
    
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "A") = "* Please submit this report including all your SCs even without sale (for monitoring purposes),"
    xlSheet1.Cells(xKawnt, "A").Font.Color = &HFF&
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "A") = "otherwise HARI will consider the SC as a resigned employee"
    xlSheet1.Cells(xKawnt, "A").Font.Color = &HFF&
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "A") = "* To be submitted on the 2nd day of the following month."
    xlSheet1.Cells(xKawnt, "A").Font.Color = &HFF&
    xlSheet1.Cells(xKawnt, "A").Font.Bold = True
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "A") = "* If it falls on a weekend, submission of the report shall be on the first"
    xlSheet1.Cells(xKawnt, "A").Font.Color = &HFF&
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "A") = "working day of the succeeding month."
    xlSheet1.Cells(xKawnt, "A").Font.Color = &HFF&
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "A") = "* Fax no. 812-1556/894-5866"
    xlSheet1.Cells(xKawnt, "A").Font.Color = &HFF&
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "A") = "* Email thru gmalabuyoc@hyundai-asia.com"
    xlSheet1.Cells(xKawnt, "A").Font.Color = &HFF&
    xKawnt = xKawnt + 1
    xlSheet1.Cells(xKawnt, "A") = " - cc scalma@hyundai-asia.com"
    xlSheet1.Cells(xKawnt, "A").Font.Color = &HFF&
    
    'CODE FOR PER SALES ACCOUNT EXECUTIVE SHEET
    Call SALES_ACCOUNT_EXECUTIVE_PER_SHEET
    
    MsgBox "Reports Generation Completed", vbInformation, "SC PRODUCTIVITY"
    Frame2.Visible = False
    
    
    'Sheets("CONSO").Move Before:=ActiveWorkbook.Sheets(1)
    
    xlApp.Visible = True
    Set xlBook = Nothing
    Set xlSheet1 = Nothing
    Set xlApp = Nothing

End Sub

Sub SALES_ACCOUNT_EXECUTIVE_PER_SHEET()
    Dim rsSOLO_SAE                  As ADODB.Recordset
    Dim rsCOUNT_SAE_RELEASE         As ADODB.Recordset
    Dim rsModel                     As ADODB.Recordset
    Dim rsSAE_CODE                  As ADODB.Recordset
    Dim rsSAE_INFO                  As ADODB.Recordset
    Dim rsSUM_MODEL_PER_YEAR        As ADODB.Recordset
    Dim rsSUM_WHOLE_YEAR            As ADODB.Recordset
    
    Dim x_MONTH                     As Integer
    Dim XXX                         As Integer
    Dim YYY                         As Integer
    
    Dim X_HIRED                     As String
    Dim X_RESIGNED                  As String
    Dim X_TIN                       As String
    Dim X_COMMISSION                As String
    Dim X_IS_CERTIFIED              As String
    Dim Y_N                         As String
    Dim X_CONTRACT                  As String
    
    XXX = 17
    
    Set rsSOLO_SAE = gconDMIS.Execute("Select DISTINCT SALESAE as NIYM FROM SMIS_PURCHAGREE where STATUS = 'P' ORDER BY SALESAE DESC")
    If Not rsSOLO_SAE.EOF And Not rsSOLO_SAE.EOF Then
        Do While Not rsSOLO_SAE.EOF
            
            Set rsSAE_CODE = gconDMIS.Execute("Select SAECODE from SMIS_vw_Srep where NAME = '" & rsSOLO_SAE!NIYM & "'")
            If Not rsSAE_CODE.EOF And Not rsSAE_CODE.BOF Then
                Set rsSAE_INFO = gconDMIS.Execute("Select DATE_HIRED,DATE_RESIGNED,TINNO,HARI_CERTIFIED,CONTRACT from SMIS_SALESTEAM")
                If Not rsSAE_INFO.EOF And Not rsSAE_INFO.BOF Then
                    X_HIRED = Null2String(rsSAE_INFO!Date_Hired)
                    X_RESIGNED = Null2String(rsSAE_INFO!Date_Resigned)
                    X_TIN = Null2String(rsSAE_INFO!TINNO)
                    X_IS_CERTIFIED = Null2String(rsSAE_INFO!HARI_Certified)
                    If Null2String(rsSAE_INFO!Contract) = "C" Or IsNull(rsSAE_INFO!Contract) = True Then
                        X_CONTRACT = "COMMISSION"
                    Else
                        X_CONTRACT = "SALARIED"
                    End If
                End If
            End If
            
            If IsNull(X_IS_CERTIFIED) = True Or X_IS_CERTIFIED = "N" Then
                Y_N = "NO"
            Else
                Y_N = "YES"
            End If
                 
            
            Set xlSheet1 = xlBook.Worksheets.Add
            
            xlSheet1.Name = Null2String(rsSOLO_SAE!NIYM)
            xlSheet1.Cells(5, "A") = "Dealers' Sales Consultants (SC) Performance Sheet"
            xlSheet1.Cells(5, "A").Font.Bold = True
            xlSheet1.Range("A5:L5").Interior.Color = &HFF0000
            xlSheet1.Cells(6, "A") = "SC Information"
            xlSheet1.Cells(6, "A").Font.Italic = True
            xlSheet1.Cells(7, "A") = "I"
            xlSheet1.Cells(7, "B") = "Dealer:"
            xlSheet1.Cells(7, "B").Font.Bold = True
            xlSheet1.Cells(7, "D") = COMPANY_NAME
            xlSheet1.Cells(7, "D").Font.Bold = True
            xlSheet1.Range("D7:G7").Borders(xlEdgeBottom).Weight = xlThin

            xlSheet1.Cells(8, "A") = "II"
            xlSheet1.Cells(8, "B") = "Report for the month of :"
            xlSheet1.Cells(8, "B").Font.Bold = True
            xlSheet1.Cells(8, "D") = cboMonthly_Month & " " & txtMonthly_Year
            xlSheet1.Cells(8, "D").Font.Bold = True
            xlSheet1.Range("D8:G8").Borders(xlEdgeBottom).Weight = xlThin

            xlSheet1.Cells(9, "A") = "III"
            xlSheet1.Cells(9, "B") = "SC (Complete Name) :"
            xlSheet1.Cells(9, "B").Font.Bold = True
            xlSheet1.Cells(9, "D") = Null2String(rsSOLO_SAE!NIYM)
            xlSheet1.Cells(9, "D").Font.Bold = True
            xlSheet1.Range("D9:G9").Borders(xlEdgeBottom).Weight = xlThin

            xlSheet1.Cells(10, "A") = "IV"
            xlSheet1.Cells(10, "B") = "DATE HIRED"
            xlSheet1.Cells(10, "B").Font.Bold = True
            xlSheet1.Cells(10, "C") = X_HIRED
            xlSheet1.Cells(10, "C").Font.Bold = True
            xlSheet1.Range("D10:G10").Borders(xlEdgeBottom).Weight = xlThin

            xlSheet1.Cells(11, "A") = "V"
            xlSheet1.Cells(11, "B") = "Date Resigned :"
            xlSheet1.Cells(11, "B").Font.Bold = True
            xlSheet1.Cells(11, "C") = X_RESIGNED
            xlSheet1.Cells(11, "C").Font.Bold = True
            xlSheet1.Range("D11:G11").Borders(xlEdgeBottom).Weight = xlThin


            xlSheet1.Cells(12, "A") = "VI"
            xlSheet1.Cells(12, "B") = "TIN #:"
            xlSheet1.Cells(12, "D") = X_TIN
            xlSheet1.Cells(12, "B").Font.Bold = True
            xlSheet1.Range("D10:G10").Borders(xlEdgeBottom).Weight = xlThin

            xlSheet1.Cells(13, "A") = "VII"
            xlSheet1.Cells(13, "B") = "Contract (Salaried / Commissioned) :"
            xlSheet1.Cells(13, "B").Font.Bold = True
            xlSheet1.Cells(13, "D") = X_CONTRACT

            xlSheet1.Cells(14, "A") = "VIII"
            xlSheet1.Cells(14, "B") = "HARI Certified (Yes / No) : "
            xlSheet1.Cells(14, "B").Font.Bold = True
            xlSheet1.Cells(14, "D") = Y_N
            xlSheet1.Cells(14, "D").Font.Bold = True

            xlSheet1.Cells(14, "E") = "BATCH #:"
            xlSheet1.Cells(14, "E").Font.Bold = True
            'CODE FOR BATCH APPEAR HERE

            'xlSheet1.Range("D" & XXX & ":" & "O" & XXX).Cells.Merge
            xlSheet1.Range("B17:P17").MergeCells = True
            xlSheet1.Range("B17:P17") = "PRODUCTIVITY"
            xlSheet1.Range("B17:P17").Cells.HorizontalAlignment = xlCenter
            'xlSheet1.Range("B17:Q17").Interior.Color = &HFF0000
            xlSheet1.Range("B17:P17").Font.Bold = True
            
            XXX = XXX + 1

            xlSheet1.Cells(XXX, "A") = "IX"
            xlSheet1.Cells(XXX, "B") = "MODEL"
            xlSheet1.Cells(XXX, "D") = "JAN"
            xlSheet1.Cells(XXX, "E") = "FEB"
            xlSheet1.Cells(XXX, "F") = "MAR"
            xlSheet1.Cells(XXX, "G") = "APR"
            xlSheet1.Cells(XXX, "H") = "MAY"
            xlSheet1.Cells(XXX, "I") = "JUN"
            xlSheet1.Cells(XXX, "J") = "JUL"
            xlSheet1.Cells(XXX, "K") = "AUG"
            xlSheet1.Cells(XXX, "L") = "SEP"
            xlSheet1.Cells(XXX, "M") = "OCT"
            xlSheet1.Cells(XXX, "N") = "NOV"
            xlSheet1.Cells(XXX, "O") = "DEC"
            xlSheet1.Cells(XXX, "P") = "TOTAL"
            
            xlSheet1.Cells(XXX, "P").Interior.Color = &H80FFFF
            xlSheet1.Cells(XXX, "P").Borders(xlEdgeTop).Weight = xlThick
            xlSheet1.Cells(XXX, "P").Borders(xlEdgeLeft).Weight = xlThick
            xlSheet1.Cells(XXX, "P").Borders(xlEdgeRight).Weight = xlThick
    
            
            xlSheet1.Range("B" & XXX & ":" & "O" & XXX).Interior.Color = vbBlue
            xlSheet1.Range("B" & XXX & ":" & "P" & XXX).Font.Bold = True
            xlSheet1.Range("B" & XXX & ":" & "P" & XXX).Cells.HorizontalAlignment = xlCenter
            xlSheet1.Range("B" & XXX & ":" & "O" & XXX).BorderAround ColorIndex:=1, Weight:=xlThin
            XXX = XXX + 1
                                  
            Set rsModel = gconDMIS.Execute("Select DISTINCT MODEL as vehModel FROM ALL_MODEL WHERE MODEL IS NOT NULL ORDER BY MODEL ASC")
                If Not rsModel.EOF And Not rsModel.BOF Then
                    Do While Not rsModel.EOF
                      xlSheet1.Range("B" & XXX & ":" & "C" & XXX).MergeCells = True
                      xlSheet1.Cells(XXX, "B") = Null2String(rsModel!vehModel)
                      xlSheet1.Range("B" & XXX & ":" & "C" & XXX).BorderAround ColorIndex:=1, Weight:=xlThin
                        For x_MONTH = 1 To 12
                        lblLoading.Caption = "Loading...."
                            Set rsCOUNT_SAE_RELEASE = gconDMIS.Execute("Select COUNT(DATERELEASED) as SOLDOUT from SMIS_PURCHAGREE WHERE month(DATERELEASED) = " & x_MONTH & " AND YEAR(DATERELEASED) = '" & txtMonthly_Year & "' and STATUS = 'P' and SALESAE = '" & rsSOLO_SAE!NIYM & "' and Model = '" & rsModel!vehModel & "'")
                            If x_MONTH = 1 Then
                                xlSheet1.Cells(XXX, "D") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "D").Cells.HorizontalAlignment = xlCenter
                            ElseIf x_MONTH = 2 Then
                                xlSheet1.Cells(XXX, "E") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "E").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "E").Cells.HorizontalAlignment = xlCenter
                            ElseIf x_MONTH = 3 Then
                                xlSheet1.Cells(XXX, "F") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "F").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "F").Cells.HorizontalAlignment = xlCenter
                            ElseIf x_MONTH = 4 Then
                                xlSheet1.Cells(XXX, "G") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "G").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "G").Cells.HorizontalAlignment = xlCenter
                            ElseIf x_MONTH = 5 Then
                                xlSheet1.Cells(XXX, "H") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "H").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "H").Cells.HorizontalAlignment = xlCenter
                            ElseIf x_MONTH = 6 Then
                                xlSheet1.Cells(XXX, "I") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "I").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "I").Cells.HorizontalAlignment = xlCenter
                            ElseIf x_MONTH = 7 Then
                                xlSheet1.Cells(XXX, "J") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "J").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "J").Cells.HorizontalAlignment = xlCenter
                            ElseIf x_MONTH = 8 Then
                                xlSheet1.Cells(XXX, "K") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "K").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "K").Cells.HorizontalAlignment = xlCenter
                            ElseIf x_MONTH = 9 Then
                                xlSheet1.Cells(XXX, "L") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "L").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "L").Cells.HorizontalAlignment = xlCenter
                            ElseIf x_MONTH = 10 Then
                                xlSheet1.Cells(XXX, "M") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "M").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "M").Cells.HorizontalAlignment = xlCenter
                            ElseIf x_MONTH = 11 Then
                                xlSheet1.Cells(XXX, "N") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "N").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "N").Cells.HorizontalAlignment = xlCenter
                            ElseIf x_MONTH = 12 Then
                                xlSheet1.Cells(XXX, "O") = Null2String(rsCOUNT_SAE_RELEASE!SOLDOUT)
                                xlSheet1.Cells(XXX, "O").BorderAround ColorIndex:=1, Weight:=xlThin
                                xlSheet1.Cells(XXX, "O").Cells.HorizontalAlignment = xlCenter
                            End If
                            lblLoading.Caption = "Loading........"
                        Next x_MONTH
                       
                       Set rsSUM_MODEL_PER_YEAR = gconDMIS.Execute("Select COUNT(DATERELEASED) as SOLD_PER_YEAR from SMIS_PURCHAGREE where SALESAE = '" & rsSOLO_SAE!NIYM & "'  and YEAR(DATERELEASED) = '" & txtMonthly_Year & "' and model = '" & rsModel!vehModel & "' and STATUS = 'P'")
                       If Not rsSUM_MODEL_PER_YEAR.EOF And Not rsSUM_MODEL_PER_YEAR.BOF Then
                            xlSheet1.Cells(XXX, "P") = Null2String(rsSUM_MODEL_PER_YEAR!SOLD_PER_YEAR)
                            'xlSheet1.Cells(XXX, "p").BorderAround ColorIndex:=1, Weight:=xlThin
                            xlSheet1.Cells(XXX, "P").Borders(xlEdgeTop).Weight = xlThin
                            xlSheet1.Cells(XXX, "P").Borders(xlEdgeBottom).Weight = xlThin
                            xlSheet1.Cells(XXX, "P").Borders(xlEdgeRight).Weight = xlThick
                            xlSheet1.Cells(XXX, "P").Borders(xlEdgeLeft).Weight = xlThick
                            xlSheet1.Cells(XXX, "p").Cells.HorizontalAlignment = xlCenter
                            xlSheet1.Cells(XXX, "p").Interior.Color = &H80FFFF
                       End If
                       
                       XXX = XXX + 1
                       rsModel.MoveNext
                    Loop
                End If
                'TOTAL PER MONTH OF ALL MODEL
                        Dim rsSUM_PER_MONTH     As ADODB.Recordset
                        Dim x_Bilang            As Integer
                        
                        xlSheet1.Range("B" & XXX & ":" & "C" & XXX).MergeCells = True
                        xlSheet1.Range("B" & XXX & ":" & "C" & XXX) = "TOTAL"
                        xlSheet1.Range("B" & XXX & ":" & "C" & XXX).Interior.Color = &H80FFFF
                        xlSheet1.Range("B" & XXX & ":" & "C" & XXX).Font.Bold = True
                        
                        For x_Bilang = 1 To 12
                        lblLoading.Caption = "Loading...."
                            Set rsSUM_PER_MONTH = gconDMIS.Execute("Select count(DATERELEASED) as SUM_MODEL_PERMONHT from SMIS_PURCHAGREE WHERE MONTH(DATERELEASED) = " & x_Bilang & "  AND YEAR(DATERELEASED) = '" & txtMonthly_Year & "' AND STATUS = 'P' and SALESAE = '" & rsSOLO_SAE!NIYM & "'")
                                If Not rsSUM_PER_MONTH.EOF And Not rsSUM_PER_MONTH.BOF Then
                                    If x_Bilang = 1 Then
                                        xlSheet1.Cells(XXX, "D") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "D").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "D").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "D").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "D").Font.Bold = True
                                    ElseIf x_Bilang = 2 Then
                                        xlSheet1.Cells(XXX, "E") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "E").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "E").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "E").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "E").Font.Bold = True
                                    ElseIf x_Bilang = 3 Then
                                        xlSheet1.Cells(XXX, "F") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "F").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "F").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "F").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "F").Font.Bold = True
                                    ElseIf x_Bilang = 4 Then
                                        xlSheet1.Cells(XXX, "G") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "G").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "G").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "G").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "G").Font.Bold = True
                                    ElseIf x_Bilang = 5 Then
                                        xlSheet1.Cells(XXX, "H") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "H").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "H").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "H").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "H").Font.Bold = True
                                    ElseIf x_Bilang = 6 Then
                                        xlSheet1.Cells(XXX, "I") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "I").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "I").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "I").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "I").Font.Bold = True
                                    ElseIf x_Bilang = 7 Then
                                        xlSheet1.Cells(XXX, "J") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "J").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "J").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "J").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "J").Font.Bold = True
                                    ElseIf x_Bilang = 8 Then
                                        xlSheet1.Cells(XXX, "K") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "K").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "K").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "K").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "K").Font.Bold = True
                                    ElseIf x_Bilang = 9 Then
                                        xlSheet1.Cells(XXX, "L") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "L").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "L").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "L").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "L").Font.Bold = True
                                    ElseIf x_Bilang = 10 Then
                                        xlSheet1.Cells(XXX, "M") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "M").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "M").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "M").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "M").Font.Bold = True
                                    ElseIf x_Bilang = 11 Then
                                        xlSheet1.Cells(XXX, "N") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "N").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "N").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "N").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "N").Font.Bold = True
                                    ElseIf x_Bilang = 12 Then
                                        xlSheet1.Cells(XXX, "O") = Null2String(rsSUM_PER_MONTH!SUM_MODEL_PERMONHT)
                                        xlSheet1.Cells(XXX, "O").Cells.HorizontalAlignment = xlCenter
                                        xlSheet1.Cells(XXX, "O").Borders(xlEdgeLeft).Weight = xlThin
                                        xlSheet1.Cells(XXX, "O").Interior.Color = &H80FFFF
                                        xlSheet1.Cells(XXX, "O").Font.Bold = True
                                    End If
                                End If
                        lblLoading.Caption = "Loading........"
                        Next x_Bilang
                        
                xlSheet1.Range("B" & XXX & ":" & "P" & XXX).BorderAround ColorIndex:=1, Weight:=xlThick
                
                Set rsSUM_WHOLE_YEAR = gconDMIS.Execute("Select COUNT(DATERELEASED) AS SUM_YER_SOLD from SMIS_PURCHAGREE WHERE YEAR(DATERELEASED) = '" & txtMonthly_Year & "' AND STATUS = 'P' and SALESAE = '" & rsSOLO_SAE!NIYM & "'")
                If Not rsSUM_WHOLE_YEAR.EOF And Not rsSUM_WHOLE_YEAR.BOF Then
                    xlSheet1.Cells(XXX, "P") = Null2String(rsSUM_WHOLE_YEAR!SUM_YER_SOLD)
                    xlSheet1.Cells(XXX, "P").Borders(xlEdgeLeft).Weight = xlThick
                    xlSheet1.Cells(XXX, "P").Cells.HorizontalAlignment = xlCenter
                    xlSheet1.Cells(XXX, "P").Interior.Color = &H80FFFF
                    xlSheet1.Cells(XXX, "P").Font.Bold = True
                End If
                                        
                XXX = XXX + 2
                xlSheet1.Cells(XXX, "B") = "Prepared by:"
                xlSheet1.Cells(XXX, "J") = "Approved by:"
                
                XXX = XXX + 2
                xlSheet1.Range("B" & XXX & ":" & "D" & XXX).MergeCells = True
                xlSheet1.Range("B" & XXX & ":" & "D" & XXX).Borders(xlEdgeBottom).Weight = xlThick
                xlSheet1.Range("J" & XXX & ":" & "L" & XXX).MergeCells = True
                xlSheet1.Range("J" & XXX & ":" & "L" & XXX).Borders(xlEdgeBottom).Weight = xlThick
                
                
                XXX = XXX + 1
                xlSheet1.Cells(XXX, "B") = "SALES ADMIN. SUPERVISOR"
                xlSheet1.Cells(XXX, "B").Font.Bold = True
                
                XXX = XXX + 2
                xlSheet1.Cells(XXX, "A") = "NOTE"
                xlSheet1.Cells(XXX, "A").Font.Color = &HFF&
                xlSheet1.Cells(XXX, "A").Font.Bold = True
                
                XXX = XXX + 1
                xlSheet1.Cells(XXX, "A") = "* Fill-up SC Information completely except V (if not yet resigned)"
                xlSheet1.Cells(XXX, "A").Font.Color = &HFF&
                                        
                XXX = XXX + 1
                xlSheet1.Cells(XXX, "B") = " - Submission is per SC  per month"
                xlSheet1.Cells(XXX, "B").Font.Color = &HFF&
                
                XXX = XXX + 1
                xlSheet1.Cells(XXX, "A") = "* Please submit this report including all your SCs even without sale (for monitoring purposes),"
                xlSheet1.Cells(XXX, "A").Font.Color = &HFF&
                
                XXX = XXX + 1
                xlSheet1.Cells(XXX, "A") = "otherwise HARI will consider the SC as a resigned employee"
                xlSheet1.Cells(XXX, "A").Font.Color = &HFF&
                
                XXX = XXX + 1
                xlSheet1.Cells(XXX, "A") = "* To be submitted on the 2nd day of the following month."
                xlSheet1.Cells(XXX, "A").Font.Color = &HFF&
                xlSheet1.Cells(XXX, "A").Font.Bold = True
                
                XXX = XXX + 1
                xlSheet1.Cells(XXX, "A") = "* If it falls on a weekend, submission of the report shall be on the first"
                xlSheet1.Cells(XXX, "A").Font.Color = &HFF&
                
                XXX = XXX + 1
                xlSheet1.Cells(XXX, "A") = "working day of the succeeding month."
                xlSheet1.Cells(XXX, "A").Font.Color = &HFF&
                
                XXX = XXX + 1
                xlSheet1.Cells(XXX, "A") = "* Fax no. 812-1556/894-5866"
                xlSheet1.Cells(XXX, "A").Font.Color = &HFF&
                
                XXX = XXX + 1
                xlSheet1.Cells(XXX, "A") = "* Email thru gmalabuyoc@hyundai-asia.com"
                xlSheet1.Cells(XXX, "A").Font.Color = &HFF&
                
                XXX = XXX + 1
                xlSheet1.Cells(XXX, "B") = " - cc scalma@hyundai-asia.com"
                xlSheet1.Cells(XXX, "B").Font.Color = &HFF&
                
                XXX = 17
                rsSOLO_SAE.MoveNext
        Loop
    End If
    
    Set rsSUM_WHOLE_YEAR = Nothing
    Set rsSUM_PER_MONTH = Nothing
    Set rsSOLO_SAE = Nothing
    Set rsCOUNT_SAE_RELEASE = Nothing
    Set rsModel = Nothing
    Set rsSAE_CODE = Nothing
    Set rsSAE_INFO = Nothing
    Set rsSUM_MODEL_PER_YEAR = Nothing
End Sub
Sub COUNT_SC_SALARIED()
    Dim rsCOUNT_SALARIED          As ADODB.Recordset
    Dim rsSC_CODE                 As ADODB.Recordset
    Dim rsSALESTEAM               As ADODB.Recordset
    
    Dim x_MONTH_COUNT             As Integer
    Dim x_SALARIED                As Integer
    
    For x_MONTH_COUNT = 1 To 12
        lblLoading.Caption = "Loading...."
        Set rsCOUNT_SALARIED = gconDMIS.Execute("Select DISTINCT SALESAE FROM SMIS_PURCHAGREE WHERE MONTH(DATERELEASED) = " & x_MONTH_COUNT & " AND YEAR(DATERELEASED) = '" & txtMonthly_Year & "' and status = 'P' ORDER BY SALESAE ASC")
        If Not rsCOUNT_SALARIED.EOF And Not rsCOUNT_SALARIED.BOF Then
            Do While Not rsCOUNT_SALARIED.EOF
                Set rsSC_CODE = gconDMIS.Execute("Select SAECODE from SMIS_vw_Srep where NAME = '" & rsCOUNT_SALARIED!salesae & "'")
                If Not rsSC_CODE.EOF And Not rsSC_CODE.BOF Then
                    Do While Not rsSC_CODE.EOF
                        Set rsSALESTEAM = gconDMIS.Execute("Select * from SMIS_SALESTEAM WHERE SAECODE = '" & rsSC_CODE!SAECODE & "' AND CONTRACT = 'S'")
                        If Not rsSALESTEAM.EOF And Not rsSALESTEAM.BOF Then
                            x_SALARIED = x_SALARIED + 1
                        End If
                        rsSC_CODE.MoveNext
                    Loop
                End If
                rsCOUNT_SALARIED.MoveNext
            Loop
        End If
        
        If x_MONTH_COUNT = 1 Then
            xlSheet1.Cells(xKawnt, "D") = x_SALARIED
            xlSheet1.Cells(xKawnt, "D").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "D").Cells.HorizontalAlignment = xlCenter
        ElseIf x_MONTH_COUNT = 2 Then
            xlSheet1.Cells(xKawnt, "E") = x_SALARIED
            xlSheet1.Cells(xKawnt, "E").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "E").Cells.HorizontalAlignment = xlCenter
        ElseIf x_MONTH_COUNT = 3 Then
            xlSheet1.Cells(xKawnt, "F") = x_SALARIED
            xlSheet1.Cells(xKawnt, "F").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "F").Cells.HorizontalAlignment = xlCenter
        ElseIf x_MONTH_COUNT = 4 Then
            xlSheet1.Cells(xKawnt, "G") = x_SALARIED
            xlSheet1.Cells(xKawnt, "G").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "G").Cells.HorizontalAlignment = xlCenter
        ElseIf x_MONTH_COUNT = 5 Then
            xlSheet1.Cells(xKawnt, "H") = x_SALARIED
            xlSheet1.Cells(xKawnt, "H").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "H").Cells.HorizontalAlignment = xlCenter
        ElseIf x_MONTH_COUNT = 6 Then
            xlSheet1.Cells(xKawnt, "I") = x_SALARIED
            xlSheet1.Cells(xKawnt, "I").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "I").Cells.HorizontalAlignment = xlCenter
        ElseIf x_MONTH_COUNT = 7 Then
            xlSheet1.Cells(xKawnt, "J") = x_SALARIED
            xlSheet1.Cells(xKawnt, "J").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "J").Cells.HorizontalAlignment = xlCenter
        ElseIf x_MONTH_COUNT = 8 Then
            xlSheet1.Cells(xKawnt, "K") = x_SALARIED
            xlSheet1.Cells(xKawnt, "K").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "K").Cells.HorizontalAlignment = xlCenter
        ElseIf x_MONTH_COUNT = 9 Then
            xlSheet1.Cells(xKawnt, "L") = x_SALARIED
            xlSheet1.Cells(xKawnt, "L").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "L").Cells.HorizontalAlignment = xlCenter
        ElseIf x_MONTH_COUNT = 10 Then
            xlSheet1.Cells(xKawnt, "M") = x_SALARIED
            xlSheet1.Cells(xKawnt, "M").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "M").Cells.HorizontalAlignment = xlCenter
        ElseIf x_MONTH_COUNT = 11 Then
            xlSheet1.Cells(xKawnt, "N") = x_SALARIED
            xlSheet1.Cells(xKawnt, "N").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "N").Cells.HorizontalAlignment = xlCenter
        ElseIf x_MONTH_COUNT = 12 Then
            xlSheet1.Cells(xKawnt, "O") = x_SALARIED
            xlSheet1.Cells(xKawnt, "O").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "O").Cells.HorizontalAlignment = xlCenter
        End If
        
        'RESET COUNTER
        x_SALARIED = 0
        lblLoading.Caption = "Loading........"
    Next x_MONTH_COUNT
    
    Set rsCOUNT_SALARIED = Nothing
    Set rsSC_CODE = Nothing
    Set rsSALESTEAM = Nothing

End Sub

Sub COUNT_SC_COMMISSION()
    Dim rsCOMMISSION              As ADODB.Recordset
    Dim rsSC_CODE                 As ADODB.Recordset
    Dim rsSALESTEAM               As ADODB.Recordset
    
    Dim x_COUNT                   As Integer
    Dim x_COMMISION               As Integer
    
    For x_COUNT = 1 To 12
    lblLoading.Caption = "Loading...."
        Set rsCOMMISSION = gconDMIS.Execute("Select DISTINCT SALESAE FROM SMIS_PURCHAGREE WHERE MONTH(DATERELEASED) = " & x_COUNT & " AND YEAR(DATERELEASED) = '" & txtMonthly_Year & "' and status = 'P' ORDER BY SALESAE ASC")
        If Not rsCOMMISSION.EOF And Not rsCOMMISSION.BOF Then
            Do While Not rsCOMMISSION.EOF
                Set rsSC_CODE = gconDMIS.Execute("Select SAECODE from SMIS_vw_Srep where NAME = '" & rsCOMMISSION!salesae & "'")
                If Not rsSC_CODE.EOF And Not rsSC_CODE.BOF Then
                    Do While Not rsSC_CODE.EOF
                        Set rsSALESTEAM = gconDMIS.Execute("Select * from SMIS_SALESTEAM WHERE SAECODE = '" & rsSC_CODE!SAECODE & "' AND (CONTRACT = 'C' OR CONTRACT IS NULL)")
                        If Not rsSALESTEAM.EOF And Not rsSALESTEAM.BOF Then
                            x_COMMISION = x_COMMISION + 1
                        End If
                        rsSC_CODE.MoveNext
                    Loop
                End If
                rsCOMMISSION.MoveNext
            Loop
        End If
        
        If x_COUNT = 1 Then
            xlSheet1.Cells(xKawnt, "D") = x_COMMISION
            xlSheet1.Cells(xKawnt, "D").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "D").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 2 Then
            xlSheet1.Cells(xKawnt, "E") = x_COMMISION
            xlSheet1.Cells(xKawnt, "E").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "E").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 3 Then
            xlSheet1.Cells(xKawnt, "F") = x_COMMISION
            xlSheet1.Cells(xKawnt, "F").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "F").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 4 Then
            xlSheet1.Cells(xKawnt, "G") = x_COMMISION
            xlSheet1.Cells(xKawnt, "G").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "G").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 5 Then
            xlSheet1.Cells(xKawnt, "H") = x_COMMISION
            xlSheet1.Cells(xKawnt, "H").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "H").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 6 Then
            xlSheet1.Cells(xKawnt, "I") = x_COMMISION
            xlSheet1.Cells(xKawnt, "I").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "I").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 7 Then
            xlSheet1.Cells(xKawnt, "J") = x_COMMISION
            xlSheet1.Cells(xKawnt, "J").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "J").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 8 Then
            xlSheet1.Cells(xKawnt, "K") = x_COMMISION
            xlSheet1.Cells(xKawnt, "K").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "K").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 9 Then
            xlSheet1.Cells(xKawnt, "L") = x_COMMISION
            xlSheet1.Cells(xKawnt, "L").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "L").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 10 Then
            xlSheet1.Cells(xKawnt, "M") = x_COMMISION
            xlSheet1.Cells(xKawnt, "M").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "M").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 11 Then
            xlSheet1.Cells(xKawnt, "N") = x_COMMISION
            xlSheet1.Cells(xKawnt, "N").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "N").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 12 Then
            xlSheet1.Cells(xKawnt, "O") = x_COMMISION
            xlSheet1.Cells(xKawnt, "O").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "O").Cells.HorizontalAlignment = xlCenter
        End If
        
        'RESET COUNTER
        x_COMMISION = 0
        lblLoading.Caption = "Loading........"
    Next x_COUNT
    
    Set rsCOMMISSION = Nothing
    Set rsSC_CODE = Nothing
    Set rsSALESTEAM = Nothing
End Sub

Sub COUNT_SC_CERTIFIED_PER_MONTH()
    Dim rsCOUNT_SC As ADODB.Recordset
    Dim rsSC_CODE As ADODB.Recordset
    Dim rsSALESTEAM As ADODB.Recordset
    Dim x_COUNT As Integer
    
    For x_COUNT = 1 To 12
    lblLoading.Caption = "Loading...."
        Set rsCOUNT_SC = gconDMIS.Execute("Select DISTINCT SALESAE FROM SMIS_PURCHAGREE WHERE MONTH(DATERELEASED) = " & x_COUNT & " AND YEAR(DATERELEASED) = '" & txtMonthly_Year & "' and status = 'P' ORDER BY SALESAE ASC")
        If Not rsCOUNT_SC.EOF And Not rsCOUNT_SC.BOF Then
            Do While Not rsCOUNT_SC.EOF
                Set rsSC_CODE = gconDMIS.Execute("Select SAECODE from SMIS_vw_Srep where NAME = '" & rsCOUNT_SC!salesae & "'")
                If Not rsSC_CODE.EOF And Not rsSC_CODE.BOF Then
                    Do While Not rsSC_CODE.EOF
                        Set rsSALESTEAM = gconDMIS.Execute("Select * from SMIS_SALESTEAM WHERE SAECODE = '" & rsSC_CODE!SAECODE & "' AND HARI_CERTIFIED = 'Y'")
                        If Not rsSALESTEAM.EOF And Not rsSALESTEAM.BOF Then
                            HARI_Certified = HARI_Certified + 1
                        End If
                        rsSC_CODE.MoveNext
                    Loop
                End If
                rsCOUNT_SC.MoveNext
            Loop
        End If
        
        If x_COUNT = 1 Then
            xlSheet1.Cells(xKawnt, "D") = HARI_Certified
            xlSheet1.Cells(xKawnt, "D").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "D").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 2 Then
            xlSheet1.Cells(xKawnt, "E") = HARI_Certified
            xlSheet1.Cells(xKawnt, "E").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "E").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 3 Then
            xlSheet1.Cells(xKawnt, "F") = HARI_Certified
            xlSheet1.Cells(xKawnt, "F").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "F").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 4 Then
            xlSheet1.Cells(xKawnt, "G") = HARI_Certified
            xlSheet1.Cells(xKawnt, "G").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "G").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 5 Then
            xlSheet1.Cells(xKawnt, "H") = HARI_Certified
            xlSheet1.Cells(xKawnt, "H").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "H").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 6 Then
            xlSheet1.Cells(xKawnt, "I") = HARI_Certified
            xlSheet1.Cells(xKawnt, "I").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "I").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 7 Then
            xlSheet1.Cells(xKawnt, "J") = HARI_Certified
            xlSheet1.Cells(xKawnt, "J").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "J").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 8 Then
            xlSheet1.Cells(xKawnt, "K") = HARI_Certified
            xlSheet1.Cells(xKawnt, "K").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "K").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 9 Then
            xlSheet1.Cells(xKawnt, "L") = HARI_Certified
            xlSheet1.Cells(xKawnt, "L").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "L").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 10 Then
            xlSheet1.Cells(xKawnt, "M") = HARI_Certified
            xlSheet1.Cells(xKawnt, "M").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "M").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 11 Then
            xlSheet1.Cells(xKawnt, "N") = HARI_Certified
            xlSheet1.Cells(xKawnt, "N").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "N").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 12 Then
            xlSheet1.Cells(xKawnt, "O") = HARI_Certified
            xlSheet1.Cells(xKawnt, "O").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "O").Cells.HorizontalAlignment = xlCenter
        End If
        
        'RESET COUNTER
        HARI_Certified = 0
    lblLoading.Caption = "Loading........"
    Next x_COUNT
    
    Set rsCOUNT_SC = Nothing
    Set rsSC_CODE = Nothing
End Sub

Sub COUNT_SC_NON_CERTIFIED_PER_MONTH()
    Dim rsCOUNT_SC As ADODB.Recordset
    Dim rsSC_CODE As ADODB.Recordset
    Dim rsSALESTEAM As ADODB.Recordset
    Dim x_COUNT As Integer
    
    For x_COUNT = 1 To 12
    lblLoading.Caption = "Loading...."
'        Set rsCOUNT_SC = gconDMIS.Execute("Select DISTINCT SALESAE FROM SMIS_PURCHAGREE WHERE MONTH(DATERELEASED) = " & x_COUNT & " AND YEAR(DATERELEASED) = '" & txtMonthly_Year & "' and status <> 'C' ORDER BY SALESAE ASC")
'        If Not rsCOUNT_SC.EOF And Not rsCOUNT_SC.BOF Then
'            Do While Not rsCOUNT_SC.EOF
'                Set rsSC_CODE = gconDMIS.Execute("Select DISTINCT SAECODE from SMIS_vw_Srep where NAME = '" & rsCOUNT_SC!salesae & "'")
'                If Not rsSC_CODE.EOF And Not rsSC_CODE.BOF Then
'                    Do While Not rsSC_CODE.EOF
'                        Set rsSALESTEAM = gconDMIS.Execute("Select * from SMIS_SALESTEAM WHERE SAECODE = '" & rsSC_CODE!SAECODE & "' AND (HARI_CERTIFIED = 'N' or HARI_CERTIFIED IS NULL)")
'                        If Not rsSALESTEAM.EOF And Not rsSALESTEAM.BOF Then
'                            NON_HARI_CERTIFIED = NON_HARI_CERTIFIED + 1
'                        End If
'                        rsSC_CODE.MoveNext
'                    Loop
'                End If
'                rsCOUNT_SC.MoveNext
'            Loop
'        End If
        
        
        Set rsCOUNT_SC = gconDMIS.Execute("Select count(DISTINCT SALESAE) as NON_HARI from SMIS_purchagree where SALESAE IN (Select [name] from SMIS_vw_Srep where SAECODE IN (Select SAECODE from SMIS_SALESTEAM WHERE HARI_CERTIFIED = 'N' or HARI_CERTIFIED IS NULL)) and month(datereleased) = '1' and year(datereleased) = '2009' and STATUS <> 'C'")
        
        NON_HARI_CERTIFIED = rsCOUNT_SC!NON_HARI
        
        If x_COUNT = 1 Then
            xlSheet1.Cells(xKawnt, "D") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "D").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "D").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 2 Then
            xlSheet1.Cells(xKawnt, "E") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "E").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "E").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 3 Then
            xlSheet1.Cells(xKawnt, "F") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "F").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "F").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 4 Then
            xlSheet1.Cells(xKawnt, "G") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "G").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "G").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 5 Then
            xlSheet1.Cells(xKawnt, "H") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "H").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "H").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 6 Then
            xlSheet1.Cells(xKawnt, "I") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "I").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "I").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 7 Then
            xlSheet1.Cells(xKawnt, "J") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "J").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "J").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 8 Then
            xlSheet1.Cells(xKawnt, "K") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "K").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "K").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 9 Then
            xlSheet1.Cells(xKawnt, "L") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "L").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "L").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 10 Then
            xlSheet1.Cells(xKawnt, "M") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "M").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "M").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 11 Then
            xlSheet1.Cells(xKawnt, "N") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "N").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "N").Cells.HorizontalAlignment = xlCenter
        ElseIf x_COUNT = 12 Then
            xlSheet1.Cells(xKawnt, "O") = NON_HARI_CERTIFIED
            xlSheet1.Cells(xKawnt, "O").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(xKawnt, "O").Cells.HorizontalAlignment = xlCenter
        End If
        
        'RESET COUNTER
        NON_HARI_CERTIFIED = 0
    lblLoading.Caption = "Loading........"
    Next x_COUNT
    
    Set rsCOUNT_SC = Nothing
    Set rsSC_CODE = Nothing
End Sub
Sub COUNT_TOTAL_SC_PER_MONTH()
    Dim rsSC             As ADODB.Recordset
    Dim JUN_BORDADO      As Integer
         
        For JUN_BORDADO = 1 To 12
        lblLoading.Caption = "Loading...."
            Set rsSC = gconDMIS.Execute("Select COUNT(DISTINCT SALESAE)as SC_COUNT from SMIS_PURCHAGREE where month(Datereleased) = " & JUN_BORDADO & " and year(Datereleased) = '" & txtMonthly_Year & "' AND status = 'P'")
            If JUN_BORDADO = 1 Then
                xlSheet1.Cells(xKawnt, "D") = rsSC!SC_COUNT
                'xlSheet1.Cells(xKawnt, "D").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "D").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "D").Cells.HorizontalAlignment = xlCenter
            ElseIf JUN_BORDADO = 2 Then
                xlSheet1.Cells(xKawnt, "E") = rsSC!SC_COUNT
                xlSheet1.Cells(xKawnt, "E").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "E").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "E").Cells.HorizontalAlignment = xlCenter
            ElseIf JUN_BORDADO = 3 Then
                xlSheet1.Cells(xKawnt, "F") = rsSC!SC_COUNT
                xlSheet1.Cells(xKawnt, "F").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "F").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "F").Cells.HorizontalAlignment = xlCenter
            ElseIf JUN_BORDADO = 4 Then
                xlSheet1.Cells(xKawnt, "G") = rsSC!SC_COUNT
                xlSheet1.Cells(xKawnt, "G").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "G").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "G").Cells.HorizontalAlignment = xlCenter
            ElseIf JUN_BORDADO = 5 Then
                xlSheet1.Cells(xKawnt, "H") = rsSC!SC_COUNT
                xlSheet1.Cells(xKawnt, "H").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "H").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "H").Cells.HorizontalAlignment = xlCenter
            ElseIf JUN_BORDADO = 6 Then
                xlSheet1.Cells(xKawnt, "I") = rsSC!SC_COUNT
                xlSheet1.Cells(xKawnt, "I").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "I").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "I").Cells.HorizontalAlignment = xlCenter
            ElseIf JUN_BORDADO = 7 Then
                xlSheet1.Cells(xKawnt, "J") = rsSC!SC_COUNT
                xlSheet1.Cells(xKawnt, "J").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "J").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "J").Cells.HorizontalAlignment = xlCenter
            ElseIf JUN_BORDADO = 8 Then
                xlSheet1.Cells(xKawnt, "K") = rsSC!SC_COUNT
                xlSheet1.Cells(xKawnt, "K").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "K").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "K").Cells.HorizontalAlignment = xlCenter
            ElseIf JUN_BORDADO = 9 Then
                xlSheet1.Cells(xKawnt, "L") = rsSC!SC_COUNT
                xlSheet1.Cells(xKawnt, "L").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "L").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "L").Cells.HorizontalAlignment = xlCenter
            ElseIf JUN_BORDADO = 10 Then
                xlSheet1.Cells(xKawnt, "M") = rsSC!SC_COUNT
                xlSheet1.Cells(xKawnt, "M").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "M").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "M").Cells.HorizontalAlignment = xlCenter
            ElseIf JUN_BORDADO = 11 Then
                xlSheet1.Cells(xKawnt, "N") = rsSC!SC_COUNT
                xlSheet1.Cells(xKawnt, "N").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "N").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "N").Cells.HorizontalAlignment = xlCenter
                
            ElseIf JUN_BORDADO = 12 Then
                xlSheet1.Cells(xKawnt, "O") = rsSC!SC_COUNT
                xlSheet1.Cells(xKawnt, "O").Borders(xlEdgeBottom).Weight = xlThin
                xlSheet1.Cells(xKawnt, "O").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "O").Cells.HorizontalAlignment = xlCenter
            End If
        lblLoading.Caption = "Loading........"
        Next JUN_BORDADO
    
    Set rsSC = Nothing
End Sub

Sub COUNT_SC_RESIGNED_THIS_MONTH()
    Dim rsRESIGNED          As ADODB.Recordset
    Dim X_TANGGAL           As Integer
        For X_TANGGAL = 1 To 12
        lblLoading.Caption = "Loading...."
            Set rsRESIGNED = gconDMIS.Execute("Select COUNT(SAECODE) AS COUNT_RESIGNED from SMIS_SalesTeam where month(DATE_RESIGNED) =" & X_TANGGAL & " AND Year(DATE_RESIGNED) ='" & txtMonthly_Year & "'")
            If X_TANGGAL = 1 Then
                xlSheet1.Cells(xKawnt, "D") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "D").Cells.HorizontalAlignment = xlCenter
            ElseIf X_TANGGAL = 2 Then
                xlSheet1.Cells(xKawnt, "E") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "E").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "E").Cells.HorizontalAlignment = xlCenter
            ElseIf X_TANGGAL = 3 Then
                xlSheet1.Cells(xKawnt, "F") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "F").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "F").Cells.HorizontalAlignment = xlCenter
            ElseIf X_TANGGAL = 4 Then
                xlSheet1.Cells(xKawnt, "G") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "G").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "G").Cells.HorizontalAlignment = xlCenter
            ElseIf X_TANGGAL = 5 Then
                xlSheet1.Cells(xKawnt, "H") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "H").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "H").Cells.HorizontalAlignment = xlCenter
            ElseIf X_TANGGAL = 6 Then
                xlSheet1.Cells(xKawnt, "I") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "I").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "I").Cells.HorizontalAlignment = xlCenter
            ElseIf X_TANGGAL = 7 Then
                xlSheet1.Cells(xKawnt, "J") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "J").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "J").Cells.HorizontalAlignment = xlCenter
            ElseIf X_TANGGAL = 8 Then
                xlSheet1.Cells(xKawnt, "K") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "K").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "K").Cells.HorizontalAlignment = xlCenter
            ElseIf X_TANGGAL = 9 Then
                xlSheet1.Cells(xKawnt, "L") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "L").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "L").Cells.HorizontalAlignment = xlCenter
            ElseIf X_TANGGAL = 10 Then
                xlSheet1.Cells(xKawnt, "M") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "M").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "M").Cells.HorizontalAlignment = xlCenter
            ElseIf X_TANGGAL = 11 Then
                xlSheet1.Cells(xKawnt, "N") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "N").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "N").Cells.HorizontalAlignment = xlCenter
            ElseIf X_TANGGAL = 12 Then
                xlSheet1.Cells(xKawnt, "O") = rsRESIGNED!COUNT_RESIGNED
                xlSheet1.Cells(xKawnt, "O").Borders(xlEdgeLeft).Weight = xlThin
                xlSheet1.Cells(xKawnt, "O").Cells.HorizontalAlignment = xlCenter
            End If
        lblLoading.Caption = "Loading........"
        Next X_TANGGAL
    Set rsRESIGNED = Nothing
End Sub

'COMMENTED BY: JUN - 02-15-2008 --------------THIS IS AN OLD REPORT PRINTING---------------------------
'Sub PRINT_DEALER_SALES_CONSULTANTS_PERFORMANCE_SHEET()
'    Dim SQL                                            As String
'    Dim rsCust                                         As ADODB.Recordset
'
'
'    If Len(Dir(SMIS_REPORT_PATH & "SMIS_EXCEL\SC Perf.xlt")) = 0 Then
'        MsgBox "Excel Directory For Sales Managment Information Could Not be Located", vbInformation
'        Exit Sub
'    End If
'
'
'    Set rsCust = gconDMIS.Execute("SELECT CUSTNAME, HOMEADDRESS,  MODELDESCRIPTION,FRAMENO,ENGINENO, DATERELEASED, TERM,SALESAE  FROM SMIS_SALESORDER WHERE month(INVOICEDDATE)=" & What_month(cboMonthly_Month) & " AND year(INVOICEDDATE)= " & (txtMonthly_Year) & "  AND STATUS<>'C' ORDER BY INVOICEDDATE")
'
'    Dim emprs                                          As ADODB.Recordset
'    Set emprs = gconDMIS.Execute(" SELECT SMIS_vw_Srep.NAME,TINNO, DATEHIRED  FROM SMIS_VW_SREP INNER JOIN HRMS_EMPINFO ON HRMS_EMPINFO.EMPNO=SMIS_VW_SREP.SAECODE ")
'
'
'
'    Dim xlApp
'    Dim xlBook
'    Dim xlSheet1
'    Set xlApp = CreateObject("Excel.Application")
'    Set xlBook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "SMIS_EXCEL\DEALER SALES CONSULTANTS PERFORMANCE SHEET.xlt")
'    Set xlSheet1 = xlBook.Worksheets(1)
'
'
'    Dim i                                              As Integer
'    Dim j                                              As Integer
'
'
'    xlSheet1.Cells(5, "C") = COMPANY_NAME
'    xlSheet1.Cells(6, "C") = cboMonthly_Month & " " & txtMonthly_Year
'    xlSheet1.Cells(11 + j, "A") = j + 1
'
'    Dim rsModels                                       As ADODB.Recordset
'    Dim X                                              As Integer
'    SQL = " SELECT  DISTINCT MODEL FROM SMIS_SALESORDER "
'    SQL = SQL & "    WHERE STATUS <>'C' AND INVOICEDDATE IS NOT NULL AND"
'    SQL = SQL & " Month(INVOICEDDATE) = " & What_month(cboMonthly_Month) & " And Year(INVOICEDDATE) = " & txtMonthly_Year
'
'    Set rsModels = gconDMIS.Execute(SQL)
'    While Not rsModels.EOF
'        xlSheet1.Cells(10, 6 + X) = rsModels!Model
'        X = X + 1
'        rsModels.MoveNext
'    Wend
'
'    Dim RSPER                                          As ADODB.Recordset
'    Do While Not emprs.EOF
'        xlSheet1.Cells(11 + j, "B") = Null2String(emprs!Name)
'        If IsDate(emprs!DATEHIRED) = True Then
'            xlSheet1.Cells(11 + j, "c") = MonthName(Month(emprs!DATEHIRED))
'        End If
'
'
'        SQL = " SELECT  MODEL, count(MODEL) AS TCOUNT FROM SMIS_SALESORDER "
'        SQL = SQL & "    WHERE STATUS <>'C' AND INVOICEDDATE IS NOT NULL AND"
'        SQL = SQL & " Month(INVOICEDDATE) = " & What_month(cboMonthly_Month) & " And Year(INVOICEDDATE) = " & txtMonthly_Year
'        SQL = SQL & " AND SALESAE='" & Null2String(emprs!Name) & "' GROUP BY MODEL"
'
'        Set RSPER = gconDMIS.Execute(SQL)
'
'        While Not RSPER.EOF
'            Select Case Null2String(RSPER!Model)
'                Case "Starex", "Starex Pro"
'                    xlSheet1.Cells(11 + j, "D") = NumericVal(RSPER!TCOUNT)
'                Case "VERA CRUZ"
'                    xlSheet1.Cells(11 + j, "E") = NumericVal(RSPER!TCOUNT)
'                Case "Santa Fe"
'                    xlSheet1.Cells(11 + j, "F") = NumericVal(RSPER!TCOUNT)
'                Case "Matrix"
'                    xlSheet1.Cells(11 + j, "G") = NumericVal(RSPER!TCOUNT)
'                Case "Tucson"
'                    xlSheet1.Cells(11 + j, "H") = NumericVal(RSPER!TCOUNT)
'                Case "Accent"
'                    xlSheet1.Cells(11 + j, "I") = NumericVal(RSPER!TCOUNT)
'                Case "Getz"
'                    xlSheet1.Cells(11 + j, "J") = NumericVal(RSPER!TCOUNT)
'
'                Case "Porter"
'                    xlSheet1.Cells(11 + j, "K") = NumericVal(RSPER!TCOUNT)
'
'            End Select
'            RSPER.MoveNext
'        Wend
'
'
'
'
'        j = j + 1
'        emprs.MoveNext
'
'    Loop
'
'    xlApp.Visible = True
'    Set xlBook = Nothing
'    Set xlSheet1 = Nothing
'    Set xlApp = Nothing
'End Sub
'COMMENTED BY: JUN - 02-15-2008 -----------THIS IS AN OLD REPORT PRINTING------------------------------

Sub PRINT_DAILY_VEHICLE_RETAIL_SALES()
    If Len(Dir(SMIS_REPORT_PATH & "SMIS_EXCEL\DAILY VEHICLE RETAIL SALES.xlt")) = 0 Then
        MsgBox "Excel Directory For Sales Managment Information Could Not be Located", vbInformation
        Exit Sub
    End If
    Dim SQL                                            As String
    Dim rsCust                                         As ADODB.Recordset
    Dim rsProsPEK                                      As ADODB.Recordset
    Dim rsGetOrNum                                     As ADODB.Recordset
    Dim rsAllCust                                      As ADODB.Recordset
    Dim getProsPECT                                    As String
    Dim getOrNum                                       As String
    Dim i                                              As Integer
    Dim j                                              As Integer
    Dim q                                              As Integer
    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet1                                       As Excel.Worksheet
    q = 1

    If gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESORDER WHERE (DAY(DateReleased)=" & Day(dtDay.Value) & " AND month(DateReleased)= " & Month(dtDay) & " AND year(DateReleased)= " & Year(dtDay) & ")").Fields(0).Value = 0 Then
        MsgBox " NO SALES RECORD FOR THE DATE ", vbInformation
        'MsgSpeech " NO SALES RECORD FOR THE DATE "
        Exit Sub
    End If

    Set rsAllCust = New ADODB.Recordset

    'Set rsCust = gconDMIS.Execute("SELECT CUSTNAME, HOMEADDRESS,certific8,vi_no,VDR_NO,MODELDESCRIPTION, VINO, ENGINENO, DATERELEASED,Certific8,IGNKEY_NO,TERM,SalesAE,code FROM SMIS_SALESORDER WHERE month(INVOICEDDATE)=" & What_month(cboMonthly_Month) & " AND year(INVOICEDDATE)= " & (txtMonthly_Year) & "  AND STATUS<>'C' ORDER BY INVOICEDDATE")
    Set rsCust = gconDMIS.Execute("SELECT * FROM SMIS_SALESORDER WHERE (DAY(DateReleased)=" & Day(dtDay.Value) & " AND month(DateReleased)= " & Month(dtDay) & " AND year(DateReleased)= " & Year(dtDay) & ") ORDER BY INVOICEDDATE")
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "SMIS_EXCEL\DAILY VEHICLE RETAIL SALES.xlt")
    Set xlSheet1 = xlBook.Worksheets(1)

    xlSheet1.Cells(6, "D") = COMPANY_NAME
    xlSheet1.Cells(7, "E") = cboMonthly_Month & " " & txtMonthly_Year

    Do While Not rsCust.EOF
        getProsPECT = Null2String(rsCust!CODE)
        If Not rsCust.EOF And Not rsCust.BOF Then
            Dim vProspectType, vleadsource             As String
            Set rsProsPEK = gconDMIS.Execute("Select ProspectType,leadsource from CRIS_Prospects where cuscde = '" & getProsPECT & "'")
            If Not rsProsPEK.EOF And Not rsProsPEK.BOF Then
                vProspectType = Null2String(rsProsPEK!ProspectType)
                vleadsource = Null2String(rsProsPEK!LeadSource)
            End If
        End If
        getOrNum = Null2String(rsCust!VI_NO)
        If Not rsCust.EOF And Not rsCust.BOF Then
            Dim Rnum                                   As String
            Set rsGetOrNum = gconDMIS.Execute("Select OR_NUM from CMIS_off_Dt where TRANTYPE = 'VI' and invoiceno = '" & getOrNum & "'")
            If Not rsGetOrNum.EOF And Not rsGetOrNum.BOF Then
                Rnum = Null2String(rsGetOrNum!OR_NUM)
            End If
        End If
        xlSheet1.Cells(15 + j, "A") = j + q
        If vProspectType = "P" Then
            xlSheet1.Cells(15 + j, "B") = "X"
        ElseIf vProspectType = "F" Then
            xlSheet1.Cells(15 + j, "C") = "X"
        Else
            xlSheet1.Cells(15 + j, "d") = "X"
        End If

        Set rsAllCust = gconDMIS.Execute("SELECT APOD,* FROM ALL_Customer WHERE CUSCDE='" & Null2String(getProsPECT) & "'")
        Dim Apod                                       As String
        If Not (rsAllCust.EOF And rsAllCust.BOF) Then
            Apod = Null2String(rsAllCust!Apod)
        End If

        If Apod = "MS" Then
            xlSheet1.Cells(15 + j, "G") = "X"
        ElseIf Apod = "MR" Then
            xlSheet1.Cells(15 + j, "F") = "X"
        Else
            xlSheet1.Cells(15 + j, "F") = ""
        End If

        xlSheet1.Cells(15 + j, "H") = Null2String(rsCust!CustName)
        xlSheet1.Cells(15 + j, "I") = Null2String(rsCust!HomeAddress)

        If Null2String(rsAllCust!Mobile) <> "" Then
            xlSheet1.Cells(15 + j, "J") = Null2String(rsAllCust!Mobile)
        End If
        If Null2String(rsCust!HomeTelNo) <> "" Then
            If xlSheet1.Cells(15 + j, "J") <> "" Then
                xlSheet1.Cells(15 + j, "J") = xlSheet1.Cells(15 + j, "J") & "/" & Null2String(rsCust!HomeTelNo)
            Else
                xlSheet1.Cells(15 + j, "J") = Null2String(rsCust!HomeTelNo)
            End If
        End If

        If Null2String(rsCust!officetelno) <> "" Then
            If xlSheet1.Cells(15 + j, "J") <> "" Then
                xlSheet1.Cells(15 + j, "J") = xlSheet1.Cells(15 + j, "J") & "/" & Null2String(rsCust!officetelno)
            Else
                xlSheet1.Cells(15 + j, "J") = Null2String(rsCust!officetelno)
            End If
        End If


        xlSheet1.Cells(15 + j, "K") = Null2String(rsCust!modeldescription)
        xlSheet1.Cells(15 + j, "L") = Null2String(rsCust!VINO)
        xlSheet1.Cells(15 + j, "M") = Null2String(rsCust!IGNKEY_NO)
        xlSheet1.Cells(15 + j, "N") = Null2String(rsCust!VI_NO)
        xlSheet1.Cells(15 + j, "O") = Format(Null2String(rsCust!InvoicedDate), "mm/dd/yyyy")
        xlSheet1.Cells(15 + j, "P") = Null2String(rsCust!VDR_NO)
        xlSheet1.Cells(15 + j, "Q") = Format(Null2String(rsCust!DateReleased), "mm/dd/yyyy")
        xlSheet1.Cells(15 + j, "R") = Null2String(rsCust!certific8)
        xlSheet1.Cells(15 + j, "S") = Null2String(rsCust!TERM)
        If Null2String(rsCust!Insured) = "I" Then
            xlSheet1.Cells(15 + j, "T") = "X"
        Else
            xlSheet1.Cells(15 + j, "U") = "X"
        End If
        xlSheet1.Cells(15 + j, "V") = Null2String(rsCust!salesae)
        xlSheet1.Cells(15 + j, "W") = vleadsource
        j = j + 1

        xlSheet1.Cells.Range("A" & 15 + j, "X" & 15 + j).Insert
        rsCust.MoveNext
    Loop

    Set rsCust = gconDMIS.Execute("SELECT DISTINCT  MODEL , COUNT(*) AS TCOUNT  FROM SMIS_SALESORDER WHERE (DAY(DateReleased)=" & Day(dtDay.Value) & " AND month(DateReleased)= " & Month(dtDay) & " AND year(DateReleased)= " & Year(dtDay) & ") GROUP BY MODEL ")
    Dim TOTALUNIT                                      As Integer
    Dim MCOUNT                                         As Integer

    '    Do While Not rsCust.EOF
    '        MCOUNT = MCOUNT + 1
    '        xlSheet1.Cells(22 + j, "H") = Null2String(rsCust!Model)
    '        xlSheet1.Cells(22 + j, "I") = Null2String(rsCust!TCOUNT)
    '        TOTALUNIT = TOTALUNIT + Null2String(rsCust!TCOUNT)
    '        j = j + 1
    '        rsCust.MoveNext
    '    Loop
    '    xlSheet1.Cells(23 + j, "H") = "TOTAL"
    '    xlSheet1.Cells(23 + j, "I") = TOTALUNIT

    Do While Not rsCust.EOF
        MCOUNT = MCOUNT + 1
        xlSheet1.Cells.BorderAround 1 = xlThin
        xlSheet1.Cells(20 + j, "K") = Null2String(rsCust!Model)
        xlSheet1.Cells(20 + j, "K").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(20 + j, "L") = Null2String(rsCust!TCOUNT)
        xlSheet1.Cells(20 + j, "L").BorderAround ColorIndex:=1, Weight:=xlThin
        TOTALUNIT = TOTALUNIT + Null2String(rsCust!TCOUNT)
        j = j + 1
        rsCust.MoveNext
    Loop
    xlSheet1.Cells(20 + j, "K") = "TOTAL"
    xlSheet1.Cells(20 + j, "K").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(20 + j, "K").Font.Bold = True
    xlSheet1.Cells(20 + j, "L") = TOTALUNIT
    xlSheet1.Cells(20 + j, "L").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(20 + j, "L").Font.Bold = True

    '    'SELECT CRIS_PROSPECTS.LEADSOURCE , COUNT(*) FROM SMIS_SALESORDER
    '    'LEFT OUTER JOIN CRIS_PROSPECTS ON CRIS_PROSPECTS.PROSPECTID =SMIS_SALESORDER.PROSPECTID
    '    'GROUP BY CRIS_PROSPECTS.LEADSOURCE
    Screen.MousePointer = 0
    xlApp.Visible = True
    Set xlBook = Nothing
    Set xlSheet1 = Nothing
    Set xlApp = Nothing
    Set rsProsPEK = Nothing
    Set rsGetOrNum = Nothing
End Sub

Sub PRINT_MONTHLY_DEALER_RETAIL_SALES()
    If Len(Dir(SMIS_REPORT_PATH & "SMIS_EXCEL\MONTHLY DEALER RETAIL SALES.xlt")) = 0 Then
        MsgBox "Excel Directory For Sales Managment Information Could Not be Located", vbInformation
        Exit Sub
    End If
    
    Dim SQL                                            As String
    Dim rsCust                                         As ADODB.Recordset
    Dim rsProsPEK                                      As ADODB.Recordset
    Dim rsGetOrNum                                     As ADODB.Recordset
    Dim rsAllCust                                      As ADODB.Recordset
    Dim getProsPECT                                    As String
    Dim getOrNum                                       As String
    Dim i                                              As Integer
    Dim j                                              As Integer
    Dim q                                              As Integer
    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet1                                       As Excel.Worksheet
    q = 1

    If gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESORDER WHERE month(INVOICEDDATE)=" & What_month(cboMonthly_Month) & " AND year(INVOICEDDATE)= " & (txtMonthly_Year) & "  AND STATUS<>'C'").Fields(0).Value = 0 Then
        MsgBox " NO SALES RECORD FOR THE DATE ", vbInformation
        'MsgSpeech " NO SALES RECORD FOR THE DATE "
        Exit Sub
    End If

    Set rsAllCust = New ADODB.Recordset

    'Set rsCust = gconDMIS.Execute("SELECT CUSTNAME, HOMEADDRESS,certific8,vi_no,VDR_NO,MODELDESCRIPTION, VINO, ENGINENO, DATERELEASED,Certific8,IGNKEY_NO,TERM,SalesAE,code FROM SMIS_SALESORDER WHERE month(INVOICEDDATE)=" & What_month(cboMonthly_Month) & " AND year(INVOICEDDATE)= " & (txtMonthly_Year) & "  AND STATUS<>'C' ORDER BY INVOICEDDATE")
    Set rsCust = gconDMIS.Execute("SELECT * FROM SMIS_SALESORDER WHERE " & _
        " month(INVOICEDDATE) = " & What_month(cboMonthly_Month) & _
        " AND year(INVOICEDDATE) = " & (txtMonthly_Year) & _
        " AND ISNULL(STATUS,'') <> 'C' ORDER BY INVOICEDDATE")
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "SMIS_EXCEL\MONTHLY DEALER RETAIL SALES.xlt")
    Set xlSheet1 = xlBook.Worksheets(1)

    xlSheet1.Cells(6, "D") = COMPANY_NAME
    xlSheet1.Cells(7, "E") = cboMonthly_Month & " " & txtMonthly_Year

    Do While Not rsCust.EOF
        getProsPECT = Null2String(rsCust!CODE)
        If Not rsCust.EOF And Not rsCust.BOF Then
            Dim vProspectType, vleadsource             As String
            Set rsProsPEK = gconDMIS.Execute("Select ProspectType,leadsource from CRIS_Prospects where cuscde = '" & getProsPECT & "'")
            If Not rsProsPEK.EOF And Not rsProsPEK.BOF Then
                vProspectType = Null2String(rsProsPEK!ProspectType)
                vleadsource = Null2String(rsProsPEK!LeadSource)
            End If
        End If
        getOrNum = Null2String(rsCust!VI_NO)
        If Not rsCust.EOF And Not rsCust.BOF Then
            Dim Rnum                                   As String
            Set rsGetOrNum = gconDMIS.Execute("Select OR_NUM from CMIS_off_Dt where TRANTYPE = 'VI' and invoiceno = '" & getOrNum & "'")
            If Not rsGetOrNum.EOF And Not rsGetOrNum.BOF Then
                Rnum = Null2String(rsGetOrNum!OR_NUM)
            End If
        End If
        xlSheet1.Cells(15 + j, "A") = j + q
        If vProspectType = "P" Then
            xlSheet1.Cells(15 + j, "B") = "X"
        ElseIf vProspectType = "F" Then
            xlSheet1.Cells(15 + j, "C") = "X"
        Else
            xlSheet1.Cells(15 + j, "d") = "X"
        End If

        Set rsAllCust = gconDMIS.Execute("SELECT APOD,* FROM ALL_Customer WHERE CUSCDE='" & Null2String(getProsPECT) & "'")
        Dim Apod                                       As String
        If Not (rsAllCust.EOF And rsAllCust.BOF) Then
            Apod = Null2String(rsAllCust!Apod)
        End If

        If Apod = "MS" Then
            xlSheet1.Cells(15 + j, "G") = "X"
        ElseIf Apod = "MR" Then
            xlSheet1.Cells(15 + j, "F") = "X"
        Else
            xlSheet1.Cells(15 + j, "F") = ""
        End If

        xlSheet1.Cells(15 + j, "H") = Null2String(rsCust!CustName)
        xlSheet1.Cells(15 + j, "I") = Null2String(rsCust!HomeAddress)
        If Null2String(rsAllCust!Mobile) <> "" Then
            xlSheet1.Cells(15 + j, "J") = Null2String(rsAllCust!Mobile)
        End If

        If Null2String(rsCust!HomeTelNo) <> "" Then
            If Null2String(rsAllCust!Mobile) <> "" Then
                xlSheet1.Cells(15 + j, "J") = Null2String(rsAllCust!Mobile) & "/" & Null2String(rsCust!HomeTelNo)
            Else
                xlSheet1.Cells(15 + j, "J") = Null2String(rsCust!HomeTelNo)
            End If
        End If

        If Null2String(rsCust!officetelno) <> "" Then
            If Null2String(rsCust!HomeTelNo) <> "" Then
                xlSheet1.Cells(15 + j, "J") = Null2String(rsCust!HomeTelNo) & "/" & Null2String(rsCust!officetelno)
            Else
                xlSheet1.Cells(15 + j, "J") = Null2String(rsCust!officetelno)
            End If
        End If

        xlSheet1.Cells(15 + j, "K") = Null2String(rsCust!modeldescription)
        xlSheet1.Cells(15 + j, "L") = Null2String(rsCust!VINO)
        xlSheet1.Cells(15 + j, "M") = Null2String(rsCust!IGNKEY_NO)
        xlSheet1.Cells(15 + j, "N") = Null2String(rsCust!VI_NO)
        xlSheet1.Cells(15 + j, "O") = Format(Null2String(rsCust!InvoicedDate), "mm/dd/yyyy")
        xlSheet1.Cells(15 + j, "P") = Null2String(rsCust!VDR_NO)
        xlSheet1.Cells(15 + j, "Q") = Format(Null2String(rsCust!DateReleased), "mm/dd/yyyy")
        xlSheet1.Cells(15 + j, "R") = Null2String(rsCust!certific8)
        xlSheet1.Cells(15 + j, "S") = Null2String(rsCust!TERM)
        If Null2String(rsCust!Insured) = "I" Then
            xlSheet1.Cells(15 + j, "T") = "X"
        Else
            xlSheet1.Cells(15 + j, "U") = "X"
        End If
        xlSheet1.Cells(15 + j, "V") = Null2String(rsCust!salesae)
        xlSheet1.Cells(15 + j, "W") = vleadsource
        j = j + 1

        xlSheet1.Cells.Range("A" & 15 + j, "X" & 15 + j).Insert
        rsCust.MoveNext
    Loop

    Set rsCust = gconDMIS.Execute("SELECT DISTINCT  MODEL , COUNT(*) AS TCOUNT  FROM SMIS_SALESORDER WHERE month(INVOICEDDATE)=" & What_month(cboMonthly_Month) & " AND year(INVOICEDDATE)= " & (txtMonthly_Year) & "  AND STATUS<>'C' GROUP BY MODEL ")
    Dim TOTALUNIT                                      As Integer
    Dim MCOUNT                                         As Integer

    '    Do While Not rsCust.EOF
    '        MCOUNT = MCOUNT + 1
    '        xlSheet1.Cells(22 + j, "H") = Null2String(rsCust!Model)
    '        xlSheet1.Cells(22 + j, "I") = Null2String(rsCust!TCOUNT)
    '        TOTALUNIT = TOTALUNIT + Null2String(rsCust!TCOUNT)
    '        j = j + 1
    '        rsCust.MoveNext
    '    Loop
    '    xlSheet1.Cells(23 + j, "H") = "TOTAL"
    '    xlSheet1.Cells(23 + j, "I") = TOTALUNIT

    Do While Not rsCust.EOF
        MCOUNT = MCOUNT + 1
        xlSheet1.Cells.BorderAround 1 = xlThin
        xlSheet1.Cells(20 + j, "K") = Null2String(rsCust!Model)
        xlSheet1.Cells(20 + j, "K").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(20 + j, "L") = Null2String(rsCust!TCOUNT)
        xlSheet1.Cells(20 + j, "L").BorderAround ColorIndex:=1, Weight:=xlThin
        TOTALUNIT = TOTALUNIT + Null2String(rsCust!TCOUNT)
        j = j + 1
        rsCust.MoveNext
    Loop
    xlSheet1.Cells(20 + j, "K") = "TOTAL"
    xlSheet1.Cells(20 + j, "K").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(20 + j, "K").Font.Bold = True
    xlSheet1.Cells(20 + j, "L") = TOTALUNIT
    xlSheet1.Cells(20 + j, "L").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(20 + j, "L").Font.Bold = True

    '    'SELECT CRIS_PROSPECTS.LEADSOURCE , COUNT(*) FROM SMIS_SALESORDER
    '    'LEFT OUTER JOIN CRIS_PROSPECTS ON CRIS_PROSPECTS.PROSPECTID =SMIS_SALESORDER.PROSPECTID
    '    'GROUP BY CRIS_PROSPECTS.LEADSOURCE
    Screen.MousePointer = 0
    xlApp.Visible = True
    Set xlBook = Nothing
    Set xlSheet1 = Nothing
    Set xlApp = Nothing
    Set rsProsPEK = Nothing
    Set rsGetOrNum = Nothing
End Sub

Private Sub cboReportType_Click()
    Select Case cboReportType
        Case "DAILY VEHICLE RETAIL SALES"
            picMonthly.Visible = False: picRange.Visible = True
        Case "DEALER ENDING INVENTORY"
            picMonthly.Visible = False: picRange.Visible = False
        Case "DEALER SALES CONSULTANTS PERFORMANCE SHEET"
            picMonthly.Visible = True: picRange.Visible = False
        Case "MONTHLY DEALER RETAIL SALES"
            picMonthly.Visible = True: picRange.Visible = False

    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Select Case cboReportType
        Case "DAILY VEHICLE RETAIL SALES"
            Call PRINT_DAILY_VEHICLE_RETAIL_SALES
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 5:00
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Call NEW_LogAudit("V", "HYUNDAI REPORTS", "", "", "", cboReportType & " " & dtDay, "", "")
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

            'LogAudit "V", "HYUNDAI DAILY VEHICLE RETAIL SALES", dtDay
        Case "DEALER ENDING INVENTORY"
            Call PRINT_DEALER_ENDING_INVENTORY
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 5:00
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Call NEW_LogAudit("V", "HYUNDAI REPORTS", "", "", "", cboReportType, "", "")
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
            'LogAudit "V", "HYUNDAI DEALER ENDING INVENTORY", dtDay
        Case "DEALER SALES CONSULTANTS PERFORMANCE SHEET"
            'MsgBox "UNDER DEVELOPMENT": Exit Sub
            Call PRINT_DEALER_SALES_CONSULTANTS_PERFORMANCE_SHEET
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 5:00
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Call NEW_LogAudit("V", "HYUNDAI REPORTS", "", "", "", cboReportType & " " & dtDay, "", "")
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Case "MONTHLY DEALER RETAIL SALES"
            Call PRINT_MONTHLY_DEALER_RETAIL_SALES
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 5:00
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Call NEW_LogAudit("V", "HYUNDAI REPORTS", "", "", "", cboReportType & " " & cboMonthly_Month & " " & txtMonthly_Year, "", "")
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
            'LogAudit "V", "HYUNDAI MONTHLY DEALER RETAIL SALES", cboMonthly_Month
            '   Case "VEHICLES SALES PROJECTION" ''*********Updated by Ryan April 26 2008
            '     PrintSQLReport CrystalReport1, SMIS_REPORT_PATH & "VS\VehiclesSalesProjection.rpt", "", DMIS_REPORT_Connection, 1

    End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (Mitsubishi REPORTS)"
            Call frmALL_AuditInquiry.DisplayHistory("", "Mitsubishi REPORTS", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonthly_Month

    dtDay = firstDay(Date)

    txtMonthly_Year = Year(LOGDATE)
    cboMonthly_Month = MonthName(Month(LOGDATE))

    With cboReportType
        .AddItem "DAILY VEHICLE RETAIL SALES"
        .AddItem "DEALER ENDING INVENTORY"
        .AddItem "DEALER SALES CONSULTANTS PERFORMANCE SHEET"
        .AddItem "MONTHLY DEALER RETAIL SALES"
        '        .AddItem "VEHICLES SALES PROJECTION" ''*********Updated by Ryan April 27 2008
        .ListIndex = 0
    End With

End Sub

''Private Sub Timer1_Timer()
''    If lblLoading.Caption <> "" Then
''        If lblLoading.Visible = True Then
''            lblLoading.Visible = False
''        Else
''            lblLoading.Visible = True
''        End If
''    End If
''End Sub

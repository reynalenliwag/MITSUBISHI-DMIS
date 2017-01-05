VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Sales_report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Report"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
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
      Height          =   825
      Left            =   1680
      MouseIcon       =   "frm_Sales_report.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frm_Sales_report.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1320
      Width           =   885
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
      Height          =   825
      Left            =   2640
      MouseIcon       =   "frm_Sales_report.frx":05F1
      MousePointer    =   99  'Custom
      Picture         =   "frm_Sales_report.frx":0743
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   1320
      Width           =   885
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   2160
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Units Released"
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
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   661
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
      Format          =   102825985
      CurrentDate     =   39427
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   661
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
      Format          =   102825985
      CurrentDate     =   39427
   End
   Begin wizProgBar.Prg progDSSR 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   2400
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   556
      Picture         =   "frm_Sales_report.frx":0B8E
      ForeColor       =   0
      BarPicture      =   "frm_Sales_report.frx":0BAA
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
   Begin VB.Label labPercent 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   600
   End
End
Attribute VB_Name = "frm_Sales_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------KEVIN D LLANES---------------------------------------------'
'-----------------------------------SALES REPORT----------------------09-10-14-------'
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim rsCust                          As ADODB.Recordset
    Dim SQL                             As String
    Dim Value                           As Integer
    Dim j                               As Integer
    Dim xlApp
    Dim xlbook
    Dim xlSheet1
    Dim xlSheet2

    
    On Error GoTo ErrorCode
    Set rsCust = New ADODB.Recordset

SQL = "SELECT SOURCE,all_customer_table.ACCTNAME,all_customer_table.CUSTOMERADD,SMIS_SALESORDER.INVOICEDDATE ,SMIS_SALESORDER.VI_NO,MODELDESCRIPTION,SMIS_SALESORDER.VINO,SMIS_SALESORDER.ENGINENO, SMIS_SALESORDER.COLOR,SMIS_SalesOrder.TIN FROM SMIS_SalesOrder"
SQL = SQL & " INNER JOIN SMIS_MRRINV_TABLE as A ON SMIS_SALESORDER.CODE = A.CUSTOMERCODE inner join all_customer_table on a.CUSTOMERCODE = all_customer_table.cuscde WHERE SMIS_SalesOrder.INVOICEDDATE BETWEEN  '" & dtFrom & "' AND '" & dtTo & "'"


    Set rsCust = gconDMIS.Execute(SQL)

    If rsCust.EOF Or rsCust.BOF Then
        MsgSpeechBox " There Are No Records for the Specified Date"
        Exit Sub
    End If

   
    Set xlApp = CreateObject("Excel.Application")

    Set xlbook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "\SMIS_EXCEL\Sales_report.xlt")
    Set xlSheet1 = xlbook.Worksheets(1)
    Set xlSheet2 = xlbook.Worksheets(2)
 
    If Not rsCust.EOF And Not rsCust.BOF Then
        Value = 0
        progDSSR.Value = Value
        Do While Not rsCust.EOF
    If Null2String(rsCust!InvoicedDate) = "" Or Null2String(rsCust!InvoicedDate) = "" Then
                
                xlSheet1.Cells(2, 2) = COMPANY_NAME
                xlSheet1.Cells(3, 2) = COMPANY_ADDRESS
                'xlSheet1.Cells(6 + j, 2) = Null2String(rsCust!CustName)
                xlSheet1.Cells(6 + j, 3) = Null2String(rsCust!AcctName)
                xlSheet1.Cells(6 + j, 4) = Null2String(rsCust!CUSTOMERADD)
                xlSheet1.Cells(6 + j, 5) = Null2String(rsCust!InvoicedDate)
                xlSheet1.Cells(6 + j, 6) = Null2String(rsCust!VI_NO)
                xlSheet1.Cells(6 + j, 7) = Null2String(rsCust!modeldescription)
                xlSheet1.Cells(6 + j, 8) = Null2String(rsCust!Vino)
                xlSheet1.Cells(6 + j, 9) = Null2String(rsCust!EngineNo)
                xlSheet1.Cells(6 + j, 10) = Null2String(rsCust!Color)
                xlSheet1.Cells(6 + j, 11) = Null2String(rsCust!TIN)
 
               

    Else
                
                xlSheet1.Cells(2, 2) = COMPANY_NAME
                xlSheet1.Cells(3, 2) = COMPANY_ADDRESS
                xlSheet1.Cells(6 + j, 1) = Null2String(rsCust!Source)
                'xlSheet1.Cells(6 + j, 2) = Null2String(rsCust!CustName)
                xlSheet1.Cells(6 + j, 3) = Null2String(rsCust!AcctName)
                xlSheet1.Cells(6 + j, 4) = Null2String(rsCust!CUSTOMERADD)
                xlSheet1.Cells(6 + j, 5) = Null2String(rsCust!InvoicedDate)
                xlSheet1.Cells(6 + j, 6) = Null2String(rsCust!VI_NO)
                xlSheet1.Cells(6 + j, 7) = Null2String(rsCust!modeldescription)
                xlSheet1.Cells(6 + j, 8) = Null2String(rsCust!Vino)
                xlSheet1.Cells(6 + j, 9) = Null2String(rsCust!EngineNo)
                xlSheet1.Cells(6 + j, 10) = Null2String(rsCust!Color)
                xlSheet1.Cells(6 + j, 11) = Null2String(rsCust!TIN)


    End If
        j = j + 1
         Value = Value + 1
        progDSSR.Value = (Value / rsCust.RecordCount) * 100
        progDSSR.Text = Int(progDSSR.Value) & "%"
        rsCust.MoveNext
        Loop
    End If
    xlApp.Visible = True
    Set xlbook = Nothing
    Set xlSheet1 = Nothing
    Set xlSheet2 = Nothing
    Set xlApp = Nothing


    Exit Sub
ErrorCode:
    ShowVBError
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub
Private Sub txtTo_GotFocus()
    progDSSR.Value = 0
    labPercent.Caption = ""
End Sub
Private Sub Form_Load()
    dtFrom = DateValue(firstDay(LOGDATE))
    dtTo = Date
End Sub

'----------------------------------------KEVIN D LLANES---------------------------------------------'
'-----------------------------------SALES REPORT----------------------09-10-14-------'






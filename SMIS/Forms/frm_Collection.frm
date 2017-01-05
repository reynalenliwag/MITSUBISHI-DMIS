VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Collection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STATUS OF COLLECTION OF SOLD UNITS"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4245
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
      Left            =   2400
      MouseIcon       =   "frm_Collection.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frm_Collection.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   1200
      Width           =   885
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
      Height          =   825
      Left            =   1440
      MouseIcon       =   "frm_Collection.frx":059D
      MousePointer    =   99  'Custom
      Picture         =   "frm_Collection.frx":06EF
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   1200
      Width           =   885
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   1920
      Top             =   1560
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
      Left            =   1080
      TabIndex        =   2
      Top             =   120
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
      Format          =   88080385
      CurrentDate     =   39427
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   600
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
      Format          =   88080385
      CurrentDate     =   39427
   End
   Begin wizProgBar.Prg progDSSR 
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   556
      Picture         =   "frm_Collection.frx":0B8E
      ForeColor       =   0
      BarPicture      =   "frm_Collection.frx":0BAA
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
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   600
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
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   600
   End
   Begin VB.Label labPercent 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "frm_Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------KEVIN D LLANES---------------------------------------------'
'-----------------------------------COLLECTION STATUS----------------------08-11-14-------'


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

SQL = "select InvoicedDate,VI_NO,CustName,Term,Total,OR_NUM_DT,AMOUNT_DT,BALANCE,OR_NUM,AMOUNT,BalToFinanced,Model,"
SQL = SQL & "(select Hd.VoucherNo from AMIS_JOURNAL_HD HD"
SQL = SQL & " inner join AMIS_JOURNAL_DET DT on hd.JType = DT.JType and hd.VoucherNo = DT.VoucherNo"
SQL = SQL & " inner join AMIS_CHARTACCOUNT AC on AC.AcctCode = dt.Acct_Code"
SQL = SQL & " where ac.TranType3 = 'Discount' and ac.trantype2 = 'SALES'and ac.TranType1= model and  HD.JType = 'CCM' and HD.InvoiceType = 'VI' and HD.InvoiceNo  = VI_NO) as CMREF,"
SQL = SQL & " (select dt.Debit from AMIS_JOURNAL_HD HD"
SQL = SQL & " inner join AMIS_JOURNAL_DET DT on hd.JType = DT.JType and hd.VoucherNo = DT.VoucherNo"
SQL = SQL & " inner join AMIS_CHARTACCOUNT AC on AC.AcctCode = dt.Acct_Code"
SQL = SQL & " where ac.TranType3 = 'Discount' and ac.trantype2 = 'SALES'and ac.TranType1= model and  HD.JType = 'CCM' and HD.InvoiceType = 'VI' and HD.InvoiceNo  = VI_NO) as DEBIT   from"
SQL = SQL & "(select model,InvoicedDate,a.VI_NO,a.CustName,a.Term,a.Total,b.OR_NUM as OR_NUM_DT,b.AMOUNT as AMOUNT_DT,b.BALANCE ,c.OR_NUM,c.AMOUNT,a.BalToFinanced from SMIS_SalesOrder a inner join CMIS_Off_Dt   b on a.VI_NO=b.INVOICENO inner join CMIS_DEPOSITDT c on c.INVOICENO=b.INVOICENO)z"






    Set rsCust = gconDMIS.Execute(SQL)

    If rsCust.EOF Or rsCust.BOF Then
        MsgSpeechBox " There Are No Records for the Specified Date"
        Exit Sub
    End If

   
    Set xlApp = CreateObject("Excel.Application")

    Set xlbook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "\SMIS_EXCEL\Collection_Status.xlt")
    Set xlSheet1 = xlbook.Worksheets(1)
    Set xlSheet2 = xlbook.Worksheets(2)
 
    If Not rsCust.EOF And Not rsCust.BOF Then
        Value = 0
        progDSSR.Value = Value
        Do While Not rsCust.EOF
    If Null2String(rsCust!InvoicedDate) = "" Or Null2String(rsCust!InvoicedDate) = "" Then
                
                xlSheet1.Cells(3 + j, 1) = Null2String(rsCust!InvoicedDate)
                xlSheet1.Cells(3 + j, 2) = Null2String(rsCust!VI_NO)
                xlSheet1.Cells(3 + j, 3) = Null2String(rsCust!CustName)
                xlSheet1.Cells(3 + j, 4) = Null2String(rsCust!TERM)
                xlSheet1.Cells(3 + j, 5) = Null2String(rsCust!Total)
                xlSheet1.Cells(3 + j, 6) = Null2String(rsCust!CMREF)
                xlSheet1.Cells(3 + j, 7) = Null2String(rsCust!DEBIT)
                xlSheet1.Cells(3 + j, 8) = Null2String(rsCust!OR_NUM)
                xlSheet1.Cells(3 + j, 9) = Null2String(rsCust!amount)
                xlSheet1.Cells(3 + j, 10) = Null2String(rsCust!OR_NUM_DT)
                xlSheet1.Cells(3 + j, 11) = Null2String(rsCust!AMOUNT_DT)
                xlSheet1.Cells(3 + j, 12) = Null2String(rsCust!BALANCE)
                xlSheet1.Cells(3 + j, 13) = Null2String(rsCust!BALTOFINANCED)
                
               

    Else
                
                xlSheet1.Cells(3 + j, 1) = Null2String(rsCust!InvoicedDate)
                xlSheet1.Cells(3 + j, 2) = Null2String(rsCust!VI_NO)
                xlSheet1.Cells(3 + j, 3) = Null2String(rsCust!CustName)
                xlSheet1.Cells(3 + j, 4) = Null2String(rsCust!TERM)
                xlSheet1.Cells(3 + j, 5) = Null2String(rsCust!Total)
                xlSheet1.Cells(3 + j, 6) = Null2String(rsCust!CMREF)
                xlSheet1.Cells(3 + j, 7) = Null2String(rsCust!DEBIT)
                xlSheet1.Cells(3 + j, 8) = Null2String(rsCust!OR_NUM)
                xlSheet1.Cells(3 + j, 9) = Null2String(rsCust!amount)
                xlSheet1.Cells(3 + j, 10) = Null2String(rsCust!OR_NUM_DT)
                xlSheet1.Cells(3 + j, 11) = Null2String(rsCust!AMOUNT_DT)
                xlSheet1.Cells(3 + j, 12) = Null2String(rsCust!BALANCE)
                xlSheet1.Cells(3 + j, 13) = Null2String(rsCust!BALTOFINANCED)


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
'-----------------------------------COLLECTION STATUS----------------------08-11-14-------'





VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_General_Vehicle_Data 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General_Vehicle_Data"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
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
      Left            =   1560
      MouseIcon       =   "frm_General_Vehicle_Data.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frm_General_Vehicle_Data.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1200
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
      Left            =   2520
      MouseIcon       =   "frm_General_Vehicle_Data.frx":05F1
      MousePointer    =   99  'Custom
      Picture         =   "frm_General_Vehicle_Data.frx":0743
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   1200
      Width           =   885
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   2040
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
      Left            =   1200
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
      Format          =   91815937
      CurrentDate     =   39427
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   1200
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
      Format          =   91815937
      CurrentDate     =   39427
   End
   Begin wizProgBar.Prg progDSSR 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Top             =   2280
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   556
      Picture         =   "frm_General_Vehicle_Data.frx":0B8E
      ForeColor       =   0
      BarPicture      =   "frm_General_Vehicle_Data.frx":0BAA
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
      Left            =   1560
      TabIndex        =   7
      Top             =   2280
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
      Left            =   480
      TabIndex        =   6
      Top             =   120
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
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   600
   End
End
Attribute VB_Name = "frm_General_Vehicle_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------KEVIN D LLANES---------------------------------------------'
'-----------------------------------GENERAL VEHICLE DATA REPORT----------------------08-06-14-------'


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

    SQL = "SELECT FIRSTNAME, LASTNAME,CUSTOMERADD, PROVINCIALADD, CITY ,HOMEPHONE, MOBILE,SEX,ZIPCODE,MIDDLEINITIAL,CUSTYPE, "
    SQL = SQL & " invoiceddate,VI_NO,plate_no, conductionSticker,term,downpaymentrate,terms,Vino,SalesAE,smis_salesorder.BirthDate  FROM smis_salesorder "
    SQL = SQL & " INNER JOIN all_customer ON all_customer.CUSCDE=smis_salesorder.CODE WHERE smis_salesorder.datereleased BETWEEN  '" & dtFrom & "' AND '" & dtTo & "'"
    SQL = SQL & " ORDER BY DateReleased "

    Set rsCust = gconDMIS.Execute(SQL)

    If rsCust.EOF Or rsCust.BOF Then
        MsgSpeechBox " There Are No Records for the Specified Date"
        Exit Sub
    End If

   
    Set xlApp = CreateObject("Excel.Application")

    Set xlbook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "\SMIS_EXCEL\GVD.xlt")
    Set xlSheet1 = xlbook.Worksheets(1)
    Set xlSheet2 = xlbook.Worksheets(2)
 
    If Not rsCust.EOF And Not rsCust.BOF Then
        Value = 0
        progDSSR.Value = Value
        Do While Not rsCust.EOF
    If Null2String(rsCust!CUSTYPE) = "" Or Null2String(rsCust!CUSTYPE) = "" Then
                
                xlSheet1.Cells(4 + j, 1) = COMPANY_NAME
               ' xlSheet1.Cells(4 + j, 2) = Null2String(rsCust!FIRSTNAME)
                xlSheet1.Cells(4 + j, 3) = Null2String(rsCust!InvoicedDate)
                xlSheet1.Cells(4 + j, 4) = Null2String(rsCust!VI_NO)
                xlSheet1.Cells(4 + j, 5) = Null2String(rsCust!FIRSTNAME)
                xlSheet1.Cells(4 + j, 6) = Null2String(rsCust!MiddleInitial)
                xlSheet1.Cells(4 + j, 7) = Null2String(rsCust!lastname)
                xlSheet1.Cells(4 + j, 8) = Null2String(rsCust!CUSTOMERADD)
                xlSheet1.Cells(4 + j, 9) = Null2String(rsCust!provincialadd) & " " & Null2String(rsCust!CITY)
                xlSheet1.Cells(4 + j, 10) = Null2String(rsCust!BirthDate)
                xlSheet1.Cells(4 + j, 11) = Null2String(rsCust!Sex)
                xlSheet1.Cells(4 + j, 12) = Null2String(rsCust!ZIPCODE)
                xlSheet1.Cells(4 + j, 13) = Null2String(rsCust!Mobile)
                xlSheet1.Cells(4 + j, 14) = Null2String(rsCust!Vino)
                xlSheet1.Cells(4 + j, 15) = Null2String(rsCust!ConductionSticker)
                xlSheet1.Cells(4 + j, 16) = Null2String(rsCust!PLATE_NO)
                xlSheet1.Cells(4 + j, 17) = Null2String(rsCust!salesae)
                xlSheet1.Cells(4 + j, 18) = Null2String(rsCust!TERM)
                xlSheet1.Cells(4 + j, 19) = Null2String(rsCust!DOWNPAYMENTRATE)
                xlSheet1.Cells(4 + j, 20) = Null2String(rsCust!TERMS)

    Else
                
                 xlSheet1.Cells(4 + j, 1) = COMPANY_NAME
                'xlSheet1.Cells(4 + j, 2) = Null2String(rsCust!FIRSTNAME)
                xlSheet1.Cells(4 + j, 3) = Null2String(rsCust!InvoicedDate)
                xlSheet1.Cells(4 + j, 4) = Null2String(rsCust!VI_NO)
                xlSheet1.Cells(4 + j, 5) = Null2String(rsCust!FIRSTNAME)
                xlSheet1.Cells(4 + j, 6) = Null2String(rsCust!MiddleInitial)
                xlSheet1.Cells(4 + j, 7) = Null2String(rsCust!lastname)
                xlSheet1.Cells(4 + j, 8) = Null2String(rsCust!CUSTOMERADD)
                xlSheet1.Cells(4 + j, 9) = Null2String(rsCust!provincialadd) & " " & Null2String(rsCust!CITY)
                xlSheet1.Cells(4 + j, 10) = Null2String(rsCust!BirthDate)
                xlSheet1.Cells(4 + j, 11) = Null2String(rsCust!Sex)
                xlSheet1.Cells(4 + j, 12) = Null2String(rsCust!ZIPCODE)
                xlSheet1.Cells(4 + j, 13) = Null2String(rsCust!Mobile)
                xlSheet1.Cells(4 + j, 14) = Null2String(rsCust!Vino)
                xlSheet1.Cells(4 + j, 15) = Null2String(rsCust!ConductionSticker)
                xlSheet1.Cells(4 + j, 16) = Null2String(rsCust!PLATE_NO)
                xlSheet1.Cells(4 + j, 17) = Null2String(rsCust!salesae)
                xlSheet1.Cells(4 + j, 18) = Null2String(rsCust!TERM)
                xlSheet1.Cells(4 + j, 19) = Null2String(rsCust!DOWNPAYMENTRATE)
                xlSheet1.Cells(4 + j, 20) = Null2String(rsCust!TERMS)


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
'-----------------------------------GENERAL VEHICLE DATA REPORT----------------------08-06-14-------'



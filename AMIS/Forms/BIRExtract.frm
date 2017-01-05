VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmBIRExtract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BIR Data Extraction"
   ClientHeight    =   2475
   ClientLeft      =   210
   ClientTop       =   645
   ClientWidth     =   3180
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "BIRExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3180
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00F1F6F5&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   345
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   480
      Width           =   1965
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00F1F6F5&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   345
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   870
      Width           =   1965
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   210
      TabIndex        =   3
      Top             =   3030
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   556
      Picture         =   "BIRExtract.frx":030A
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "BIRExtract.frx":0326
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2160
      MouseIcon       =   "BIRExtract.frx":0342
      MousePointer    =   99  'Custom
      Picture         =   "BIRExtract.frx":0494
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit Window"
      Top             =   1290
      Width           =   720
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Process"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1455
      MouseIcon       =   "BIRExtract.frx":07FA
      MousePointer    =   99  'Custom
      Picture         =   "BIRExtract.frx":094C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Process BIR Data Extraction"
      Top             =   1290
      Width           =   720
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait until excel file appears."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   510
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   735
   End
   Begin VB.Label labCPB 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   5835
   End
End
Attribute VB_Name = "frmBIRExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'       ´$$$$`                             ,,,
'       ´$$$$$$$`                         ´$$$`
'        `$$$$$$$`      ,,        ,,      ´$$$$´
'         `$$$$$$$`    ´$$`     ´$$`    ´$$$$$´
'          `$$$$$$$`_´$$$$$`_´$$$$$`__´$$$$$$$´
'           `$$$$$$$_$$$$$$$_$$$$$$$_´$$$$$$$´
'            `$$$$$$_$$$$$$$_$$$$$$$`´$$$$$$´
'   ,,,,,    `$$$$$$_$$$$$$$_$$$$$$$_$$$$$$´
' ´$$$$$`    `$$$$$$_$$$$$$$_$$$$$$$_$$$$$$´
'´$$$$$$$$$`´$$$$$$$_SARAJOY_$$$$$$$_$$$$$´
'´$$$$$$$$$$$$$$$$$$_$$$$$$$_$$$$$$$_$$$$$´
'   `$$$$$$$$$$$$$$$_$$$$$$$_$$$$$$_$$$$$$´
'      `$$$$$$$$$$$$$_$$$$$__$$_$$$$$$_$$´
'       `$$$$$$$$$$$$$__,$$$$,_____,$$$$$´
'         `$$$$$$$$$$$$$$$$$$$$$$$$$$$$$´
'          `$$$$$$$$$$$$$$$$$$$$$$$$$$$´
'            `$$$$$$$$CREATED$ON$$$$$$´
'               `$$$$$NOV292016$$$$$´
'
'              ROCK 'N ROLL TO THE WORLD
Option Explicit

Dim rsExtractBirRelief As ADODB.Recordset
Dim rsExtractSummary As ADODB.Recordset
    
Dim oXLApp As Object
Dim oXLBook As Object
Dim oXLSheet As Object
Dim sFilePath As String
Dim CsvDatFileName As String

Private Sub cmdCheck_Click()
    sFilePath = AMIS_REPORT_PATH & "BirRelief\"
    
    If Dir(sFilePath & "Purchases.xlsx") = "" Then
      MsgBox "Directory not found.", vbCritical, "Error"
      Exit Sub
    Else
        lblStatus.Visible = True
        CsvDatFileName = "C:\BIR_RLF\" & Null2String(Left(COMPANY_TIN, 9)) & Left(EXTRACT_TYPE, 1) & cboMonth.ListIndex + 1 & cboYear.Text
        Set oXLApp = CreateObject("Excel.Application")
        Set oXLBook = oXLApp.Workbooks.Open(sFilePath & "Purchases.xlsx")
    End If
    
    Set oXLSheet = oXLBook.Worksheets(1)
    
    'oXLSheet.UsedRange.Clear
    
    If Function_Access(LOGID, "Acess_Process", "EXTRACT ENTRIES TO BIR RELIEF") = False Then Exit Sub
    
    Call ExtractSummary
    'Call ExcelHeaders
    Call OldExcelHeaders
    Call ExtractBirRelief
        
    'oXLSheet.Range("A2").CopyFromRecordset rsExtractBirRelief
    oXLSheet.Range("A15").CopyFromRecordset rsExtractBirRelief
    
    With oXLBook
        On Error Resume Next
        .Worksheets(1).Select
        Range("A1").Select
        ActiveSheet.Paste
        '.SaveAs CsvDatFileName, xlCSV
        .SaveAs CsvDatFileName
        .Close
    End With

    oXLApp.Visible = True
    Set oXLApp = Nothing
    Set oXLBook = Nothing
    Set oXLSheet = Nothing
    lblStatus.Visible = False
            
    'FileCopy CsvDatFileName & ".csv", CsvDatFileName & ".dat"
    'MsgBox "File saved to " & CsvDatFileName & ".dat", vbInformation, "Information"
    
    LogAudit "R", "B.I.R. DATA EXTRACTION", cboMonth & "-" & cboYear
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub ExtractBirRelief()
    Set rsExtractBirRelief = Nothing
    
    If EXTRACT_TYPE = "SALES" Then
        Set rsExtractBirRelief = gconDMIS.Execute("SELECT * FROM AMIS_VW_BIR_SALES WHERE MONTH(TAXABLEMONTH) = " & cboMonth.ListIndex + 1 & " AND YEAR(TAXABLEMONTH) = " & cboYear.Text & " ORDER BY TAXABLEMONTH ASC, REGISTEREDNAME ASC")
    ElseIf EXTRACT_TYPE = "PURCHASES" Then
        Set rsExtractBirRelief = gconDMIS.Execute("SELECT * FROM AMIS_VW_BIR_PURCHASES WHERE MONTH(TAXABLEMONTH) = " & cboMonth.ListIndex + 1 & " AND YEAR(TAXABLEMONTH) = " & cboYear.Text & " ORDER BY TAXABLEMONTH ASC, REGISTEREDNAME ASC")
    End If
End Sub

Sub ExtractSummary()
    Set rsExtractSummary = Nothing
    
    If EXTRACT_TYPE = "SALES" Then
        Set rsExtractSummary = gconDMIS.Execute("SELECT SUM(GROSSTAXABLEPURCHASE) AS GROSS, SUM(INPUTTAX) AS INPUT FROM [AMIS_VW_BIR_PURCHASES]")
    ElseIf EXTRACT_TYPE = "PURCHASES" Then
        Set rsExtractSummary = gconDMIS.Execute("SELECT SUM(GROSSTAXABLEPURCHASE) AS GROSS, SUM(INPUTTAX) AS INPUT FROM [AMIS_VW_BIR_PURCHASES]")
    End If
End Sub

Sub ExcelHeaders()
    oXLSheet.Cells(1, "A") = "H"
    
    If EXTRACT_TYPE = "SALES" Then
        oXLSheet.Cells(1, "B") = "S"
    ElseIf EXTRACT_TYPE = "PURCHASES" Then
        oXLSheet.Cells(1, "B") = "P"
    End If
    
    oXLSheet.Cells(1, "C") = COMPANY_TIN
    oXLSheet.Cells(1, "D") = COMPANY_NAME
    oXLSheet.Cells(1, "E") = ""
    oXLSheet.Cells(1, "F") = ""
    oXLSheet.Cells(1, "G") = ""
    oXLSheet.Cells(1, "H") = COMPANY_NAME
    oXLSheet.Cells(1, "I") = Null2String(Replace(COMPANY_ADDRESS, ",", " "))
    oXLSheet.Cells(1, "J") = "NAGA CITY CAMARINES SUR 4400" 'ZIPCODE
    oXLSheet.Cells(1, "K") = "0"
    oXLSheet.Cells(1, "L") = "0"
    oXLSheet.Cells(1, "M") = "0"
    oXLSheet.Cells(1, "N") = "0"
    
    oXLSheet.Cells(1, "O") = rsExtractSummary!GROSS 'GROSS
    oXLSheet.Cells(1, "P") = rsExtractSummary!Input 'INPUT
    oXLSheet.Cells(1, "Q") = rsExtractSummary!Input 'INPUT
    
    oXLSheet.Cells(1, "R") = "0"
    oXLSheet.Cells(1, "S") = "065" 'BIR CODE FOR NAGA
    oXLSheet.Cells(1, "T") = LOGDATE
    oXLSheet.Cells(1, "U") = "12" 'UNKNOWN FIELD
    
    oXLSheet.Columns("A:A").EntireColumn.AutoFit
    oXLSheet.Columns("B:B").EntireColumn.AutoFit
    oXLSheet.Columns("C:C").EntireColumn.AutoFit
    oXLSheet.Columns("D:D").EntireColumn.AutoFit
    oXLSheet.Columns("E:E").EntireColumn.AutoFit
    oXLSheet.Columns("F:F").EntireColumn.AutoFit
    oXLSheet.Columns("G:G").EntireColumn.AutoFit
    oXLSheet.Columns("H:H").EntireColumn.AutoFit
    oXLSheet.Columns("I:I").EntireColumn.AutoFit
    oXLSheet.Columns("J:J").EntireColumn.AutoFit
    oXLSheet.Columns("K:K").EntireColumn.AutoFit
    oXLSheet.Columns("L:L").EntireColumn.AutoFit
    oXLSheet.Columns("M:M").EntireColumn.AutoFit
    oXLSheet.Columns("N:N").EntireColumn.AutoFit
    oXLSheet.Columns("O:O").EntireColumn.AutoFit
    oXLSheet.Columns("P:P").EntireColumn.AutoFit
    oXLSheet.Columns("Q:Q").EntireColumn.AutoFit
    oXLSheet.Columns("R:R").EntireColumn.AutoFit
    oXLSheet.Columns("S:S").EntireColumn.AutoFit
    oXLSheet.Columns("T:T").EntireColumn.AutoFit
    oXLSheet.Columns("U:U").EntireColumn.AutoFit
End Sub

Sub OldExcelHeaders()
    oXLSheet.Cells(1, "A") = "PURCHASE TRANSACTION"
    oXLSheet.Cells(1, "A").Font.Bold = True
    oXLSheet.Cells(2, "A") = "RECONCILIATION OF LISTING FOR ENFORCEMENT"
    oXLSheet.Cells(2, "A").Font.Bold = True
    oXLSheet.Cells(6, "A") = "TIN : " & COMPANY_TIN
    oXLSheet.Cells(6, "A").Font.Bold = True
    oXLSheet.Cells(7, "A") = "OWNER'S NAME: " & COMPANY_NAME
    oXLSheet.Cells(7, "A").Font.Bold = True
    oXLSheet.Cells(8, "A") = "OWNER'S TRADE NAME : " & COMPANY_NAME
    oXLSheet.Cells(8, "A").Font.Bold = True
    oXLSheet.Cells(9, "A") = "OWNER'S ADDRESS: ROXAS AVE. DIVERSION RD CONCEPCION PEQUENA NAGA CITY CAMARINES SUR 4400"
    oXLSheet.Cells(9, "A").Font.Bold = True

    oXLSheet.Cells(11, "A") = "TAXABLE"
    oXLSheet.Cells(11, "A").Font.Bold = True
    oXLSheet.Cells(12, "A") = "MONTH"
    oXLSheet.Cells(12, "A").Font.Bold = True
    oXLSheet.Cells(14, "A") = "'(1)"
    oXLSheet.Cells(14, "A").Font.Bold = True

    oXLSheet.Cells(11, "B") = "TAXPAYER"
    oXLSheet.Cells(11, "B").Font.Bold = True
    oXLSheet.Cells(12, "B") = "IDENTIFICATION"
    oXLSheet.Cells(12, "B").Font.Bold = True
    oXLSheet.Cells(13, "B") = "NUMBER"
    oXLSheet.Cells(13, "B").Font.Bold = True
    oXLSheet.Cells(14, "B") = "'(2)"
    oXLSheet.Cells(14, "B").Font.Bold = True
    
    oXLSheet.Cells(11, "C") = "REGISTERED NAME"
    oXLSheet.Cells(11, "C").Font.Bold = True
    oXLSheet.Cells(12, "C") = ""
    oXLSheet.Cells(12, "C").Font.Bold = True
    oXLSheet.Cells(13, "C") = ""
    oXLSheet.Cells(13, "C").Font.Bold = True
    oXLSheet.Cells(14, "C") = "'(3)"
    oXLSheet.Cells(14, "C").Font.Bold = True
    
    oXLSheet.Cells(11, "D") = "NAME OF SUPPLIER"
    oXLSheet.Cells(11, "D").Font.Bold = True
    oXLSheet.Cells(12, "D") = "(Last Name, First Name, Middle Name)"
    oXLSheet.Cells(12, "D").Font.Bold = True
    oXLSheet.Cells(13, "D") = ""
    oXLSheet.Cells(13, "D").Font.Bold = True
    oXLSheet.Cells(14, "D") = "'(4)"
    oXLSheet.Cells(14, "D").Font.Bold = True
    
    oXLSheet.Cells(11, "E") = "SUPPLIER'S ADDRESS"
    oXLSheet.Cells(11, "E").Font.Bold = True
    oXLSheet.Cells(12, "E") = ""
    oXLSheet.Cells(12, "E").Font.Bold = True
    oXLSheet.Cells(13, "E") = ""
    oXLSheet.Cells(13, "E").Font.Bold = True
    oXLSheet.Cells(14, "E") = "'(5)"
    oXLSheet.Cells(14, "E").Font.Bold = True
    
    oXLSheet.Cells(11, "F") = "AMOUNT OF"
    oXLSheet.Cells(11, "F").Font.Bold = True
    oXLSheet.Cells(12, "F") = "GROSS PURCHASE"
    oXLSheet.Cells(12, "F").Font.Bold = True
    oXLSheet.Cells(13, "F") = ""
    oXLSheet.Cells(13, "F").Font.Bold = True
    oXLSheet.Cells(14, "F") = "'(6)"
    oXLSheet.Cells(14, "F").Font.Bold = True
    
    oXLSheet.Cells(11, "G") = "AMOUNT OF"
    oXLSheet.Cells(11, "G").Font.Bold = True
    oXLSheet.Cells(12, "G") = "EXEMPT PURCHASE"
    oXLSheet.Cells(12, "G").Font.Bold = True
    oXLSheet.Cells(13, "G") = ""
    oXLSheet.Cells(13, "G").Font.Bold = True
    oXLSheet.Cells(14, "G") = "'(7)"
    oXLSheet.Cells(14, "G").Font.Bold = True
    
    oXLSheet.Cells(11, "H") = "AMOUNT OF"
    oXLSheet.Cells(11, "H").Font.Bold = True
    oXLSheet.Cells(12, "H") = "ZERO-RATED PURCHASE"
    oXLSheet.Cells(12, "H").Font.Bold = True
    oXLSheet.Cells(13, "H") = ""
    oXLSheet.Cells(13, "H").Font.Bold = True
    oXLSheet.Cells(14, "H") = "'(8)"
    oXLSheet.Cells(14, "H").Font.Bold = True
    
    oXLSheet.Cells(11, "I") = "AMOUNT OF"
    oXLSheet.Cells(11, "I").Font.Bold = True
    oXLSheet.Cells(12, "I") = "TAXABLE PURCHASE"
    oXLSheet.Cells(12, "I").Font.Bold = True
    oXLSheet.Cells(13, "I") = ""
    oXLSheet.Cells(13, "I").Font.Bold = True
    oXLSheet.Cells(14, "I") = "'(9)"
    oXLSheet.Cells(14, "I").Font.Bold = True
    
    oXLSheet.Cells(11, "J") = "AMOUNT OF"
    oXLSheet.Cells(11, "J").Font.Bold = True
    oXLSheet.Cells(12, "J") = "PURCHASE OF SERVICES"
    oXLSheet.Cells(12, "J").Font.Bold = True
    oXLSheet.Cells(13, "J") = ""
    oXLSheet.Cells(13, "J").Font.Bold = True
    oXLSheet.Cells(14, "J") = "'(10)"
    oXLSheet.Cells(14, "J").Font.Bold = True
    
    oXLSheet.Cells(11, "K") = "AMOUNT OF"
    oXLSheet.Cells(11, "K").Font.Bold = True
    oXLSheet.Cells(12, "K") = "PURCHASE OF CAPITAL GOODS"
    oXLSheet.Cells(12, "K").Font.Bold = True
    oXLSheet.Cells(13, "K") = ""
    oXLSheet.Cells(13, "K").Font.Bold = True
    oXLSheet.Cells(14, "K") = "'(11)"
    oXLSheet.Cells(14, "K").Font.Bold = True
    
    oXLSheet.Cells(11, "L") = "AMOUNT OF"
    oXLSheet.Cells(11, "L").Font.Bold = True
    oXLSheet.Cells(12, "L") = "PURCHASE OF GOODS OTHER THAN CAPITAL GOODS"
    oXLSheet.Cells(12, "L").Font.Bold = True
    oXLSheet.Cells(13, "L") = ""
    oXLSheet.Cells(13, "L").Font.Bold = True
    oXLSheet.Cells(14, "L") = "'(12)"
    oXLSheet.Cells(14, "L").Font.Bold = True
    
    oXLSheet.Cells(11, "M") = "AMOUNT OF"
    oXLSheet.Cells(11, "M").Font.Bold = True
    oXLSheet.Cells(12, "M") = "INPUT TAX"
    oXLSheet.Cells(12, "M").Font.Bold = True
    oXLSheet.Cells(13, "M") = ""
    oXLSheet.Cells(13, "M").Font.Bold = True
    oXLSheet.Cells(14, "M") = "'(13)"
    oXLSheet.Cells(14, "M").Font.Bold = True
    
    oXLSheet.Cells(11, "N") = "AMOUNT OF"
    oXLSheet.Cells(11, "N").Font.Bold = True
    oXLSheet.Cells(12, "N") = "GROSS TAXABLE"
    oXLSheet.Cells(12, "N").Font.Bold = True
    oXLSheet.Cells(13, "N") = " PURCHASE"
    oXLSheet.Cells(13, "N").Font.Bold = True
    oXLSheet.Cells(14, "N") = "'(14)"
    oXLSheet.Cells(14, "N").Font.Bold = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    
    fillcbomonth cboMonth
    FillcboNewYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)

    lblType.Caption = EXTRACT_TYPE
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    MsgBox err.Number & vbCrLf & err.DESCRIPTION, vbCritical, "Database Connection Error!"
    Unload frmSplash
    cmdCheck.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gconBIR_RELIEF.Close
End Sub


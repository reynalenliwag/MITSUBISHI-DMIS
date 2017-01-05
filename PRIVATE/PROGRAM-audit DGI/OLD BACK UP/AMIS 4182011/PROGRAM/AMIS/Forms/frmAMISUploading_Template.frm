VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAMISUploading_Template 
   Caption         =   "Opening Balance Uploading Tool"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3150
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select File to Upload"
      Height          =   645
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   2745
   End
   Begin MSComctlLib.ListView lvExcelData 
      Height          =   6825
      Left            =   60
      TabIndex        =   1
      Top             =   1770
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   12039
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Vendor Name"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Invoice No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Invoice Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Reference No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Payment Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Particulars"
         Object.Width           =   9596
      EndProperty
   End
   Begin VB.Label lblTotal 
      Height          =   315
      Left            =   7080
      TabIndex        =   6
      Top             =   8670
      Width           =   1905
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1650
      TabIndex        =   5
      Top             =   1290
      Width           =   2805
   End
   Begin VB.Label lblAcctCode 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1650
      TabIndex        =   4
      Top             =   870
      Width           =   2805
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPTION:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1410
      Width           =   1140
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNT CODE:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   1320
   End
End
Attribute VB_Name = "frmAMISUploading_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSelect_Click()
'On Error GoTo ErrorHandler

    Dim xlApplication           As Excel.Application
    Dim xlWorkbook              As Excel.Workbook
    Dim xlWorksheet             As Excel.Worksheet
    Dim var                     As Variant
    Dim intRowCounter       As Integer
    Dim xList As ListItem
    Dim Rng As Range, ix As Long
    Dim xTotal As Double
    
    Set xlApplication = New Excel.Application
'    Set xlWorkbook = xlApplication.Workbooks.Open(App.Path & "\Sample.xls")
    
'    Set xlWorksheet = xlWorkbook.Worksheets("Sheet1")
'    var = xlWorksheet.Range("L2").Value
'
'    lblTranNo.Caption = var
    
    
    
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = _
        "Select Excel file to import from and click open..."
    CommonDialog1.Filter = _
        "MS Excel Files (*.xls)|*.xls|(*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    xExcelFile = CommonDialog1.FileName
    
'    If CommonDialog1.CancelError = False Then
'        Exit Sub
'    End If
    Set xlWorkbook = xlApplication.Workbooks.Open(xExcelFile)
    Connect
    With xlApplication
        .Workbooks(1).Worksheets(1).Select
        lblAcctCode.Caption = .Range("B1").Value
        lblDescription.Caption = .Range("B2").Value
'        For intRowCounter = 5 To 484
        For intRowCounter = 5 To CInt(.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row)
            '.ActiveSheet.UsedRange.Rows.Count
            '.ActiveSheet.Range("B1").End(xlDown).Row
            
            'CInt(.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row)
'            If StrComp(.Cells(intRowCounter, 2), "") <> "END" Then
'                FillGrid
'            End If
             '   .ActiveSheet.Cells.SpecialCells(xlCellTypeBlanks).Delete (xlShiftUp)
             
'            Set Rng = Intersect(Range("B" & intRowCounter), ActiveSheet.UsedRange)
'            For ix = Rng.Count To 1 Step -1
'                If Trim(Replace(Rng.Item(ix).Text, Chr(160), Chr(32))) = "" Then
'                  'Rng.Item(ix).EntireRow.Delete
'                  GoTo NextRecord
'                End If
            
            Set xList = lvExcelData.ListItems.Add(, , .Range("A" & intRowCounter).Value)
                xList.SubItems(1) = .Range("B" & intRowCounter).Value
                xList.SubItems(2) = .Range("C" & intRowCounter).Value
                xList.SubItems(3) = .Range("D" & intRowCounter).Value
                xList.SubItems(4) = .Range("E" & intRowCounter).Value
                xList.SubItems(5) = .Range("F" & intRowCounter).Value
                xList.SubItems(6) = .Range("G" & intRowCounter).Value
                xTotal = xTotal + xList.SubItems(5)
'NextRecord:
'            Next
        Next intRowCounter
    End With
    
    xlWorkbook.Close
    xlApplication.Quit
    
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing
    
'ErrorHandler:
'    If Err.Number = 429 Then
'        MsgBox "You cannot use this feature unless you have Microsoft Excel installed."
'        Exit Sub
'    Else
'
'        Exit Sub
'    End If
End Sub

Sub FillGrid()
    Dim xList As ListItem
    Dim rsImportData As ADODB.Recordset
    Set rsImportData = New ADODB.Recordset
    rsImportData.Open "SELECT VENDOR_NAME,INVOICE_NO,INVOICE_DATE,REFERENCE_NO,PAYMENT_TYPE,AMOUNT,REMARKS FROM [Sheet1$]", , adOpenForwardOnly
    If Not rsImportData.EOF And Not rsImportData.BOF Then
        Do While Not rsImportData.EOF
            Set xList = lvExcelData.ListItems.Add(, , rsImportData!Transno)
            xList.SubItems(1) = rsImportData!Type
            xList.SubItems(2) = rsImportData!Date
            xList.SubItems(3) = rsImportData!Num
            xList.SubItems(4) = rsImportData!Name
            xList.SubItems(5) = rsImportData!Memo
            xList.SubItems(6) = rsImportData!Account
            xList.SubItems(7) = NumericVal(rsImportData!Debit)
            xList.SubItems(8) = NumericVal(rsImportData!Credit)
            rsImportData.MoveNext
        Loop
    End If
    Set rsImportData = Nothing
End Sub

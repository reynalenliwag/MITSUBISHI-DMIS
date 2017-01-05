VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Begin VB.Form frmAPJImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Payable Import Process"
   ClientHeight    =   7680
   ClientLeft      =   345
   ClientTop       =   1110
   ClientWidth     =   14055
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "APJImport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   14055
   Begin VB.PictureBox wizFlexCrack1 
      Height          =   3765
      Left            =   1470
      ScaleHeight     =   3705
      ScaleWidth      =   6375
      TabIndex        =   9
      Top             =   7800
      Width           =   6435
   End
   Begin VB.CommandButton cmdClearJournals 
      BackColor       =   &H0080FF80&
      Caption         =   "Clear Selected Date"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   1935
   End
   Begin VB.CommandButton cmdShowTrans 
      Caption         =   "Show Transactions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3630
      MouseIcon       =   "APJImport.frx":030A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Process Import of SALES"
      Top             =   90
      Width           =   2010
   End
   Begin MSComCtl2.DTPicker dtpTranDate 
      Height          =   405
      Left            =   1770
      TabIndex        =   10
      Top             =   90
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   54067201
      CurrentDate     =   38216
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   9480
      TabIndex        =   11
      Top             =   6480
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   556
      Picture         =   "APJImport.frx":045C
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "APJImport.frx":0478
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
   Begin FlexCell.Grid Grid1 
      Height          =   4905
      Left            =   30
      TabIndex        =   12
      Top             =   1170
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin FlexCell.Grid Grid2 
      Height          =   4905
      Left            =   4740
      TabIndex        =   13
      Top             =   1140
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin FlexCell.Grid Grid3 
      Height          =   4905
      Left            =   9390
      TabIndex        =   15
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   13230
      MouseIcon       =   "APJImport.frx":0494
      MousePointer    =   99  'Custom
      Picture         =   "APJImport.frx":05E6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Window"
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Import"
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
      Height          =   795
      Left            =   12510
      MouseIcon       =   "APJImport.frx":094C
      MousePointer    =   99  'Custom
      Picture         =   "APJImport.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Process Importing of Cash Receipts "
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SERVICE SUBLET"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   9360
      TabIndex        =   14
      Top             =   630
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Only Un-Imported Invoices can be Imported"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   150
      TabIndex        =   8
      Top             =   7170
      Width           =   7995
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PARTS, ACCS && MATERIALS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   60
      TabIndex        =   7
      Top             =   630
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VEHICLES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   4710
      TabIndex        =   6
      Top             =   615
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   180
      Width           =   1875
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
      Left            =   9480
      TabIndex        =   4
      Top             =   6180
      Width           =   5835
   End
End
Attribute VB_Name = "frmAPJImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function SetTransaction(XXX As Variant) As String
    Dim rsSBOOKTransaction                                            As ADODB.Recordset
    Set rsSBOOKTransaction = New ADODB.Recordset
    Set rsSBOOKTransaction = gconDMIS.Execute("Select * from SBOOK Where BOOK = 'A' and CODE = '" & XXX & "'")
    If Not rsSBOOKTransaction.EOF And Not rsSBOOKTransaction.BOF Then
        SetTransaction = Null2String(rsSBOOKTransaction!DESCNAME)
    End If
    Set rsSBOOKTransaction = Nothing
End Function

Function SetOtherTransaction(XXX As Variant) As String
    Dim rsSBOOKOtherTransaction                                       As ADODB.Recordset
    Set rsSBOOKOtherTransaction = New ADODB.Recordset
    Set rsSBOOKOtherTransaction = gconDMIS.Execute("Select * from SBOOK Where BOOK = 'D' and CODE = '" & XXX & "'")
    If Not rsSBOOKOtherTransaction.EOF And Not rsSBOOKOtherTransaction.BOF Then
        SetOtherTransaction = Null2String(rsSBOOKOtherTransaction!DESCNAME)
    End If
    Set rsSBOOKOtherTransaction = Nothing
End Function

Function Setacctname(VVV As Variant) As String
    Dim rsChartAccount2                                               As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    If Left(VVV, 1) = "'" Then
        rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & VVV, gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!Description))
    Else
        Setacctname = ""
    End If
End Function

Function GetVoucherNo() As String
    Dim rsJournal_hd                                                  As ADODB.Recordset
    Set rsJournal_hd = New ADODB.Recordset
    Set rsJournal_hd = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'APJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_hd.EOF And Not rsJournal_hd.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_hd!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function CheckAPJExisting(VarInvoiceNo As String, VarTYPE As Variant) As Boolean
    Dim rsCheckAPJ_Journal_HD                                         As ADODB.Recordset
    Set rsCheckAPJ_Journal_HD = New ADODB.Recordset
    If VarTYPE = "PARTS" Then
        Set rsCheckAPJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'APJ' AND InvoiceType = 'PARTS' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    ElseIf VarTYPE = "ACCESSORIES" Then
        Set rsCheckAPJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'APJ' AND InvoiceType = 'ACCESSORIES' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    ElseIf VarTYPE = "MATERIALS" Then
        Set rsCheckAPJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'APJ' AND InvoiceType = 'MATERIALS' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    ElseIf VarTYPE = "VEHICLES" Then
        Set rsCheckAPJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'APJ' AND InvoiceType = 'VEHICLES' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    ElseIf VarTYPE = "" Then
        Set rsCheckAPJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'APJ' AND InvoiceType = NULL AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    End If
    If Not rsCheckAPJ_Journal_HD.EOF And Not rsCheckAPJ_Journal_HD.BOF Then
        CheckAPJExisting = True
    Else
        CheckAPJExisting = False
    End If
    Set rsCheckAPJ_Journal_HD = Nothing
End Function

Function ReturnClearing_AccountCode(XXX As String) As String
    Dim rsChartAccount                                                As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'CLEARING' AND TRANTYPE1 = '" & Trim(XXX) & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnClearing_AccountCode = Null2String(rsChartAccount!acctcode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnAP_AccountCode(XXX As String) As String
    Dim rsChartAccount                                                As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'AP' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAP_AccountCode = Null2String(rsChartAccount!acctcode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnInPutTax()
    Dim rsChartAccount                                                As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'INPUT TAX'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInPutTax = Null2String(rsChartAccount!acctcode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnInComeTax(XXX As String) As String
    Dim rsChartAccount                                                As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'INCOME TAX' AND TRANTYPE2 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInComeTax = Null2String(rsChartAccount!acctcode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnInventory(XXX As String, Optional YYY As String) As String
    Dim rsChartAccount                                                As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If Trim(YYY) = "" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & XXX & "'")
    Else
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & XXX & "' AND TRANTYPE1 = '" & YYY & "'")
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInventory = Null2String(rsChartAccount!acctcode)
    End If
    Set rsChartAccount = Nothing
End Function

Function SetSellingDealerName(XXX As String) As String
    Dim rsSellingDealer                                               As ADODB.Recordset
    Set rsSellingDealer = New ADODB.Recordset
    Set rsSellingDealer = gconDMIS.Execute("Select * from CSMS_SellingDealer Where DealerCode = '" & XXX & "'")
    If Not rsSellingDealer.EOF And Not rsSellingDealer.BOF Then
        SetSellingDealerName = Null2String(rsSellingDealer!DealerName)
    End If
End Function

Function ReturnPartNo(nard As String) As String
    Dim SQL                                                           As String
    Dim rs                                                            As New ADODB.Recordset

    SQL = "SELECT stock_ord from PMIS_Tdaytran where tranno='" & nard & "' and  status ='P'"

    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)

    If Not rs.EOF And Not rs.BOF Then
        ReturnPartNo = Null2String(rs!stock_ord)
    End If
    Set rs = Nothing
End Function

Function CheckIfORIG(ARNIE As String) As Boolean
    Dim SQL                                                           As String
    Dim rs                                                            As New ADODB.Recordset

    SQL = "SELECT GENUINE From PMIS_stockmas where Stockno='" & ARNIE & "'"

    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)

    If Not rs.EOF And Not rs.BOF Then
        If Null2String(rs!genuine) = "Y" Then
            CheckIfORIG = True
        Else
            CheckIfORIG = False
        End If
    End If
    Set rs = Nothing
End Function

Function ReturnCode(XXX As String) As String
    'Update By BTT - 07092008
    Dim SQL                                                           As String
    Dim rs                                                            As New ADODB.Recordset
    Dim MARK                                                          As String

    MARK = (Replace(XXX, " ", ""))

    SQL = "SELECT Code, replace(Nameofvendor,' ','') from ALL_Vendor_table where REPLACE(Nameofvendor,' ','') like '" & MARK & "%'"

    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)

    If Not rs.BOF And Not rs.EOF Then
        ReturnCode = Null2String(rs!code)
    Else
        ReturnCode = ""
    End If
    Set rs = Nothing
End Function

Function CheckSubletifExist(MARK As String, EVAN As String) As Boolean
    'Update By BTT - 07092008
    Dim SQL                                                           As String
    Dim rs                                                            As New ADODB.Recordset

   nard = "SELECT * from AMIS_pv_detail where MRR_no='" & MARK & "' and po_no='" & EVAN & "'"

    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(nard)

    If Not rs.EOF And Not rs.BOF Then
        CheckSubletifExist = True
    Else
        CheckSubletifExist = False
    End If
    Set rs = Nothing
End Function

Function GetVendorTerms(XXX As String) As String
    'Update By BTT - 07092008
    Dim SQL                                                           As String
    Dim rs                                                            As New ADODB.Recordset

    SQL = "SELECT terms from all_vendor_table where code='" & XXX & "'"

    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)

    If Not rs.EOF And Not rs.BOF Then
        GetVendorTerms = Null2String(rs!TERMS)
    Else
        GetVendorTerms = ""
    End If
    Set rs = Nothing
End Function

Sub InitGrids()
    With Grid1
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "Type"
        .Cell(0, 3).Text = "RR No."
        .Cell(0, 4).Text = "RR Amt."
        .Cell(0, 5).Text = "Supplier"

        .Column(0).Width = 10
        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 60
        .Column(4).Width = 80
        .Column(5).Width = 200

        .Column(1).CellType = cellCheckBox
        .Column(4).Alignment = cellRightGeneral

        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True

    End With

    With Grid2
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "Type"
        .Cell(0, 3).Text = "RR No."
        .Cell(0, 4).Text = "RR Amt."
        .Cell(0, 5).Text = "Supplier"

        .Column(0).Width = 10
        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 60
        .Column(4).Width = 80
        .Column(5).Width = 200

        .Column(1).CellType = cellCheckBox
        .Column(4).Alignment = cellRightGeneral

        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True
    End With
    With Grid3
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "Type"
        .Cell(0, 3).Text = "RR No."
        .Cell(0, 4).Text = "RR Amt."
        .Cell(0, 5).Text = "Contractor"

        .Column(0).Width = 10
        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 60
        .Column(4).Width = 80
        .Column(5).Width = 200

        .Column(1).CellType = cellCheckBox
        .Column(4).Alignment = cellRightGeneral

        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True

    End With

End Sub

Sub ImportSublet()
    Dim rsJournal_HDDup                                               As New ADODB.Recordset
    Dim RsSublet                                                      As New ADODB.Recordset
    Dim RsSublet_Det                                                  As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim GridImports                                                   As Integer
    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                                 As String
    Dim J_JNO, J_REMARKS, J_VENDORCODE, J_CUSTOMERCODE                As String
    Dim J_CUSTOMERNAME                                                As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
    Dim J_STATUS, J_JITEMNO, J_CHECKNO                                As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE                           As String
    Dim J_INVOICETYPE, J_INVOICENO                                    As String
    Dim J_CHECKDATE, J_BANKCODE                                       As String
    Dim J_REFNO, J_REFDATE                                            As String
    Dim J_TERMS, J_DEALER, J_ACCT_CODE, J_ACCT_NAME                   As String
    Dim J_ATC, J_RATE, J_TAXBASE                                      As Double
    Dim i                                                             As Integer
    Dim J_PAIDSTATUS, J_RECEIVESTATUS
    J_JTYPE = "'APJ'"
    Dim TOTAL_CREDIT                                                  As Double
    Dim TOTAL_DEBIT                                                   As Double
    Dim TheRO                                                         As String
    Dim ThePO                                                         As String
    Dim TheSublet_Cost                                                As Double
    Dim TheSublet_Vat                                                 As Double
    Dim TheSublet_Net                                                 As Double
    Dim TheRRDate                                                     As String
    Dim TheRRNO                                                       As String
    J_CUSTOMERCODE = "'999999'"
    Dim TheINVOICE_no                                                 As String
    Dim WCode                                                         As String
    Dim TERMS                                                         As String
    TOTAL_CREDIT = 0: TOTAL_DEBIT = 0
    i = 1
    For GridImports = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(i, 1).Text) = 0 Then
            SQL = "SELECT * from CSMS_PO_RC_HD where RC_NO='" & Grid3.Cell(GridImports, 3).Text & "' and status='P' and RC_DATE='" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "'"
            Set RsSublet = New ADODB.Recordset
            Set RsSublet = gconDMIS.Execute(SQL)
            If Not RsSublet.EOF And Not RsSublet.BOF Then
                TheRO = Null2String(RsSublet!ro_no)
                ThePO = Null2String(RsSublet!po_no)
                TheRRNO = Null2String(RsSublet!RC_no)
                TheINVOICE_no = Null2String(RsSublet!Invoice_no)
                J_CUSTOMERNAME = "NULL"
                J_INVOICETYPE = "'SUBLET'"
                'SUBLET DETAIL LOOKUP
                Set RsSublet_Det = New ADODB.Recordset
                Set RsSublet_Det = gconDMIS.Execute("SELECT * FROM CSMS_PO_RC_DT where PO_no='" & ThePO & "'")
                If Not RsSublet_Det.EOF And Not RsSublet_Det.BOF Then
                    J_VENDORCODE = N2Str2Null(ReturnCode(RsSublet_Det!TECHNICIAN))
                    TheSublet_Cost = NumericVal(RsSublet_Det!contractamount)    'COST
                    TheSublet_Vat = NumericVal(TheSublet_Cost) / 1.12 * 0.12
                    TheSublet_Net = NumericVal(TheSublet_Cost) - TheSublet_Vat
                    TheRRDate = Null2String(RsSublet!rc_date)
                    WCode = Null2String(RsSublet_Det!WCode)
                End If
                TERMS = GetVendorTerms(ReturnCode(RsSublet_Det!TECHNICIAN))
                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If


                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_DEBIT = 0: J_CREDIT = 0: J_TAX = 0: J_OUTBALANCE = 0
                J_INVOICEDATE = N2Date2Null(RsSublet!rc_date): J_BALANCE = 0: J_AMOUNTPAID = 0
                J_DUEDATE = N2Date2Null(TheRRDate)
                J_PAYTYPE = "'" & TERMS & "D'": J_STATUS = "'N'"
                J_TERMS = "'" & TERMS & "D'": J_DEALER = "NULL"
                J_CHECKDATE = "NULL": J_BANKCODE = "NULL"
                J_INVOICEAMT = NumericVal(TheSublet_Cost)
                J_PAIDSTATUS = "'N'": J_RECEIVESTATUS = "'N'"
                J_CHECKNO = "NULL": J_REFDATE = "NULL"
                J_AMOUNTTOPAY = Round(NumericVal(TheSublet_Cost), 2)
                J_JDATE = N2Date2Null(TheRRDate)
                J_REFNO = N2Str2Null(TheRRNO)
                J_INVOICENO = N2Str2Null(TheINVOICE_no)

                'AP
                J_REMARKS = N2Str2Null("To Record Sublet Recieving with RR No:" + TheRRNO + " (And Ro No " + TheRO + ")")
                J_JITEMNO = "'0001'"
                If COMPANY_CODE = "HGC" Then
                    J_ACCT_CODE = "'21-01002-00'"
                    J_ACCT_NAME = N2Str2Null(Setacctname("'21-01002-00'"))

                Else    ' HSB
                    J_ACCT_CODE = "'21-01004-20'"
                    J_ACCT_NAME = N2Str2Null(Setacctname("'21-01004-20'"))
                End If
                J_DEBIT = 0
                'J_CREDIT = Round(NumericVal(TheSublet_Cost) / 1.12, 2)
                J_CREDIT = Round(NumericVal(TheSublet_Cost))
                J_TAX = 0
                J_ATC = 0
                J_RATE = 0
                J_TAXBASE = 0
                J_GROSS = 0
                J_NET = 0
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                'SUBLET DETAIL
                
                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                gconDMIS.Execute SQL_STATEMENT
                
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                
                'INVENTORY
                J_JITEMNO = "'0002'"
                If COMPANY_CODE = "HGC" Then
                    J_ACCT_CODE = "'11-05006-00'"
                    J_ACCT_NAME = N2Str2Null(Setacctname("'11-05006-00'"))
                Else    ' HSB
                    J_ACCT_CODE = "'11-05006-21'"
                    J_ACCT_NAME = N2Str2Null(Setacctname("'11-05006-21'"))
                End If
                'J_DEBIT = Round(NumericVal(TheSublet_Cost) / 1.12, 2) - TheSublet_Vat
                J_DEBIT = Round(NumericVal(TheSublet_Cost)) - TheSublet_Vat
                J_CREDIT = 0: J_TAX = 0: J_ATC = 0
                J_RATE = 0: J_TAXBASE = 0
                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                'SUBLET DETAIL
                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                
                gconDMIS.Execute SQL_STATEMENT
                
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)

                'TAX
                If WCode = "C" Then
                    ' Do Nothing
                Else
                    J_JITEMNO = "'0003'"
                    J_ACCT_CODE = N2Str2Null(ReturnInPutTax())
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInPutTax()))
                    J_DEBIT = Round(NumericVal(Round((TheSublet_Cost / 1.12), 2) * 0.12), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    J_GROSS = 0
                    J_NET = 0
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)

                End If

                'SUBLET DETAIL
                J_JVOUCHERNO = J_VOUCHERNO
                PV_ITEMNO = N2Str2Null("0001")
                PV_MRRNO = N2Str2Null(TheRRNO)
                PV_PONO = N2Str2Null(ThePO)
                PV_INVNO = N2Str2Null(TheINVOICE_no)
                PV_PRODNO = "NULL"
                PV_AMOUNT = Round(NumericVal(TheSublet_Cost), 2)
                PV_STATUS = "'N'"

                SQL_STATEMENT = "insert into AMIS_PV_Detail " & _
                                 "(VoucherNo,JDATE,itemno,PO_No,MRR_No,INV_No,PROD_No,Amount,status)" & _
                               " values (" & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & ", " & PV_PONO & _
                                 ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                 ", " & PV_STATUS & ")"
                
                gconDMIS.Execute SQL_STATEMENT

                'TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_PV_Detail", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                
                
                'SUBLET HEADER
                gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                               " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
            
            End If

        End If
        Grid3.Cell(i, 1).Text = 1
        i = i + 1
    Next
End Sub

Private Sub cmdCheck_Click()

    'On Error GoTo ErrorCode:
    If Function_Access(LOGID, "Acess_Process", "IMPORT PURCHASES") = False Then Exit Sub

    Screen.MousePointer = 11
    'HEADER
    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                                 As String
    Dim J_JNO, J_REMARKS, J_VENDORCODE, J_CUSTOMERCODE                As String
    Dim J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
    Dim J_CHECKNO                                                     As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE                           As String
    Dim J_INVOICETYPE, J_INVOICENO                                    As String
    Dim J_CHECKDATE, J_BANKCODE                                       As String
    Dim J_REFNO, J_REFDATE                                            As String
    Dim J_TERMS, J_DEALER                                             As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS                                 As String

    'DETAIL
    Dim J_ACCT_CODE, J_ACCT_NAME                                      As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET                      As Double
    Dim J_STATUS, J_JITEMNO                                           As String

    Dim rsJournal_HDDup                                               As ADODB.Recordset

    Dim PMIOS_RRNO                                                    As String
    Dim PMIOS_RRDATE                                                  As String
    Dim PMIOS_PONO                                                    As String
    Dim PMIOS_PODATE                                                  As String
    Dim PMIOS_RECVD_CODE                                              As String
    Dim PMIOS_RECVD_FROM                                              As String
    Dim PMIOS_DRNO                                                    As String
    Dim PMIOS_INVNO                                                   As String
    Dim PMIOS_CLASSCODE                                               As String
    Dim PMIOS_TERMS                                                   As String
    Dim PMIOS_TOTALQTY                                                As Double
    Dim PMIOS_TTLRRAMT                                                As Double
    Dim PMIOS_DS1                                                     As Double
    Dim PMIOS_DS_AMT1                                                 As Double
    Dim PMIOS_NETRRAMT                                                As Double
    Dim PMIOS_STATUS                                                  As String
    Dim PMIOS_TYPE                                                    As String
    Dim CONDUCTION                                                    As String
    Dim AMIS_JTYPE                                                    As String

    Dim TOTAL_DEBIT, TOTAL_CREDIT                                     As Double

    Dim i                                                             As Long


    Dim PV_PONO, PV_MRRNO, PV_INVNO, PV_PRODNO                        As String
    Dim J_JVOUCHERNO                                                  As String
    Dim PV_AMOUNT                                                     As Double
    Dim PV_STATUS, PV_ITEMNO                                          As String


    Dim rsRR_HD                                                       As ADODB.Recordset

    Dim GridImport                                                    As Integer
    i = 0
    For GridImport = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(GridImport, 1).Text) = 0 Then
            Set rsRR_HD = New ADODB.Recordset
            ' Update By BTT : 08132008
            If COMPANY_CODE = "HGC" Then
                Set rsRR_HD = gconDMIS.Execute("Select * from PMIS_vw_RR_TRANS Where RRNO = '" & Grid1.Cell(GridImport, 3).Text & "' AND (CLASSCODE = 'PCG' or CLASSCODE = 'PCS') AND RRDATE = '" & CDate(dtpTranDate) & "' Order by RRNO ASC")
            Else
                Set rsRR_HD = gconDMIS.Execute("Select * from PMIS_vw_RR_TRANS Where RRNO = '" & Grid1.Cell(GridImport, 3).Text & "' AND CLASSCODE = 'PCG' AND RRDATE = '" & CDate(dtpTranDate) & "' Order by RRNO ASC")
            End If
            If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                PMIOS_RRNO = Null2String(rsRR_HD!RRNO)
                PMIOS_RRDATE = Null2String(rsRR_HD!RRDATE)
                PMIOS_PONO = Null2String(rsRR_HD!PONO)
                PMIOS_PODATE = Null2String(rsRR_HD!PODATE)
                PMIOS_RECVD_CODE = Null2String(rsRR_HD!RECVD_CODE)
                PMIOS_RECVD_FROM = Null2String(rsRR_HD!RECVD_FROM)
                PMIOS_DRNO = Null2String(rsRR_HD!DRNO)
                PMIOS_INVNO = Null2String(rsRR_HD!INVNO)
                PMIOS_CLASSCODE = Null2String(rsRR_HD!CLASSCODE)
                PMIOS_TERMS = Null2String(rsRR_HD!TERMS)
                PMIOS_TOTALQTY = Round(N2Str2Zero(rsRR_HD!TOTALQTY), 2)
                PMIOS_TTLRRAMT = Round(N2Str2Zero(rsRR_HD!TTLRRAMT), 2)
                PMIOS_DS1 = Round(N2Str2Zero(rsRR_HD!DS1), 2)
                PMIOS_DS_AMT1 = Round(N2Str2Zero(rsRR_HD!DS_AMT1), 2)
                PMIOS_NETRRAMT = Round(N2Str2Zero(rsRR_HD!NETRRAMT), 2)
                PMIOS_STATUS = Null2String(rsRR_HD!Status)
                PMIOS_TYPE = Null2String(rsRR_HD!Type)
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0

                

                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If

                J_JDATE = N2Date2Null(PMIOS_RRDATE)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_JTYPE = "'APJ'"
                If PMIOS_TYPE = "P" Then
                    AMIS_JTYPE = "PI"
                    J_REMARKS = "'To record spareparts purchases. with Ref# " & PMIOS_RRNO & "'"
                End If
                If PMIOS_TYPE = "A" Then
                    AMIS_JTYPE = "AI"
                    J_REMARKS = "'To record accessories purchases. with Ref# " & PMIOS_RRNO & "'"
                End If
                If PMIOS_TYPE = "M" Then
                    AMIS_JTYPE = "MI"
                    J_REMARKS = "'To record materials purchases. with Ref# " & PMIOS_RRNO & "'"
                End If
                J_VENDORCODE = N2Str2Null(PMIOS_RECVD_CODE)
                J_CUSTOMERCODE = "'999999'"

                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                J_AMOUNTTOPAY = Round(NumericVal(PMIOS_NETRRAMT), 2)
                J_INVOICEAMT = 0
                J_BALANCE = Round(NumericVal(PMIOS_NETRRAMT), 2)
                J_AMOUNTPAID = 0

                J_STATUS = "'N'"

                J_INVOICEDATE = N2Date2Null(PMIOS_RRDATE)
                J_INVOICENO = N2Str2Null(PMIOS_RRNO)
                J_CHECKNO = "NULL"
                J_DUEDATE = N2Date2Null(Format(DateAdd("d", NumericVal(PMIOS_TERMS), Format(PMIOS_RRDATE, "DD-MMM-YY"))))
                J_PAYTYPE = ("'" & PMIOS_TERMS & "D'")

                If PMIOS_TYPE = "P" Then
                    J_INVOICETYPE = "'PARTS'"
                ElseIf PMIOS_TYPE = "A" Then
                    J_INVOICETYPE = "'ACCESSORIES'"
                ElseIf PMIOS_TYPE = "M" Then
                    J_INVOICETYPE = "'MATERIALS'"
                Else
                    J_INVOICETYPE = "NULL"
                End If
                J_CHECKDATE = "NULL"
                J_BANKCODE = "NULL"
                J_REFNO = "NULL"
                J_REFDATE = "NULL"
                J_TERMS = ("'" & PMIOS_TERMS & "D'")
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"



                'CASH ON HAND
                If PMIOS_NETRRAMT > 0 Then
                    J_JITEMNO = "'0001'"
                    If PMIOS_TYPE = "P" Then
                        'Update By BTT: 07042008 to separate the Orig to not Orig
                        If COMPANY_CODE = "HGC" Then
                            If CheckIfORIG(ReturnPartNo(PMIOS_RRNO)) = True Then
                                'Original Parts
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            Else
                                'Non - Original Parts
                                J_ACCT_CODE = "'11-05001-00'"
                                J_ACCT_NAME = N2Str2Null(Setacctname("'11-05001-00'"))
                            End If
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                        End If
                    End If
                    If PMIOS_TYPE = "A" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA")))
                    End If
                    If PMIOS_TYPE = "M" Then
                        If COMPANY_CODE = "HSB" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIAL", "INVA"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIAL", "INVA")))

                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVM"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVM")))
                        End If
                    End If

                    'WITHOUT INPUT TAX
                    'J_DEBIT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                    'WITH INPUT TAX
                    J_DEBIT = Round(NumericVal(PMIOS_NETRRAMT - PMIOS_DS_AMT1), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                    
                    'NO INPUT TAX - HGC
                    J_JITEMNO = "'0002'"
                    J_ACCT_CODE = N2Str2Null(ReturnInPutTax())
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInPutTax()))
                    J_DEBIT = NumericVal(PMIOS_DS_AMT1)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                    
                    '                    J_JITEMNO = "'0003'"
                    '                    J_ACCT_CODE = N2Str2Null(ReturnInComeTax("EXPANDED"))
                    '                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInComeTax("EXPANDED")))
                    '                    J_DEBIT = 0
                    '                    J_CREDIT = NumericVal(PMIOS_NETRRAMT) * 0.01
                    '                    J_TAX = 0
                    '                    J_GROSS = 0
                    '                    J_NET = 0
                    '                    J_STATUS = "'N'"
                    '                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '
                    '                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         '                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                         '                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         '                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         '                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    J_JITEMNO = "'0003'"
                    'AP IS CLEARING ACCOUNT
                    If COMPANY_CODE = "HGC" Then
                        If PMIOS_RECVD_CODE = "H00001" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("HYUNDAI"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("HYUNDAI")))
                        Else
                            J_ACCT_CODE = "'21-01002-00'"
                            J_ACCT_NAME = N2Str2Null(Setacctname("'21-01002-00'"))
                        End If
                    ElseIf COMPANY_CODE = "HMH" Then
                        If PMIOS_RECVD_CODE = "H00001" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("HYUNDAI"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("HYUNDAI")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("GENERAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("GENERAL")))
                        End If
                    Else
                        If PMIOS_RECVD_CODE = "H00001" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("HYUNDAI"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("HYUNDAI")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("INVP"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("INVP")))

                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("GENERAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("GENERAL")))
                        End If
                    End If
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("AP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("AP")))
                    End If

                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                    
                End If

                J_JVOUCHERNO = J_VOUCHERNO
                PV_ITEMNO = N2Str2Null("0001")
                PV_MRRNO = N2Str2Null(PMIOS_RRNO)
                PV_PONO = N2Str2Null(PMIOS_PONO)
                PV_INVNO = N2Str2Null(PMIOS_INVNO)
                PV_PRODNO = "NULL"
                PV_AMOUNT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                PV_STATUS = "'N'"

                SQL_STATEMENT = "insert into AMIS_PV_Detail " & _
                                 "(VoucherNo,JTYPE,JDATE,itemno,PO_No,MRR_No,INV_No,PROD_No,Amount,status)" & _
                               " values (" & J_JVOUCHERNO & ",'" & AMIS_JTYPE & "'," & J_JDATE & ", " & PV_ITEMNO & ", " & PV_PONO & _
                                 ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                 ", " & PV_STATUS & ")"
                
                gconDMIS.Execute SQL_STATEMENT

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_PV_Detail", "BERNARD", N2Str2Null(AMIS_JTYPE), "Jtype"))
                NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), AMIS_JTYPE, N2Str2Null(PV_MRRNO)
                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                               " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                'SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                '               " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                '               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & ",'" & AMIS_JTYPE & "'," & PV_INVNO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & PV_MRRNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                '                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                gconDMIS.Execute SQL_STATEMENT
                
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                Grid1.Cell(GridImport, 1).Text = 1
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid1.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next


    i = 0
    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then
            Set rsRR_HD = New ADODB.Recordset
            Set rsRR_HD = gconDMIS.Execute("Select * from SMIS_MRRINV Where CODE = '" & Grid2.Cell(GridImport, 3).Text & "' AND DateReceived = '" & CDate(dtpTranDate) & "' Order by DateReceived ASC")

            If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                PMIOS_RRNO = Null2String(rsRR_HD!code)
                PMIOS_RRDATE = Null2String(rsRR_HD!DateReceived)
                PMIOS_PONO = Null2String(rsRR_HD!PONO)
                PMIOS_PODATE = Null2String(rsRR_HD!PullOutDate)
                PMIOS_RECVD_CODE = Null2String("H00001")
                PMIOS_RECVD_FROM = Null2String("HYUNDAI ASIA RESOURCES INC.")
                'PMIOS_RECVD_CODE = Null2String(rsRR_HD!Source)
                'PMIOS_RECVD_FROM = SetSellingDealerName(Null2String(rsRR_HD!Source))
                PMIOS_DRNO = Null2String(rsRR_HD!DRNO)
                'PMIOS_INVNO = Null2String(rsRR_HD!VI_NO)
                PMIOS_INVNO = Null2String(rsRR_HD!refpono)
                PMIOS_CLASSCODE = Null2String(rsRR_HD!MODEL)
                CONDUCTION = Null2String(rsRR_HD!ignkey)
                PMIOS_TERMS = "CSH"
                PMIOS_TOTALQTY = 1

                
                If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HAS" Then
                    'Update By BTT - 06272008 : Net of subsidy
                    PMIOS_TTLRRAMT = Round(N2Str2Zero(rsRR_HD!PurchPrice), 2) - Round(N2Str2Zero(rsRR_HD!mmpcsubs), 2)
                    PMIOS_DS_AMT1 = Round(((N2Str2Zero(rsRR_HD!PurchPrice) - Round(N2Str2Zero(rsRR_HD!mmpcsubs), 2)) / 1.12) * 0.12, 2)
                    PMIOS_NETRRAMT = Round((N2Str2Zero(rsRR_HD!PurchPrice) - Round(N2Str2Zero(rsRR_HD!mmpcsubs), 2)) - PMIOS_DS_AMT1, 2)
                    PMIOS_DS1 = 12
                Else
                    PMIOS_DS1 = 12
                    PMIOS_TTLRRAMT = Round(N2Str2Zero(rsRR_HD!PurchPrice), 2)
                    PMIOS_DS_AMT1 = Round((N2Str2Zero(rsRR_HD!PurchPrice) / 1.12) * 0.12, 2)
                    PMIOS_NETRRAMT = Round(N2Str2Zero(rsRR_HD!PurchPrice) - PMIOS_DS_AMT1, 2)
                    PMIOS_STATUS = Null2String(rsRR_HD!IStatus)
                End If
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0

                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If

                J_JDATE = N2Date2Null(PMIOS_RRDATE)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_JTYPE = "'APJ'"
                If COMPANY_CODE = "HAS" Then
                    J_REMARKS = "'To record New Vehicle Purchases. with Ref# " & PMIOS_RRNO + ":" + CONDUCTION & "'"
                Else
                    J_REMARKS = "'To record New Vehicle Purchases. with Ref# " & PMIOS_RRNO + "'"
                End If
                J_VENDORCODE = N2Str2Null(PMIOS_RECVD_CODE)
                J_CUSTOMERCODE = "'999999'"

                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                J_AMOUNTTOPAY = Round(NumericVal(PMIOS_TTLRRAMT), 2)
                J_INVOICEAMT = 0
                J_BALANCE = Round(NumericVal(PMIOS_TTLRRAMT), 2)
                J_AMOUNTPAID = 0

                J_STATUS = "'N'"

                J_INVOICEDATE = N2Date2Null(PMIOS_RRDATE)
                J_INVOICENO = N2Str2Null(rsRR_HD!code)
                J_CHECKNO = "NULL"
                J_DUEDATE = N2Date2Null(Format(DateAdd("d", NumericVal(PMIOS_TERMS), Format(PMIOS_RRDATE, "DD-MMM-YY"))))
                J_PAYTYPE = "'" & PMIOS_TERMS & "'"
                J_INVOICETYPE = "'VEHICLES'"
                J_CHECKDATE = "NULL"
                J_BANKCODE = "NULL"
                J_REFNO = "NULL"
                J_REFDATE = "NULL"
                J_TERMS = "'" & PMIOS_TERMS & "'"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"

                'CASH ON HAND
                If PMIOS_NETRRAMT > 0 Then
                    J_JITEMNO = "'0001'"
                    ' Update By : BTT - 06252008
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "SALES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "SALES")))
                        J_DEBIT = NumericVal(PMIOS_NETRRAMT)
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    Else

                        If COMPANY_CODE = "HGC" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "IN TRANSIT"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "IN TRANSIT")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("HYUNDAI"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("HYUNDAI")))
                        End If
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(PMIOS_TTLRRAMT), 2)
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    End If
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                    
                    J_JITEMNO = "'0002'"
                    'Update By : BTT - 06252008
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("INVENTORY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("INVENTORY")))
                        J_DEBIT = 0
                        J_CREDIT = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT))
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT


                    ElseIf COMPANY_CODE = "HSB" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "INVA"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "INVA")))
                        J_DEBIT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    Else

                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "Vehicles"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "Vehicles")))
                        J_DEBIT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    End If
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                    
                    J_JITEMNO = "'0003'"
                    J_ACCT_CODE = N2Str2Null(ReturnInPutTax())
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInPutTax()))
                    J_DEBIT = Round(NumericVal(PMIOS_DS_AMT1), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                    
                    
                    '                J_JITEMNO = "'0003'"
                    '                J_ACCT_CODE = N2Str2Null(ReturnInComeTax("EXPANDED"))
                    '                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInComeTax("EXPANDED")))
                    '                J_DEBIT = 0
                    '                J_CREDIT = NumericVal(PMIOS_NETRRAMT) * 0.01
                    '                J_TAX = 0
                    '                J_GROSS = 0
                    '                J_NET = 0
                    '                J_STATUS = "'N'"
                    '                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '
                    '                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     '                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     '                                 " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     '                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     '                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & "," & J_STATUS & ")"
                    '                J_JITEMNO = "'0004'"


                End If

                J_JVOUCHERNO = J_VOUCHERNO
                PV_ITEMNO = N2Str2Null("0001")
                PV_MRRNO = N2Str2Null(PMIOS_RRNO)
                PV_PONO = N2Str2Null(PMIOS_PONO)
                PV_INVNO = N2Str2Null(PMIOS_INVNO)
                PV_PRODNO = "NULL"
                PV_AMOUNT = Round(NumericVal(PMIOS_TTLRRAMT), 2)
                PV_STATUS = "'N'"

                SQL_STATEMENT = "insert into AMIS_PV_Detail " & _
                                 "(VoucherNo,JTYPE,JDATE,itemno,PO_No,MRR_No,INV_No,PROD_No,Amount,status)" & _
                               " values (" & J_JVOUCHERNO & ",'VI'," & J_JDATE & ", " & PV_ITEMNO & ", " & PV_PONO & _
                                 ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                 ", " & PV_STATUS & ")"
                gconDMIS.Execute SQL_STATEMENT

                NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(PV_MRRNO)
                
                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                               " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                gconDMIS.Execute SQL_STATEMENT
                
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                
                '=========================================================================================================================
                'Entry For Clering Accoung PSB : Update By BTT - 06232008
                If COMPANY_CODE = "HBK" Then

                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")

                    Dim J_ClearingNo                                  As String
                    Dim VendorCode                                    As String
                    Dim VendorName                                    As String

                    VendorCode = "P00007"
                    VendorName = "PS Bank"
                    J_AMOUNTTOPAY = NumericVal(0)
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                        J_ClearingNo = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    Else
                        J_ClearingNo = "'000001'"
                    End If
                    J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                    'Detail
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("INVENTORY"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("INVENTORY")))
                    J_DEBIT = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT))
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_ClearingNo & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    J_JITEMNO = "'0002'"
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("AP", "FLOORSTOCK"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("AP", "FLOORSTOCK")))
                    J_DEBIT = 0
                    J_CREDIT = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT))
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_ClearingNo & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    J_JVOUCHERNO = J_VOUCHERNO
                    PV_ITEMNO = N2Str2Null("0001")
                    PV_MRRNO = N2Str2Null(PMIOS_RRNO)
                    PV_PONO = N2Str2Null(PMIOS_PONO)
                    PV_INVNO = N2Str2Null(PMIOS_INVNO)
                    PV_PRODNO = "NULL"
                    PV_AMOUNT = Round(NumericVal(PMIOS_TTLRRAMT), 2)
                    PV_STATUS = "'N'"

                    gconDMIS.Execute "insert into AMIS_PV_Detail " & _
                                     "(VoucherNo,JDATE,itemno,PO_No,MRR_No,INV_No,PROD_No,Amount,status)" & _
                                   " values (" & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & ", " & PV_PONO & _
                                     ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                     ", " & PV_STATUS & ")"
                    'Header
                    gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                                   " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", '" & VendorCode & "','" & VendorName & "', " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                     ", " & J_ClearingNo & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"

                End If
                '
                'End of Clearing Acoount
                '*****************************************************************************************************

            End If
            Grid2.Cell(GridImport, 1).Text = 1
        End If
        i = i + 1
        progCPB.Value = (i / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
    Next
    'Sublet
    ImportSublet
    Screen.MousePointer = 0
    MsgBox "Import Successfully Completed!", vbInformation, "Finish"
    LogAudit "R", "ACCOUNTS PAYABLE IMPORT", dtpTranDate
    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdClearJournals_Click()
    Dim rsCHATCheckControlIfExistRecordInJournalHD                    As ADODB.Recordset
    Set rsCHATCheckControlIfExistRecordInJournalHD = New ADODB.Recordset
    Set rsCHATCheckControlIfExistRecordInJournalHD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'APJ' and Jdate = '" & CDate(dtpTranDate) & "'")
    If Not rsCHATCheckControlIfExistRecordInJournalHD.EOF And Not rsCHATCheckControlIfExistRecordInJournalHD.BOF Then
        Screen.MousePointer = 0
        If MsgBox("Clear Unposted Data for this Particular Date?", vbQuestion + vbYesNo, "Confirm...") = vbNo Then
            Exit Sub
        Else
            gconDMIS.Execute ("Delete from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'APJ' and JDate = '" & CDate(dtpTranDate) & "'")
            gconDMIS.Execute ("Delete from AMIS_Journal_Det Where STATUS <> 'P' AND Jtype = 'APJ' and JDate = '" & CDate(dtpTranDate) & "'")
            gconDMIS.Execute ("Delete from AMIS_PV_Detail Where STATUS <> 'P' AND JDate = '" & CDate(dtpTranDate) & "'")
            cmdShowTrans.Value = True
            Screen.MousePointer = 0
            MsgBox "Existing Data Successfully deleted.", vbInformation, "Deleted"
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdShowTrans_Click()
    Screen.MousePointer = 11
InitGrids:     DoEvents: cmdCheck.Enabled = False: cmdClearJournals.Enabled = False
    Grid1.Rows = 2: Grid2.Rows = 2: KIM = 0
    Dim RRTYPE                                                        As String
    Dim IS_Exist                                                      As Byte
    Dim rsRR_HD                                                       As ADODB.Recordset
    Dim rsPURCH_AGREE                                                 As ADODB.Recordset
    Set rsRR_HD = New ADODB.Recordset
    Set rsRR_HD = gconDMIS.Execute("Select * from PMIS_vw_RR_TRANS Where STATUS = 'P' AND (CLASSCODE = 'PCG' or CLASSCODE = 'PCS' ) AND RRDATE = '" & CDate(dtpTranDate) & "' Order by RRNO ASC")
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        rsRR_HD.MoveFirst: KIM = 0
        Grid1.AutoRedraw = False
        Do While Not rsRR_HD.EOF
            KIM = KIM + 1
            If Null2String(rsRR_HD!Type) = "P" Then
                RRTYPE = "PARTS"
            ElseIf Null2String(rsRR_HD!Type) = "A" Then
                RRTYPE = "ACCESSORIES"
            ElseIf Null2String(rsRR_HD!Type) = "M" Then
                RRTYPE = "MATERIALS"
            Else
                RRTYPE = ""
            End If
            If CheckAPJExisting(Null2String(rsRR_HD!RRNO), RRTYPE) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            Grid1.AddItem IS_Exist & Chr(9) & RRTYPE & Chr(9) & Null2String(rsRR_HD!RRNO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsRR_HD!TTLRRAMT)) & Chr(9) & Null2String(rsRR_HD!RECVD_FROM)
            rsRR_HD.MoveNext
        Loop
        If KIM > 0 Then Grid1.RemoveItem 1
        Grid1.AutoRedraw = True
        Grid1.Refresh
    End If
    Set rsPURCH_AGREE = New ADODB.Recordset
    Set rsPURCH_AGREE = gconDMIS.Execute("Select SMIS_MRRINV.*,CSMS_SELLINGDEALER.DEALERNAME from SMIS_MRRINV left outer JOIN CSMS_SELLINGDEALER ON SMIS_MRRINV.SOURCE = CSMS_SELLINGDEALER.DEALERCODE Where STATUS = 'P' AND DateReceived = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' Order by DateReceived ASC")
    If Not rsPURCH_AGREE.EOF And Not rsPURCH_AGREE.BOF Then
        rsPURCH_AGREE.MoveFirst: KIM = 0
        Grid2.AutoRedraw = False
        Do While Not rsPURCH_AGREE.EOF
            KIM = KIM + 1
            If CheckAPJExisting(Null2String(rsPURCH_AGREE!code), "VEHICLES") = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            Grid2.AddItem IS_Exist & Chr(9) & "VEHICLES" & Chr(9) & Null2String(rsPURCH_AGREE!code) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPURCH_AGREE!PurchPrice)) & Chr(9) & Null2String(rsPURCH_AGREE!DealerName)
            rsPURCH_AGREE.MoveNext
        Loop
        If KIM > 0 Then Grid2.RemoveItem 1
        Grid2.AutoRedraw = True
        Grid2.Refresh
    End If
    If KIM > 0 Then
        cmdCheck.Enabled = True
        cmdClearJournals.Enabled = True
    End If
    Screen.MousePointer = 0
    'Update By : BTT 07142008 : to Process the Sublet in CSMS
    Set RsSublet = New ADODB.Recordset
    Set RsSublet = gconDMIS.Execute("Select * from CSMS_PO_RC_HD Where STATUS = 'P' AND Rc_DATE = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' Order by RC_date ASC")
    Grid3.Rows = 1
    If Not RsSublet.EOF And Not RsSublet.BOF Then
        RsSublet.MoveFirst: KIM = 0
        Grid3.AutoRedraw = False
        Do While Not RsSublet.EOF
            If CheckSubletifExist(Null2String(RsSublet!RC_no), Null2String(RsSublet!po_no)) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            Grid3.AddItem IS_Exist & Chr(9) & "SUBLET" & Chr(9) & Null2String(RsSublet!RC_no) & Chr(9) & ToDoubleNumber(N2Str2Zero(RsSublet!sublet_Total_net_AMT)) & Chr(9) & Null2String(RsSublet!contractor_name)
            RsSublet.MoveNext
        Loop
    End If
    If KIM > 0 Then Grid3.RemoveItem 1
    Grid3.AutoRedraw = True
    Grid3.Refresh
    cmdCheck.Enabled = True
    cmdClearJournals.Enabled = True
    'End of Update
    Screen.MousePointer = 0
End Sub

Private Sub dtpTranDate_Change()
InitGrids:     DoEvents:
    Grid1.Rows = 1
    Grid2.Rows = 1
    cmdCheck.Enabled = False
    cmdClearJournals.Enabled = False
End Sub

Private Sub Form_Load()
    On Error GoTo Errorcode
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpTranDate = LOGDATE
    InitGrids
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    Screen.MousePointer = 0
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Database Connection Error!"
    Unload frmSplash
    cmdCheck.Enabled = False
End Sub


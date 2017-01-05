VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Begin VB.Form frmCRJImport_HAI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Receipts Import Process"
   ClientHeight    =   7815
   ClientLeft      =   345
   ClientTop       =   1110
   ClientWidth     =   14010
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmCRJImport_HAI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   14010
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1485
      Left            =   9360
      ScaleHeight     =   1425
      ScaleWidth      =   4515
      TabIndex        =   18
      Top             =   6180
      Width           =   4575
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Paid Invoices If not yet imported, will be automatically imported in Sales Journal since its already paid."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1275
         Left            =   60
         TabIndex        =   19
         Top             =   90
         Width           =   4395
      End
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
      Left            =   11940
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
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
      Left            =   3660
      MouseIcon       =   "frmCRJImport_HAI.frx":030A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Process Import of SALES"
      Top             =   120
      Width           =   2010
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Deposited Official Receipts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   1
      Top             =   6540
      Width           =   4155
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Un-Deposited Official Receipts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   6210
      Value           =   -1  'True
      Width           =   4155
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
      Left            =   8580
      MouseIcon       =   "frmCRJImport_HAI.frx":045C
      MousePointer    =   99  'Custom
      Picture         =   "frmCRJImport_HAI.frx":05AE
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit Window"
      Top             =   6870
      Width           =   720
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
      Left            =   7875
      MouseIcon       =   "frmCRJImport_HAI.frx":0914
      MousePointer    =   99  'Custom
      Picture         =   "frmCRJImport_HAI.frx":0A66
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Process Importing of Cash Receipts "
      Top             =   6870
      Width           =   720
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   4740
      TabIndex        =   4
      Top             =   6450
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   556
      Picture         =   "frmCRJImport_HAI.frx":0D01
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "frmCRJImport_HAI.frx":0D1D
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
   Begin MSComCtl2.DTPicker dtpTranDate 
      Height          =   405
      Left            =   1860
      TabIndex        =   2
      Top             =   120
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
      Format          =   48037889
      CurrentDate     =   38216
   End
   Begin FlexCell.Grid Grid1 
      Height          =   4905
      Left            =   90
      TabIndex        =   8
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin FlexCell.Grid Grid2 
      Height          =   4905
      Left            =   4740
      TabIndex        =   9
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   1320
      TabIndex        =   15
      Top             =   1590
      Visible         =   0   'False
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin FlexCell.Grid Grid3 
      Height          =   4905
      Left            =   9360
      TabIndex        =   16
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PMIS/CSMS/SMIS"
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
      TabIndex        =   17
      Top             =   660
      Width           =   4545
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
      Left            =   180
      TabIndex        =   14
      Top             =   7200
      Width           =   7995
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNDEPOSITED OR'S"
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
      Left            =   90
      TabIndex        =   11
      Top             =   660
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DEPOSITED OR'S"
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
      Left            =   4740
      TabIndex        =   10
      Top             =   660
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
      Left            =   120
      TabIndex        =   5
      Top             =   210
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
      Left            =   4770
      TabIndex        =   3
      Top             =   6180
      Width           =   5835
   End
End
Attribute VB_Name = "frmCRJImport_HAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gconBIRData                         As ADODB.Connection
Dim LIM As Integer

Private Sub cmdCheck_Click()

    'CASH RECEIPTS
    COA_CASH_ON_HAND = "'11-01005-00'"
    COA_CUSTOMER_DEPOSIT = "'21-02402-00'"

    COA_INSURANCE_PREMIUM_PAYABLE = "'21-02400-00'"
    COA_INSURANCE_PREMIUM_RENEWAL = "'21-02401-00'"
    COA_LTO_PAYMENT = "'21-02403-00'"


    COA_CHATTEL_MORTGAGE_FEE_PAYABLE = "'21-02204-00'"
    COA_NEW_VEHICLE_REGISTRATION = "'61-01203-10'"
    COA_WARRANTY_CLAIMS_RECEIVABLE = "'11-02000-00'"

    COA_ACCOUNTS_RECEIVABLE_NONTRADE_EMPLOYEES = "'11-03300-00'"    ' advances to employees
    COA_INCIDENTAL_CHARGES_UNITS = "''"
    COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD = "'11-02102-00'"
    COA_PRE_DELIVERY = "'71-43100-10'"

    COA_CORPORATE_TAX_WHELD = "'11-07200-00'"
    COA_CORPORATE_VAT_WHELD = "''"

    COA_COST_OF_SALES_PARTS = "'61-03101-30'"
    COA_COST_OF_SALES_GOL = "'61-03100-30"
    COA_COST_OF_SALES_VEHICLES = "XXXX"

    COA_INVENTORIES_PARTS = "'11-05300-00'"
    COA_INVENTORIES_GOL = "'11-05100-00'"
    COA_INVENTORIES_VEHICLES = "'11-05508-00'"

    COA_INPUT_TAX = "'11-07000-00'"
    COA_INCOME_TAX_WITHHELD = "'21-04000-00'"
    COA_ACCOUNTS_PAYABLE = "'21-01100-00'"

    COA_AR_TRADE_UNITS = "'11-02100-00'"
    COA_AR_TRADE_SERVICE = "'11-02200-00'"
    COA_AR_TRADE_PARTS = "'11-02300-00'"


    If Function_Access(LOGID, "Acess_Process", "IMPORT CASH RECEIPTS") = False Then Exit Sub
    Screen.MousePointer = 11
    Dim rsCHATCheckControlIfExistRecordInJournalHD As ADODB.Recordset
    If Option1.Value = True Then
        Call ImportPMISSales
        Call ImportCSMSSales
        Call ImportSMISSales
        Call ImportUnDeposit
    End If
    If Option2.Value = True Then
        Call ImportDeposited
    End If
    MsgBox "Import Successfully Completed!", vbInformation, "Finish"
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Function SetOTHChartCodes(XXX As String) As String
    Dim rsSBOOK_CHARTCODES              As ADODB.Recordset
    Set rsSBOOK_CHARTCODES = New ADODB.Recordset
    Set rsSBOOK_CHARTCODES = gconDMIS.Execute("Select * from CMIS_SBOOK where BOOK = 'D' and CODE = '" & XXX & "'")
    If Not rsSBOOK_CHARTCODES.EOF And Not rsSBOOK_CHARTCODES.BOF Then
        SetOTHChartCodes = Null2String(rsSBOOK_CHARTCODES!CHARTCODES)
    End If
    Set rsSBOOK_CHARTCODES = Nothing
End Function

Sub ImportUnDeposit()
    'HEADER
    Dim J_JDATE, J_VOUCHERNO, J_JTYPE   As String
    Dim J_JNO, J_REMARKS, J_VENDORCODE, J_CUSTOMERCODE As String
    Dim J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
    Dim J_CHECKNO                       As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE As String
    Dim J_INVOICETYPE, J_INVOICENO      As String
    Dim J_CHECKDATE, J_BANKCODE         As String
    Dim J_REFNO, J_REFDATE              As String
    Dim J_TERMS, J_DEALER               As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS   As String

    'DETAIL
    Dim J_ACCT_CODE, J_ACCT_NAME        As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET As Double
    Dim J_STATUS, J_JITEMNO             As String

    Dim rsJournal_HDDup                 As ADODB.Recordset

    Dim CMIS_OR_NUM                     As String
    Dim CMIS_OR_DATE                    As String
    Dim CMIS_OR_AMT                     As String
    Dim CMIS_DISCOUNT                   As String
    Dim CMIS_TAX                        As String
    Dim CMIS_CASHAMOUNT                 As Double
    Dim CMIS_CHKAMOUNT                  As Double
    Dim CMIS_CARDAMOUNT                 As Double
    Dim CMIS_CUSCDE                     As String
    Dim CMIS_CUSNAME                    As String
    Dim CMIS_DEPOSIT                    As String
    Dim CMIS_BANKCODE                   As String
    Dim CMIS_TSEKE                      As String
    Dim CMIS_CHECKDATE                  As String
    Dim CMIS_STATUS                     As String
    Dim CMIS_TYPE_PAYMENT               As String

    Dim CMIS_DT_TRANTYPE                As String
    Dim CMIS_DT_REFERENCE               As String
    Dim CMIS_DT_CUSCDE                  As String
    Dim CMIS_DT_DESCRIPT                As String
    Dim CMIS_DT_AMOUNT                  As Double
    Dim CMIS_DT_DOCDTE                  As String
    Dim CMIS_DT_PAYMENT                 As Double
    Dim CMIS_DT_DISCOUNT                As Double
    Dim CMIS_DT_TAX                     As Double
    Dim CMIS_DT_PAIDFOR                 As String
    Dim CMIS_IS_VAT                     As Boolean

    Dim TOTAL_DEBIT, TOTAL_CREDIT       As Double

    Dim rsOFF_HD                        As ADODB.Recordset
    Dim rsOFF_DT                        As ADODB.Recordset
    Dim I                               As Long

    Dim rsSJ_DATA                       As ADODB.Recordset

    Dim PV_PONO, PV_MRRNO, PV_INVNO, PV_PRODNO As String
    Dim J_JVOUCHERNO                    As String
    Dim PV_AMOUNT                       As Double
    Dim PV_STATUS, PV_ITEMNO            As String

    Dim SJ_PV_ITEMNO                    As Integer
    Dim rsCheckJournal_HD               As ADODB.Recordset

    Dim rsREPOR, rsORD_HD               As ADODB.Recordset

    Dim GridImport                      As Integer
    I = 0
    For GridImport = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(GridImport, 1).Text) = 0 Then
            Set rsOFF_HD = New ADODB.Recordset
            If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD Where OR_NUM = '" & Grid1.Cell(GridImport, 3).Text & "' AND VAT = 1 AND OR_DATE = '" & CDate(dtpTranDate) & "' Order by OR_NUM ASC")
            Else
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD Where OR_NUM = '" & Grid1.Cell(GridImport, 3).Text & "' AND VAT = 0 AND OR_DATE = '" & CDate(dtpTranDate) & "' Order by OR_NUM ASC")
            End If
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                CMIS_OR_NUM = Null2String(rsOFF_HD!OR_NUM)
                CMIS_OR_DATE = Null2Date(rsOFF_HD!OR_DATE)
                CMIS_OR_AMT = Null2String(rsOFF_HD!OR_AMT)
                CMIS_DISCOUNT = Null2String(rsOFF_HD!DISCOUNT)
                CMIS_TAX = Null2String(rsOFF_HD!TAX)
                CMIS_CASHAMOUNT = N2Str2Zero(rsOFF_HD!CASHAMOUNT)
                CMIS_CHKAMOUNT = N2Str2Zero(rsOFF_HD!CHKAMOUNT)
                CMIS_CARDAMOUNT = N2Str2Zero(rsOFF_HD!CARDAMOUNT)
                CMIS_CUSCDE = Null2String(rsOFF_HD!CUSCDE)
                CMIS_CUSNAME = Null2String(rsOFF_HD!CUSNAME)
                CMIS_DEPOSIT = Null2String(rsOFF_HD!DEPOSIT)
                CMIS_BANKCODE = Null2String(rsOFF_HD!BankCode)
                CMIS_TSEKE = Null2String(rsOFF_HD!TSEKE) & Null2String(rsOFF_HD!CARDNUMBER)
                CMIS_TYPE_PAYMENT = Null2String(rsOFF_HD!TOF)
                CMIS_CHECKDATE = Format(Null2Date(rsOFF_HD!CheckDate), "MM/DD/YYYY")
                CMIS_STATUS = Null2String(rsOFF_HD!Status)
                CMIS_IS_VAT = Null2Bool(rsOFF_HD!VAT)
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0

                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If

                J_JDATE = N2Date2Null(CMIS_OR_DATE)
                J_VOUCHERNO = N2Str2Null(GetCRJVoucherNo())
                J_JTYPE = "'CRJ'"

                Set rsOFF_DT = New ADODB.Recordset
                Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE OR_NUM = '" & CMIS_OR_NUM & "'")
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst
                    Do While Not rsOFF_DT.EOF
                        If Null2String(rsOFF_DT!TRANTYPE) = "OTH" Then
                            J_REMARKS = SetOtherTransaction(Null2String(rsOFF_DT!PAIDFOR)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!PAYMENT))
                        Else
                            J_REMARKS = SetTransaction(Null2String(rsOFF_DT!TRANTYPE)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!PAYMENT))
                        End If
                        rsOFF_DT.MoveNext
                        If Not rsOFF_DT.EOF Then J_REMARKS = "" & Chr(9)
                    Loop
                    J_REMARKS = N2Str2Null(J_REMARKS)
                Else
                    J_REMARKS = "NULL"
                End If
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(CMIS_CUSCDE)

                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = NumericVal(CMIS_OR_AMT)
                J_BALANCE = 0
                J_AMOUNTPAID = 0

                J_STATUS = "'N'"

                J_INVOICEDATE = N2Date2Null(CMIS_OR_DATE)
                If CMIS_IS_VAT = True Then
                    J_INVOICENO = N2Str2Null(Left(CMIS_OR_NUM, 6))
                Else
                    J_INVOICENO = N2Str2Null("NV" & Left(CMIS_OR_NUM, 6))
                End If
                J_CHECKNO = N2Str2Null(CMIS_TSEKE)
                J_DUEDATE = N2Date2Null(CMIS_CHECKDATE)
                If Null2String(rsOFF_HD!TOF) = "1" Then
                    J_PAYTYPE = "'CASH'"
                ElseIf Null2String(rsOFF_HD!TOF) = "2" Then
                    J_PAYTYPE = "'CHECK'"
                ElseIf Null2String(rsOFF_HD!TOF) = "3" Then
                    J_PAYTYPE = "'CARD'"
                Else
                    J_PAYTYPE = "NULL"
                End If
                J_INVOICETYPE = "'CI'"
                J_CHECKDATE = N2Str2Null(CMIS_CHECKDATE)
                J_BANKCODE = N2Str2Null(CMIS_BANKCODE)
                J_REFNO = N2Str2Null(CMIS_TSEKE)
                J_REFDATE = N2Date2Null(CMIS_CHECKDATE)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"

                'CASH ON HAND
                If CMIS_TYPE_PAYMENT = "1" Or CMIS_TYPE_PAYMENT = "2" Then
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = N2Str2Null(COA_CASH_ON_HAND)
                    J_ACCT_NAME = N2Str2Null(Setacctname(COA_CASH_ON_HAND))
                    If CMIS_CASHAMOUNT > 0 Then
                        J_DEBIT = NumericVal(CMIS_CASHAMOUNT)
                    Else
                        J_DEBIT = NumericVal(CMIS_CHKAMOUNT)
                    End If
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CMIS_TYPE_PAYMENT = "3" Then
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = N2Str2Null(COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD)
                    J_ACCT_NAME = N2Str2Null(Setacctname(COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD))
                    J_DEBIT = NumericVal(CMIS_CARDAMOUNT)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                Set rsOFF_DT = New ADODB.Recordset
                Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT Where OR_NUM = " & N2Str2Null(CMIS_OR_NUM))
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst: SJ_PV_ITEMNO = 0
                    Do While Not rsOFF_DT.EOF
                        CMIS_DT_TRANTYPE = Null2String(rsOFF_DT!TRANTYPE)
                        CMIS_DT_REFERENCE = Null2String(rsOFF_DT!InvoiceNo)
                        CMIS_DT_CUSCDE = Null2String(rsOFF_DT!CUSCDE)
                        CMIS_DT_DESCRIPT = Null2String(rsOFF_DT!DESCRIPT)
                        CMIS_DT_AMOUNT = N2Str2Zero(rsOFF_DT!Amount)
                        CMIS_DT_DOCDTE = Null2String(rsOFF_DT!DOCDTE)
                        CMIS_DT_PAYMENT = N2Str2Zero(rsOFF_DT!PAYMENT)
                        CMIS_DT_DISCOUNT = N2Str2Zero(rsOFF_DT!DISCOUNT)
                        CMIS_DT_TAX = N2Str2Zero(rsOFF_DT!TAX)
                        CMIS_DT_PAIDFOR = Null2String(rsOFF_DT!PAIDFOR)
                        SJ_PV_ITEMNO = SJ_PV_ITEMNO + 1
                        PV_MRRNO = "'" & CMIS_DT_TRANTYPE & "'"
                        Set rsSJ_DATA = New ADODB.Recordset
                        Set rsSJ_DATA = gconDMIS.Execute("Select * from AMIS_Journal_HD Where jtype = 'SJ' and invoicetype = " & PV_MRRNO & " and invoiceno = " & N2Str2Null(CMIS_DT_REFERENCE))
                        If Not rsSJ_DATA.EOF And Not rsSJ_DATA.BOF Then
                            rsSJ_DATA.MoveFirst
                            Do While Not rsSJ_DATA.EOF
                                J_JVOUCHERNO = J_VOUCHERNO
                                PV_ITEMNO = N2Str2Null(Format(SJ_PV_ITEMNO, "0000"))
                                PV_INVNO = N2Str2Null(CMIS_DT_REFERENCE)
                                PV_PRODNO = N2Date2Null(rsSJ_DATA!InvoiceDate)
                                'PV_AMOUNT = N2Str2Zero(rsSJ_DATA!InvoiceAmt)
                                'MODIFIED BY FML: 07262008
                                PV_AMOUNT = CMIS_DT_PAYMENT
                                PV_STATUS = "'N'"
                                gconDMIS.Execute "Delete from AMIS_CRJ_Detail Where VoucherNo = " & J_JVOUCHERNO & " AND JDate = " & J_JDATE & _
                                                 " AND ItemNo = " & PV_ITEMNO & " AND INVOICETYPE = " & PV_MRRNO & _
                                                 " AND INVOICENO = " & PV_INVNO & " AND INVOICEDATE = " & PV_PRODNO & " AND INVOICEAMOUNT = " & PV_AMOUNT
                                gconDMIS.Execute "insert into AMIS_CRJ_Detail " & _
                                                 "(VoucherNo,Jdate,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status)" & _
                                               " values (" & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & _
                                                 ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                                 ", " & PV_STATUS & ")"
                                Set rsCheckJournal_HD = New ADODB.Recordset
                                Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where ID  = " & rsSJ_DATA!ID)
                                If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
                                    If N2Str2Zero(rsCheckJournal_HD!InvoiceAmt) <= PV_AMOUNT Then
                                        gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                                       " ReceiveStatus = 'Y' " & "," & _
                                                       " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                       " Balance = Balance - " & PV_AMOUNT & _
                                                       " where ID = " & rsSJ_DATA!ID
                                    Else
                                        gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                                       " ReceiveStatus = 'N' " & "," & _
                                                       " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                       " Balance = Balance - " & PV_AMOUNT & _
                                                       " where ID = " & rsSJ_DATA!ID
                                    End If
                                Else
                                    Set rsCheckJournal_HD = New ADODB.Recordset
                                    Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'")
                                    If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
                                        If N2Str2Zero(rsCheckJournal_HD!InvoiceAmt) <= PV_AMOUNT Then
                                            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                                           " ReceiveStatus = 'Y' " & "," & _
                                                           " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                           " Balance = Balance - " & PV_AMOUNT & _
                                                           " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"
                                        Else
                                            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                                           " ReceiveStatus = 'N' " & "," & _
                                                           " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                           " Balance = Balance - " & PV_AMOUNT & _
                                                           " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"
                                        End If
                                    End If
                                End If
                                rsSJ_DATA.MoveNext
                            Loop
                        End If

                        J_JITEMNO = "'0002'"
                        If CMIS_DT_TRANTYPE = "RO" Or CMIS_DT_TRANTYPE = "SI" Then
                            J_ACCT_CODE = COA_AR_TRADE_SERVICE
                            J_ACCT_NAME = N2Str2Null(Setacctname(COA_AR_TRADE_SERVICE))
                        End If
                        If CMIS_DT_TRANTYPE = "PI" Then
                            J_ACCT_CODE = COA_AR_TRADE_PARTS
                            J_ACCT_NAME = N2Str2Null(Setacctname(COA_AR_TRADE_PARTS))
                        End If
                        If CMIS_DT_TRANTYPE = "AI" Then
                            J_ACCT_CODE = "'11-02003-00'"
                            J_ACCT_NAME = N2Str2Null(Setacctname("'11-02003-00'"))
                        End If
                        If CMIS_DT_TRANTYPE = "VI" Then
                            J_ACCT_CODE = COA_AR_TRADE_UNITS
                            J_ACCT_NAME = N2Str2Null(Setacctname(COA_AR_TRADE_UNITS))
                        End If
                        If CMIS_DT_TRANTYPE = "EST" Then
                            J_ACCT_CODE = COA_CUSTOMER_DEPOSIT
                            J_ACCT_NAME = N2Str2Null(Setacctname(COA_CUSTOMER_DEPOSIT))
                        End If

                        If CMIS_DT_TRANTYPE = "OTH" Then
                            CMIS_DT_AMOUNT = CMIS_DT_PAYMENT
                            J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(CMIS_DT_PAIDFOR))
                            J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(CMIS_DT_PAIDFOR)))
                            If CMIS_DT_PAIDFOR = "429" Then
                                J_JITEMNO = "'0002'"
                                J_ACCT_CODE = N2Str2Null(COA_SALES_DISCOUNT_SERVICE_CHARGE_LABOR)
                                J_ACCT_NAME = N2Str2Null(Setacctname(COA_SALES_DISCOUNT_SERVICE_CHARGE_LABOR))
                                J_DEBIT = NumericVal(CMIS_DT_DISCOUNT)
                                J_CREDIT = 0
                                J_TAX = 0
                                J_GROSS = 0
                                J_NET = 0
                                J_STATUS = "'N'"
                                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                                J_JITEMNO = "'0003'"
                                J_ACCT_CODE = N2Str2Null(COA_CORPORATE_TAX_WHELD)
                                J_ACCT_NAME = N2Str2Null(Setacctname(COA_CORPORATE_TAX_WHELD))
                                J_DEBIT = NumericVal(CMIS_DT_TAX)
                                J_CREDIT = 0
                                J_TAX = 0
                                J_GROSS = 0
                                J_NET = 0
                                J_STATUS = "'N'"
                                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                                J_JITEMNO = "'0004'"
                                J_ACCT_CODE = COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD
                                J_ACCT_NAME = N2Str2Null(Setacctname(COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD))
                            End If
                        End If
                        'J_GROSS = NumericVal(CMIS_DT_AMOUNT)
                        'MODIFIED FML: 07/26/2008
                        J_GROSS = NumericVal(CMIS_DT_PAYMENT)
                        J_TAX = 0
                        J_NET = NumericVal(J_GROSS - J_TAX)
                        J_DEBIT = 0
                        J_CREDIT = NumericVal(J_NET)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        rsOFF_DT.MoveNext
                    Loop
                End If
                gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                               " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                Grid1.Cell(GridImport, 1).Text = 1
            End If
        End If
        I = I + 1
        progCPB.Value = (I / (Grid1.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
    Screen.MousePointer = 0
End Sub

Sub InitGrid3()
With Grid3
    .Rows = 1
    .Cell(0, 1).Text = "Imported"
    .Cell(0, 2).Text = "Inv. Type"
    .Cell(0, 3).Text = "Inv. No."
    .Cell(0, 4).Text = "Inv. Amt."
    .Cell(0, 5).Text = "Customer"
    
    .Column(0).Width = 10
    .Column(1).Width = 50
    .Column(2).Width = 80
    .Column(3).Width = 80
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

Sub ShowUnImportedPaidInvoices(VarTranType As String, VarTranno As String)
Screen.MousePointer = 11
Dim InvoiceType, InvoiceTypeCode As String
Dim IS_Exist As Byte
LIM = 0
If VarTranType = "PI" Or VarTranType = "MI" Or VarTranType = "AI" Then
    Set rsPMIOS_ORD_HD = New ADODB.Recordset
    Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where TYPE = '" & Left(VarTranType, 1) & "' AND (TranType = 'CSH' OR  TranType = 'CHG') and STATUS = 'P' AND tranno = '" & VarTranno & "' order by Tranno ASC")
    If Not rsPMIOS_ORD_HD.EOF And Not rsPMIOS_ORD_HD.BOF Then
       rsPMIOS_ORD_HD.MoveFirst:
       Do While Not rsPMIOS_ORD_HD.EOF
          LIM = LIM + 1
          If Null2String(rsPMIOS_ORD_HD!Type) = "P" Then
             InvoiceType = "Parts"
             InvoiceTypeCode = "PI"
          ElseIf Null2String(rsPMIOS_ORD_HD!Type) = "A" Then
             InvoiceType = "Accessories"
             InvoiceTypeCode = "AI"
          ElseIf Null2String(rsPMIOS_ORD_HD!Type) = "M" Then
             InvoiceType = "Materials"
             InvoiceTypeCode = "MI"
          Else
             InvoiceType = "Unknown"
             InvoiceTypeCode = ""
          End If
          If CheckSJExisting(InvoiceTypeCode, Null2String(rsPMIOS_ORD_HD!tranno)) = True Then
             IS_Exist = 1
          Else
             IS_Exist = 0
          End If
          Grid3.AddItem IS_Exist & Chr(9) & UCase(InvoiceType) & Chr(9) & Null2String(rsPMIOS_ORD_HD!TRANTYPE) & "-" & Null2String(rsPMIOS_ORD_HD!tranno) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPMIOS_ORD_HD!NetInvAmt)) & Chr(9) & Null2String(rsPMIOS_ORD_HD!CUSTNAME)
          rsPMIOS_ORD_HD.MoveNext
       Loop
    End If
End If
If VarTranType = "SI" Then
    Set rsCSMIOS_REPOR = New ADODB.Recordset
    Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where invoice <> 'NO CHG' AND invoice <> 'PDI RO' and invoice = '" & VarTranno & "' order by invoice ASC")
    If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
       rsCSMIOS_REPOR.MoveFirst:
       Do While Not rsCSMIOS_REPOR.EOF
          LIM = LIM + 1
          If CheckSJExisting("SI", Null2String(rsCSMIOS_REPOR!Invoice)) = True Then
             IS_Exist = 1
          Else
             IS_Exist = 0
          End If
          'Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!Invoice) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!Amount)) & Chr(9) & Null2String(rsCSMIOS_REPOR!Niym)
          Grid3.AddItem IS_Exist & Chr(9) & "SERVICE" & Chr(9) & Null2String(rsCSMIOS_REPOR!Invoice) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!Amount)) & Chr(9) & Null2String(rsCSMIOS_REPOR!Niym)
          rsCSMIOS_REPOR.MoveNext
       Loop
    End If
End If
If VarTranType = "VI" Then
    Set rsSMIS_PURCHAGREE = New ADODB.Recordset
    Set rsSMIS_PURCHAGREE = gconDMIS.Execute("Select * from SMIS_PurchAgree Where VI_NO = '" & VarTranno & "' order by VI_NO ASC")
    If Not rsSMIS_PURCHAGREE.EOF And Not rsSMIS_PURCHAGREE.BOF Then
       rsSMIS_PURCHAGREE.MoveFirst:
       Do While Not rsSMIS_PURCHAGREE.EOF
          LIM = LIM + 1
          If CheckSJExisting("VI", Null2String(rsSMIS_PURCHAGREE!VI_NO)) = True Then
             IS_Exist = 1
          Else
             IS_Exist = 0
          End If
          'Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsSMIS_PURCHAGREE!IGNKEY_NO) & Chr(9) & Null2String(rsSMIS_PURCHAGREE!VI_NO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsSMIS_PURCHAGREE!TOTAL)) & Chr(9) & SetCustomerName(Null2String(rsSMIS_PURCHAGREE!Code))
          Grid3.AddItem IS_Exist & Chr(9) & "SALES" & Chr(9) & Null2String(rsSMIS_PURCHAGREE!VI_NO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsSMIS_PURCHAGREE!TOTAL)) & Chr(9) & SetCustomerName(Null2String(rsSMIS_PURCHAGREE!code))
          rsSMIS_PURCHAGREE.MoveNext
       Loop
    End If
End If
Screen.MousePointer = 0
End Sub

Function SetCustomerName(VVV As Variant) As String
    Dim rsCustomer2                                    As ADODB.Recordset
    Set rsCustomer2 = New ADODB.Recordset
    rsCustomer2.Open "Select CustCode,acctname from ALL_CUSTMASTER_AMIS where CustCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer2.EOF And Not rsCustomer2.BOF Then
        SetCustomerName = UCase(Null2String(rsCustomer2!AcctName))
    Else
        SetCustomerName = ""
    End If
End Function

Sub ImportDeposited()
    Dim J_JDATE, J_VOUCHERNO, J_JTYPE   As String
    Dim J_JNO, J_REMARKS, J_VENDORCODE, J_CUSTOMERCODE As String
    Dim J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
    Dim J_CHECKNO                       As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE As String
    Dim J_INVOICETYPE, J_INVOICENO      As String
    Dim J_CHECKDATE, J_BANKCODE         As String
    Dim J_REFNO, J_REFDATE              As String
    Dim J_TERMS, J_DEALER               As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS   As String

    Dim J_ACCT_CODE, J_ACCT_NAME        As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET As Double
    Dim J_STATUS, J_JITEMNO             As String

    Dim rsJournal_HDDup                 As ADODB.Recordset

    Dim CMIS_OR_NUM                     As String
    Dim CMIS_OR_DATE                    As String
    Dim CMIS_OR_AMT                     As String
    Dim CMIS_DISCOUNT                   As String
    Dim CMIS_TAX                        As String
    Dim CMIS_CASHAMOUNT                 As Double
    Dim CMIS_CHKAMOUNT                  As Double
    Dim CMIS_CARDAMOUNT                 As Double
    Dim CMIS_CUSCDE                     As String
    Dim CMIS_CUSNAME                    As String
    Dim CMIS_DEPOSIT                    As String
    Dim CMIS_BANKCODE                   As String
    Dim CMIS_TSEKE                      As String
    Dim CMIS_CHECKDATE                  As String
    Dim CMIS_STATUS                     As String
    Dim CMIS_TYPE_PAYMENT               As String

    Dim CMIS_DT_TRANTYPE                As String
    Dim CMIS_DT_REFERENCE               As String
    Dim CMIS_DT_CUSCDE                  As String
    Dim CMIS_DT_DESCRIPT                As String
    Dim CMIS_DT_AMOUNT                  As Double
    Dim CMIS_DT_DOCDTE                  As String
    Dim CMIS_DT_PAYMENT                 As Double
    Dim CMIS_DT_DISCOUNT                As Double
    Dim CMIS_DT_TAX                     As Double
    Dim CMIS_DT_PAIDFOR                 As String
    Dim CMIS_IS_VAT                     As Boolean
    Dim CMIS_BANK_DEPOSITED             As String

    Dim TOTAL_DEBIT, TOTAL_CREDIT       As Double

    Dim rsOFF_HD                        As ADODB.Recordset
    Dim rsOFF_DT                        As ADODB.Recordset
    Dim I                               As Long

    Dim rsSJ_DATA                       As ADODB.Recordset

    Dim PV_PONO, PV_MRRNO, PV_INVNO, PV_PRODNO As String
    Dim J_JVOUCHERNO                    As String
    Dim PV_AMOUNT                       As Double
    Dim PV_STATUS, PV_ITEMNO            As String

    Dim SJ_PV_ITEMNO                    As Integer
    Dim rsCheckJournal_HD               As ADODB.Recordset

    Dim rsREPOR, rsORD_HD               As ADODB.Recordset

    Dim GridImport                      As Integer
    I = 0
    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then
            Set rsOFF_HD = New ADODB.Recordset
            If Grid2.Cell(GridImport, 2).Text = "VAT" Then
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD_Deposited Where DEPOSIT = 1 AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' AND OR_NUM = '" & Grid2.Cell(GridImport, 3).Text & "' AND VAT = 1 Order by OR_NUM ASC")
            Else
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD_Deposited Where DEPOSIT = 1 AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' AND OR_NUM = '" & Grid2.Cell(GridImport, 3).Text & "' AND VAT = 0 Order by OR_NUM ASC")
            End If
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                CMIS_OR_NUM = Null2String(rsOFF_HD!OR_NUM)
                CMIS_OR_DATE = Null2Date(rsOFF_HD!DATDEPOSIT)
                CMIS_OR_AMT = Null2String(rsOFF_HD!OR_AMT)
                CMIS_DISCOUNT = Null2String(rsOFF_HD!DISCOUNT)
                CMIS_TAX = Null2String(rsOFF_HD!TAX)
                CMIS_CASHAMOUNT = N2Str2Zero(rsOFF_HD!CASHAMOUNT)
                CMIS_CHKAMOUNT = N2Str2Zero(rsOFF_HD!CHKAMOUNT)
                CMIS_CARDAMOUNT = N2Str2Zero(rsOFF_HD!CARDAMOUNT)
                CMIS_CUSCDE = Null2String(rsOFF_HD!CUSCDE)
                CMIS_CUSNAME = Null2String(rsOFF_HD!CUSNAME)
                CMIS_DEPOSIT = Null2String(rsOFF_HD!DEPOSIT)
                CMIS_BANKCODE = Null2String(rsOFF_HD!BankCode)
                CMIS_TSEKE = Null2String(rsOFF_HD!TSEKE) & Null2String(rsOFF_HD!CARDNUMBER)
                CMIS_TYPE_PAYMENT = Null2String(rsOFF_HD!TOF)

                CMIS_CHECKDATE = Format(Null2Date(rsOFF_HD!CheckDate), "MM/DD/YYYY")
                CMIS_STATUS = Null2String(rsOFF_HD!Status)
                CMIS_IS_VAT = Null2Bool(rsOFF_HD!VAT)
                CMIS_BANK_DEPOSITED = Null2String(rsOFF_HD!BankAccountNo)
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0

                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If

                J_JDATE = N2Date2Null(CMIS_OR_DATE)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_JTYPE = "'DRJ'"

                Set rsOFF_DT = New ADODB.Recordset
                Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE OR_NUM = '" & CMIS_OR_NUM & "'")
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst
                    Do While Not rsOFF_DT.EOF
                        If Null2String(rsOFF_DT!TRANTYPE) = "OTH" Then
                            J_REMARKS = SetOtherTransaction(Null2String(rsOFF_DT!PAIDFOR)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!PAYMENT))
                        Else
                            J_REMARKS = SetTransaction(Null2String(rsOFF_DT!TRANTYPE)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!PAYMENT))
                        End If
                        rsOFF_DT.MoveNext
                        If Not rsOFF_DT.EOF Then J_REMARKS = "" & Chr(9)
                    Loop
                    J_REMARKS = N2Str2Null(J_REMARKS)
                Else
                    J_REMARKS = "NULL"
                End If
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(CMIS_CUSCDE)

                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = NumericVal(CMIS_OR_AMT)
                J_BALANCE = 0
                J_AMOUNTPAID = 0

                J_STATUS = "'N'"

                J_INVOICEDATE = N2Date2Null(CMIS_OR_DATE)
                If CMIS_IS_VAT = True Then
                    J_INVOICENO = N2Str2Null(Left(CMIS_OR_NUM, 6))
                Else
                    J_INVOICENO = N2Str2Null("NV" & Left(CMIS_OR_NUM, 6))
                End If
                J_CHECKNO = N2Str2Null(CMIS_TSEKE)
                J_DUEDATE = N2Date2Null(CMIS_CHECKDATE)
                If Null2String(rsOFF_HD!TOF) = "1" Then
                    J_PAYTYPE = "'CASH'"
                ElseIf Null2String(rsOFF_HD!TOF) = "2" Then
                    J_PAYTYPE = "'CHECK'"
                ElseIf Null2String(rsOFF_HD!TOF) = "3" Then
                    J_PAYTYPE = "'CARD'"
                Else
                    J_PAYTYPE = "NULL"
                End If
                J_INVOICETYPE = "'CI'"
                J_CHECKDATE = N2Date2Null(CMIS_CHECKDATE)
                J_BANKCODE = N2Str2Null(CMIS_BANKCODE)
                J_REFNO = N2Str2Null(CMIS_TSEKE)
                J_REFDATE = N2Date2Null(CMIS_CHECKDATE)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"

                If CMIS_TYPE_PAYMENT = "1" Or CMIS_TYPE_PAYMENT = "2" Then
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = N2Str2Null(CMIS_BANK_DEPOSITED)
                    J_ACCT_NAME = N2Str2Null(Setacctname(CMIS_BANK_DEPOSITED))
                    If CMIS_CASHAMOUNT > 0 Then
                        J_DEBIT = NumericVal(CMIS_CASHAMOUNT)
                    Else
                        J_DEBIT = NumericVal(CMIS_CHKAMOUNT)
                    End If
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    J_JITEMNO = "'0002'"
                    J_ACCT_CODE = N2Str2Null(COA_CASH_ON_HAND)
                    J_ACCT_NAME = N2Str2Null(Setacctname(COA_CASH_ON_HAND))
                    If CMIS_CASHAMOUNT > 0 Then
                        J_CREDIT = NumericVal(CMIS_CASHAMOUNT)
                    Else
                        J_CREDIT = NumericVal(CMIS_CHKAMOUNT)
                    End If
                    J_DEBIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CMIS_TYPE_PAYMENT = "3" Then
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = N2Str2Null(CMIS_BANK_DEPOSITED)
                    J_ACCT_NAME = N2Str2Null(Setacctname(CMIS_BANK_DEPOSITED))
                    J_DEBIT = NumericVal(CMIS_CARDAMOUNT)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    J_JITEMNO = "'0002'"
                    J_ACCT_CODE = N2Str2Null(COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD)
                    J_ACCT_NAME = N2Str2Null(Setacctname(COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD))
                    J_DEBIT = 0
                    J_CREDIT = NumericVal(CMIS_CARDAMOUNT)
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                               " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
            End If
            Grid2.Cell(GridImport, 1).Text = 1
        End If
        I = I + 1
        progCPB.Value = (I / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next

    Screen.MousePointer = 0
End Sub

Private Sub cmdClearJournals_Click()
    If Option1.Value = True Then
        Set rsCHATCheckControlIfExistRecordInJournalHD = New ADODB.Recordset
        Set rsCHATCheckControlIfExistRecordInJournalHD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'CRJ' and Jdate = '" & CDate(dtpTranDate) & "'")
        If Not rsCHATCheckControlIfExistRecordInJournalHD.EOF And Not rsCHATCheckControlIfExistRecordInJournalHD.BOF Then
            Screen.MousePointer = 0
            If LOGLEVEL = "ADM" Then
                If MsgBox("Clear Unposted Data for this Particular Date?", vbQuestion + vbYesNo, "Purge Data") = vbYes Then
                    Screen.MousePointer = 11
                    gconDMIS.Execute ("delete from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'CRJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    gconDMIS.Execute ("delete from AMIS_Journal_DET Where STATUS <> 'P' AND Jtype = 'CRJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    gconDMIS.Execute ("delete from AMIS_CRJ_Detail Where STATUS <> 'P' AND Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    cmdShowTrans.Value = True
                    Screen.MousePointer = 0
                    MsgBox "Existing Data Successfully deleted.", vbInformation, "Deleted"
                    Exit Sub
                End If
            End If
        End If
    End If
    If Option2.Value = True Then
        Set rsCHATCheckControlIfExistRecordInJournalHD = New ADODB.Recordset
        Set rsCHATCheckControlIfExistRecordInJournalHD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'DRJ' and Jdate = '" & CDate(dtpTranDate) & "'")
        If Not rsCHATCheckControlIfExistRecordInJournalHD.EOF And Not rsCHATCheckControlIfExistRecordInJournalHD.BOF Then
            Screen.MousePointer = 0
            If LOGLEVEL = "ADM" Then
                If MsgBox("Clear Unposted Data for this Particular Date?", vbQuestion + vbYesNo, "Purge Data") = vbYes Then
                    Screen.MousePointer = 11
                    gconDMIS.Execute ("delete from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'DRJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    gconDMIS.Execute ("delete from AMIS_Journal_DET Where STATUS <> 'P' AND Jtype = 'DRJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    cmdShowTrans.Value = True
                    Screen.MousePointer = 0
                    MsgBox "Existing Data Successfully deleted.", vbInformation, "Deleted"
                    Exit Sub
                End If
            End If
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Function SetTransaction(XXX As Variant) As String
    Dim rsSBOOKTransaction              As ADODB.Recordset
    Set rsSBOOKTransaction = New ADODB.Recordset
    Set rsSBOOKTransaction = gconDMIS.Execute("Select * from CMIS_SBOOK Where BOOK = 'A' and CODE = '" & XXX & "'")
    If Not rsSBOOKTransaction.EOF And Not rsSBOOKTransaction.BOF Then
        SetTransaction = Null2String(rsSBOOKTransaction!DESCNAME)
    End If
    Set rsSBOOKTransaction = Nothing
End Function

Function SetOtherTransaction(XXX As Variant) As String
    Dim rsSBOOKOtherTransaction         As ADODB.Recordset
    Set rsSBOOKOtherTransaction = New ADODB.Recordset
    Set rsSBOOKOtherTransaction = gconDMIS.Execute("Select * from CMIS_SBOOK Where BOOK = 'D' and CODE = '" & XXX & "'")
    If Not rsSBOOKOtherTransaction.EOF And Not rsSBOOKOtherTransaction.BOF Then
        SetOtherTransaction = Null2String(rsSBOOKOtherTransaction!DESCNAME)
    End If
    Set rsSBOOKOtherTransaction = Nothing
End Function

Function Setacctname(VVV As Variant) As String
    Dim rsChartAccount2                 As ADODB.Recordset
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

Function GetCRJVoucherNo() As String
    Dim rsJournal_HD                    As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'CRJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetCRJVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetCRJVoucherNo = "000001"
    End If
End Function

Function GetVoucherNo() As String
    Dim rsJournal_HD                    As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select  CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'DRJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function GetSJVoucherNo() As String
    Dim rsJournal_HD                                   As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'SJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetSJVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetSJVoucherNo = "000001"
    End If
End Function

Private Sub cmdShowTrans_Click()
    Screen.MousePointer = 11
    cmdCheck.Enabled = False: cmdClearJournals.Enabled = False: InitGrids: InitGrids: DoEvents
    If KIM > 0 Then
       cmdCheck.Enabled = True
       cmdClearJournals.Enabled = True
    End If
        
    Grid3.AutoRedraw = False
    Grid1.Rows = 2: Grid2.Rows = 2: Grid3.Rows = 2: KIM = 0: LIM = 0
    Dim ORType, ORTypeCode              As String
    Dim IS_Exist                        As Byte
    Dim rsOR_UNDEPOSITED                As ADODB.Recordset
    Dim rsOR_DEPOSITED                  As ADODB.Recordset
    
    Dim rsUNDEPOSITED_INVOICES As ADODB.Recordset
    Set rsOR_UNDEPOSITED = New ADODB.Recordset
    'Set rsOR_UNDEPOSITED = gconDMIS.Execute("Select * from CMIS_OFF_HD where OR_DATE = '" & CDate(dtpTranDate) & "' AND (DEPOSIT = FALSE OR DEPOSIT = 0) order by OR_NUM ASC")
    Set rsOR_UNDEPOSITED = gconDMIS.Execute("Select * from CMIS_OFF_HD where OR_DATE = '" & CDate(dtpTranDate) & "' order by OR_NUM ASC")
    If Not rsOR_UNDEPOSITED.EOF And Not rsOR_UNDEPOSITED.BOF Then
        rsOR_UNDEPOSITED.MoveFirst: KIM = 0: LIM = 0
        Grid1.AutoRedraw = False
        Do While Not rsOR_UNDEPOSITED.EOF
            KIM = KIM + 1
            If CheckCRJExisting(Null2String(rsOR_UNDEPOSITED!OR_NUM), N2Str2Zero(rsOR_UNDEPOSITED!VAT)) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            If N2Str2Zero(rsOR_UNDEPOSITED!VAT) = 1 Then
                ORType = "VAT"
            Else
                ORType = "NON VAT"
            End If
            Grid1.AddItem IS_Exist & Chr(9) & ORType & Chr(9) & Null2String(rsOR_UNDEPOSITED!OR_NUM) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOR_UNDEPOSITED!OR_AMT)) & Chr(9) & Null2String(rsOR_UNDEPOSITED!CUSNAME)
            Set rsUNDEPOSITED_INVOICES = New ADODB.Recordset
            Set rsUNDEPOSITED_INVOICES = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE (TRANTYPE = 'VI' OR TRANTYPE = 'SI' OR TRANTYPE = 'PI' OR TRANTYPE = 'AI' OR TRANTYPE = 'MI') AND OR_NUM = " & N2Str2Null(rsOR_UNDEPOSITED!OR_NUM) & " AND VAT = " & N2Str2Zero(rsOR_UNDEPOSITED!VAT))
            If Not rsUNDEPOSITED_INVOICES.EOF And Not rsUNDEPOSITED_INVOICES.BOF Then
               ShowUnImportedPaidInvoices Null2String(rsUNDEPOSITED_INVOICES!TRANTYPE), Null2String(rsUNDEPOSITED_INVOICES!InvoiceNo)
            End If
            rsOR_UNDEPOSITED.MoveNext
        Loop
        If KIM > 0 Then Grid1.RemoveItem 1
        Grid1.AutoRedraw = True
        Grid1.Refresh
    End If
    Set rsOR_DEPOSITED = New ADODB.Recordset
    Set rsOR_DEPOSITED = gconDMIS.Execute("Select * from CMIS_OFF_HD_Deposited Where DEPOSIT = 1 AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' Order by OR_NUM ASC")
    If Not rsOR_DEPOSITED.EOF And Not rsOR_DEPOSITED.BOF Then
        rsOR_DEPOSITED.MoveFirst: KIM = 0
        Grid2.AutoRedraw = False
        Do While Not rsOR_DEPOSITED.EOF
            KIM = KIM + 1
            If CheckDRJExisting(Null2String(rsOR_DEPOSITED!OR_NUM), N2Str2Zero(rsOR_DEPOSITED!VAT)) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            If N2Str2Zero(rsOR_DEPOSITED!VAT) = 1 Then
                ORType = "VAT"
            Else
                ORType = "NON VAT"
            End If
            Grid2.AddItem IS_Exist & Chr(9) & ORType & Chr(9) & Null2String(rsOR_DEPOSITED!OR_NUM) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOR_DEPOSITED!OR_AMT)) & Chr(9) & Null2String(rsOR_DEPOSITED!CUSNAME)
            rsOR_DEPOSITED.MoveNext
        Loop
        If KIM > 0 Then Grid2.RemoveItem 1
        Grid2.AutoRedraw = True
        Grid2.Refresh
    End If
    If KIM > 0 Then
        cmdCheck.Enabled = True
        cmdClearJournals.Enabled = True
    End If
    If LIM > 0 Then Grid3.RemoveItem 1
    Grid3.AutoRedraw = True
    Grid3.Refresh
    Screen.MousePointer = 0
End Sub

Private Sub dtpTranDate_Change()
    InitGrids
    InitGrid3
    DoEvents:
    Grid1.Rows = 1
    Grid2.Rows = 1
    Grid3.Rows = 1
    cmdCheck.Enabled = False
    cmdClearJournals.Enabled = False
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    dtpTranDate = LOGDATE
    InitGrids
    InitGrid3
    Option1_Click
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Database Connection Error!"
    Unload frmSplash
    cmdCheck.Enabled = False
End Sub

Sub InitGrids()
    With Grid1
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "OR Type"
        .Cell(0, 3).Text = "OR No."
        .Cell(0, 4).Text = "OR Amt."
        .Cell(0, 5).Text = "Customer"

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
        .Cell(0, 2).Text = "OR Type"
        .Cell(0, 3).Text = "OR No."
        .Cell(0, 4).Text = "OR Amt."
        .Cell(0, 5).Text = "Customer"

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
    InitGrid3
End Sub

Function CheckCRJExisting(VarInvoiceNo As String, VarVAT As Variant) As Boolean
    Dim rsCheckCRJ_Journal_HD           As ADODB.Recordset
    Set rsCheckCRJ_Journal_HD = New ADODB.Recordset
    If VarVAT = 0 Then
        Set rsCheckCRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'CRJ' AND LEFT(InvoiceNo,2) = 'NV' AND RIGHT(InvoiceNo,6) = " & N2Str2Null(VarInvoiceNo))
    Else
        Set rsCheckCRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'CRJ' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    End If
    If Not rsCheckCRJ_Journal_HD.EOF And Not rsCheckCRJ_Journal_HD.BOF Then
        CheckCRJExisting = True
    Else
        CheckCRJExisting = False
    End If
    Set rsCheckCRJ_Journal_HD = Nothing
End Function

Function CheckDRJExisting(VarInvoiceNo As String, VarVAT As Variant) As Boolean
    Dim rsCheckDRJ_Journal_HD           As ADODB.Recordset
    Set rsCheckDRJ_Journal_HD = New ADODB.Recordset
    If VarVAT = 0 Then
        Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'DRJ' AND LEFT(InvoiceNo,2) = 'NV' AND RIGHT(InvoiceNo,6) = " & N2Str2Null(VarInvoiceNo))
    Else
        Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'DRJ' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    End If
    If Not rsCheckDRJ_Journal_HD.EOF And Not rsCheckDRJ_Journal_HD.BOF Then
        CheckDRJExisting = True
    Else
        CheckDRJExisting = False
    End If
    Set rsCheckDRJ_Journal_HD = Nothing
End Function

Function CheckSJExisting(VarInvoiceType As String, VarInvoiceNo As String) As Boolean
Dim rsCheckSJ_Journal_HD As ADODB.Recordset
Set rsCheckSJ_Journal_HD = New ADODB.Recordset
Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
If Not rsCheckSJ_Journal_HD.EOF And Not rsCheckSJ_Journal_HD.BOF Then
   CheckSJExisting = True
Else
   CheckSJExisting = False
End If
Set rsCheckSJ_Journal_HD = Nothing
End Function

Sub ImportCSMSSales()
    Screen.MousePointer = 11
    Dim WARRANTY_JNO As String
    Dim WARRANTY_VOUCHERNO  As String
    Dim WARRANTY_ItemCnt As Integer
    Dim WARRANTY_J_JITEMNO As String
    
    Dim WARRANTY_J_AMOUNTTOPAY As Double
    Dim WARRANTY_J_INVOICEAMT As Double
    Dim WARRANTY_J_BALANCE As Double
    Dim WARRANTY_J_AMOUNTPAID As Double
        
    Dim rsCSMIOS_REPOR                                 As ADODB.Recordset
    Dim rsCSMIOS_TINSPAINT                             As ADODB.Recordset
    Dim rsCSMIOS_SUBLET                                As ADODB.Recordset
    Dim rsCSMIOS_AIRCON                                As ADODB.Recordset
    Dim rsCSMIOS_LABOR                                 As ADODB.Recordset
    Dim rsCSMIOS_PARTS                                 As ADODB.Recordset
    Dim rsCSMIOS_MATERIALS                             As ADODB.Recordset
    Dim rsCSMIOS_ACCESSORIES                           As ADODB.Recordset

    Dim CSMIOS_REP_OR                                  As String
    Dim CSMIOS_ACCT_NO                                 As String
    Dim CSMIOS_PLATE_NO                                As String
    Dim CSMIOS_NIYM                                    As String
    Dim CSMIOS_PARTICIPAT                              As String
    Dim CSMIOS_TERM                                    As String
    Dim CSMIOS_DTE_REL                                 As String
    Dim CSMIOS_INVOICE                                 As String

    Dim CSMIOS_LABOR                                   As Double
    Dim CSMIOS_PARTS                                   As Double
    Dim CSMIOS_MATERIALS                               As Double
    Dim CSMIOS_ACCESSORIES                             As Double

    Dim CSMIOS_LABOR_COST                              As Double
    Dim CSMIOS_PARTS_COST                              As Double
    Dim CSMIOS_MATERIALS_COST                          As Double
    Dim CSMIOS_ACCESSORIES_COST                        As Double

    Dim CSMIOS_RO_AMOUNT                               As Double

    Dim CSMIOS_TINSPAINT                               As Double
    Dim CSMIOS_SUBLET                                  As Double
    Dim CSMIOS_AIRCON                                  As Double
    Dim CSMIOS_PMS                                     As Double

    'FOR PDI
    Dim CSMIOS_PDI_LABOR                               As Double
    Dim CSMIOS_PDI_PARTS                               As Double
    Dim CSMIOS_PDI_MATERIALS                           As Double

    Dim CSMIOS_PDI_TINSPAINT                           As Double
    Dim CSMIOS_PDI_SUBLET                              As Double
    Dim CSMIOS_PDI_AIRCON                              As Double
    'END PDI
    
    Dim CSMIOS_TINSPAINT_DISCOUNT                      As Double
    Dim CSMIOS_SUBLET_DISCOUNT                         As Double
    Dim CSMIOS_AIRCON_DISCOUNT                         As Double

    Dim CSMIOS_LABOR_DISCOUNT                          As Double
    Dim CSMIOS_PARTS_DISCOUNT                          As Double
    Dim CSMIOS_MATERIALS_DISCOUNT                      As Double
    Dim CSMIOS_ACCESSORIES_DISCOUNT                    As Double

    Dim WARRANTY_DIRECT_EXPENSE_LABOR                  As Double
    Dim WARRANTY_DIRECT_EXPENSE_SPAREPARTS             As Double
    Dim WARRANTY_DIRECT_EXPENSE_GOL                    As Double
    Dim WARRANTY_DIRECT_EXPENSE_ACCESSORIES            As Double
    
    Dim WARRANTY_CSMIOS_PARTS_COST                     As Double
    Dim WARRANTY_CSMIOS_MATERIALS_COST                 As Double
    Dim WARRANTY_CSMIOS_ACCESSORIES_COST               As Double
                
    Dim COMPANY_DIRECT_EXPENSE_LABOR                   As Double
    Dim COMPANY_DIRECT_EXPENSE_SPAREPARTS              As Double
    Dim COMPANY_DIRECT_EXPENSE_GOL                     As Double
    Dim COMPANY_DIRECT_EXPENSE_ACCESSORIES             As Double

    Dim SALES_DIRECT_EXPENSE_LABOR                     As Double
    Dim SALES_DIRECT_EXPENSE_SPAREPARTS                As Double
    Dim SALES_DIRECT_EXPENSE_GOL                       As Double
    Dim SALES_DIRECT_EXPENSE_ACCESSORIES               As Double

    Dim INSURANCE_DIRECT_EXPENSE_LABOR                 As Double
    Dim INSURANCE_DIRECT_EXPENSE_SPAREPARTS            As Double
    Dim INSURANCE_DIRECT_EXPENSE_GOL                   As Double
    Dim INSURANCE_DIRECT_EXPENSE_ACCESSORIES           As Double
    Dim TOTAL_INSURANCE_AMOUNT                         As Double
    
    Dim CSMS_VAT_EXEMPT                                As Boolean
    Dim ItemCnt                                        As Integer
        
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            If Grid3.Cell(GridImport, 2).Text = "SERVICE" Then
            Else
               GoTo NextGrid
            End If
            Set rsCSMIOS_REPOR = New ADODB.Recordset
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where invoice = '" & Grid3.Cell(GridImport, 3).Text & "' order by invoice ASC")
            If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
                ItemCnt = 0
                CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
                
                CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
    
                CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!Plate_No)
                CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)
                CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!DTE_COMP)
                CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!Invoice)
                
                CSMS_VAT_EXEMPT = Null2Bool(rsCSMIOS_REPOR!VAT_EXEMPT)
                
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If
    
                CSMIOS_RO_AMOUNT = Round(N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT), 2)
            
                
                '=======================================================================================================================================================================
                'CUSTOMER
                
                'LABOR - MECHANICAL / BODY AND PAINT / AIRCON
                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABOR Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                    CSMIOS_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
                    CSMIOS_LABOR_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_LABOR!DISCOUNT), 2)
                Else
                    CSMIOS_LABOR = 0: CSMIOS_LABOR_DISCOUNT = 0
                End If
                
                Set rsCSMIOS_TINSPAINT = New ADODB.Recordset
                Set rsCSMIOS_TINSPAINT = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS TINSPAINT,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_TINSPAINT Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_TINSPAINT.EOF And Not rsCSMIOS_TINSPAINT.BOF Then
                    CSMIOS_TINSPAINT = Round(N2Str2Zero(rsCSMIOS_TINSPAINT!TINSPAINT), 2)
                    CSMIOS_TINSPAINT_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_TINSPAINT!DISCOUNT), 2)
                Else
                    CSMIOS_TINSPAINT = 0: CSMIOS_TINSPAINT_DISCOUNT = 0
                End If
    
                Set rsCSMIOS_AIRCON = New ADODB.Recordset
                Set rsCSMIOS_AIRCON = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS AIRCON,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_AIRCON Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_AIRCON.EOF And Not rsCSMIOS_AIRCON.BOF Then
                    CSMIOS_AIRCON = Round(N2Str2Zero(rsCSMIOS_AIRCON!AIRCON), 2)
                    CSMIOS_AIRCON_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_AIRCON!DISCOUNT), 2)
                Else
                    CSMIOS_AIRCON = 0: CSMIOS_AIRCON_DISCOUNT = 0
                End If

'                Set rsCSMIOS_SUBLET = New ADODB.Recordset
'                Set rsCSMIOS_SUBLET = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS SUBLET,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_SUBLET Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
'                If Not rsCSMIOS_SUBLET.EOF And Not rsCSMIOS_SUBLET.BOF Then
'                    CSMIOS_SUBLET = N2Str2Zero(rsCSMIOS_SUBLET!SUBLET)
'                    CSMIOS_SUBLET_DISCOUNT = N2Str2Zero(rsCSMIOS_SUBLET!DISCOUNT)
'                Else
'                    CSMIOS_SUBLET = 0: CSMIOS_SUBLET_DISCOUNT = 0
'                End If

        
                'PARTS
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTS Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                    CSMIOS_PARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
                    CSMIOS_PARTS_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_PARTS!DISCOUNT), 2)
                Else
                    CSMIOS_PARTS = 0: CSMIOS_PARTS_DISCOUNT = 0
                End If
    
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS PARTS_COST from CSMIOS_vw_PARTSCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                    CSMIOS_PARTS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS_COST), 2)
                Else
                    CSMIOS_PARTS_COST = 0:
                End If
    
                'MATERIALS
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALS Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                    CSMIOS_MATERIALS = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
                    CSMIOS_MATERIALS_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_MATERIALS!DISCOUNT), 2)
                Else
                    CSMIOS_MATERIALS = 0: CSMIOS_MATERIALS_DISCOUNT = 0
                End If
                
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS MAT_COST from CSMIOS_vw_MATCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                    CSMIOS_MATERIALS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!MAT_COST), 2)
                Else
                    CSMIOS_MATERIALS_COST = 0:
                End If
    
                'ACCESSORIES
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_ACCESSORIES Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                    CSMIOS_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES), 2)
                    CSMIOS_ACCESSORIES_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!DISCOUNT), 2)
                Else
                    CSMIOS_ACCESSORIES = 0: CSMIOS_ACCESSORIES_DISCOUNT = 0
                End If
                
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS ACC_COST from CSMIOS_vw_ACCCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                    CSMIOS_ACCESSORIES_COST = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACC_COST), 2)
                Else
                    CSMIOS_ACCESSORIES_COST = 0:
                End If
    
                '=======================================================================================================================================================================
                'WARRANTY
                
                WARRANTY_DIRECT_EXPENSE_LABOR = 0: WARRANTY_DIRECT_EXPENSE_SPAREPARTS = 0: WARRANTY_DIRECT_EXPENSE_GOL = 0: WARRANTY_DIRECT_EXPENSE_ACCESSORIES = 0
                WARRANTY_CSMIOS_PARTS_COST = 0: WARRANTY_CSMIOS_MATERIALS_COST = 0: WARRANTY_CSMIOS_ACCESSORIES_COST = 0
                
                'LABOR
                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORWarranty Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                   WARRANTY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
                End If
    
                'PARTS
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSWarranty Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                   WARRANTY_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
                End If
                
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS PARTS_COST from CSMIOS_vw_WARRANTY_PARTSCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                    WARRANTY_CSMIOS_PARTS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS_COST), 2)
                Else
                    WARRANTY_CSMIOS_PARTS_COST = 0:
                End If
                
                'MATERIALS
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSWarranty Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                   WARRANTY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
                End If
        
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS MAT_COST from CSMIOS_vw_WARRANTY_MATCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                    WARRANTY_CSMIOS_MATERIALS_COST = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MAT_COST), 2)
                Else
                    WARRANTY_CSMIOS_MATERIALS_COST = 0:
                End If
    
                'ACCESSORIES
                
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_ACCESSORIESWarranty Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                   WARRANTY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES), 2)
                End If
                
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS ACC_COST from CSMIOS_vw_WARRANTY_ACCCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                    WARRANTY_CSMIOS_ACCESSORIES_COST = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACC_COST), 2)
                Else
                    WARRANTY_CSMIOS_ACCESSORIES_COST = 0:
                End If
        
                '=======================================================================================================================================================================
                
                '=======================================================================================================================================================================
                'COMPANY
                
                COMPANY_DIRECT_EXPENSE_LABOR = 0: COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0: COMPANY_DIRECT_EXPENSE_GOL = 0: COMPANY_DIRECT_EXPENSE_ACCESSORIES = 0
                
                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                   COMPANY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
                End If
                
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                   COMPANY_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
                End If
                
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                   COMPANY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
                End If
                
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_ACCESSORIESCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                   COMPANY_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES), 2)
                End If
                
                '=======================================================================================================================================================================
    
                '=======================================================================================================================================================================
                'SALES
                
                SALES_DIRECT_EXPENSE_LABOR = 0: SALES_DIRECT_EXPENSE_SPAREPARTS = 0: SALES_DIRECT_EXPENSE_GOL = 0: SALES_DIRECT_EXPENSE_ACCESSORIES = 0
    
                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                   SALES_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
                End If
    
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                   SALES_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
                End If
                
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                   SALES_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
                End If
    
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_ACCESSORIESSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                   SALES_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES), 2)
                End If
    
                '=======================================================================================================================================================================
    
                '=======================================================================================================================================================================
                'INSURANCE - HAI OLD
                     
                'INSURANCE_DIRECT_EXPENSE_LABOR = 0: INSURANCE_DIRECT_EXPENSE_SPAREPARTS = 0: INSURANCE_DIRECT_EXPENSE_GOL = 0
                
                'Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                'Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select * from CSMIOS_INSURANCE Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                'If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                '   INSURANCE_DIRECT_EXPENSE_LABOR = N2Str2Zero(rsCSMIOS_MATERIALS!INSLABOR)
                '   INSURANCE_DIRECT_EXPENSE_GOL = N2Str2Zero(rsCSMIOS_MATERIALS!INSMATERIALS)
                '   INSURANCE_DIRECT_EXPENSE_SPAREPARTS = N2Str2Zero(rsCSMIOS_MATERIALS!INSPARTS)
                'End If
                
                '====================================================================================================================================================================================
                'INSURANCE - FML 03122008 11:35 AM
                
                INSURANCE_DIRECT_EXPENSE_LABOR = 0: INSURANCE_DIRECT_EXPENSE_SPAREPARTS = 0: INSURANCE_DIRECT_EXPENSE_GOL = 0: INSURANCE_DIRECT_EXPENSE_ACCESSORIES = 0
                TOTAL_INSURANCE_AMOUNT = 0
                
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select * from CSMIOS_INSURANCE Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                   INSURANCE_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSLABOR), 2)
                   INSURANCE_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSMATERIALS), 2)
                   INSURANCE_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSPARTS), 2)
                   INSURANCE_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSACCESSORIES), 2)
                   
                   If (CSMIOS_LABOR + CSMIOS_SUBLET + CSMIOS_TINSPAINT + CSMIOS_PMS) - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                        If CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                           CSMIOS_LABOR = Round(CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR, 2)
                           GoTo PAKSIW
                        Else
                           If CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                              CSMIOS_LABOR = Round(CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR, 2)
                              GoTo PAKSIW
                           Else
                              INSURANCE_DIRECT_EXPENSE_LABOR = Round(INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR, 2)
                              CSMIOS_LABOR = 0
                           End If
                        End If
                        If CSMIOS_SUBLET > 0 And CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                           CSMIOS_SUBLET = Round(CSMIOS_SUBLET - (INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR), 2)
                           GoTo PAKSIW
                        Else
                           If CSMIOS_SUBLET - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                              CSMIOS_SUBLET = Round(CSMIOS_SUBLET - INSURANCE_DIRECT_EXPENSE_LABOR, 2)
                              GoTo PAKSIW
                           Else
                              INSURANCE_DIRECT_EXPENSE_LABOR = Round(INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_SUBLET, 2)
                              CSMIOS_SUBLET = 0
                           End If
                        End If
                        If CSMIOS_TINSPAINT > 0 And CSMIOS_LABOR - CSMIOS_SUBLET - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                           CSMIOS_TINSPAINT = Round(CSMIOS_TINSPAINT - (INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR - CSMIOS_SUBLET), 2)
                           GoTo PAKSIW
                        Else
                           If CSMIOS_TINSPAINT - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                              CSMIOS_TINSPAINT = Round(CSMIOS_TINSPAINT - INSURANCE_DIRECT_EXPENSE_LABOR, 2)
                              GoTo PAKSIW
                           Else
                              INSURANCE_DIRECT_EXPENSE_LABOR = Round(INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_TINSPAINT, 2)
                              CSMIOS_TINSPAINT = 0
                           End If
                        End If
                        If CSMIOS_PMS > 0 And CSMIOS_LABOR - CSMIOS_SUBLET - CSMIOS_TINSPAINT - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                           CSMIOS_PMS = Round(CSMIOS_PMS - (INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR - CSMIOS_SUBLET - CSMIOS_TINSPAINT), 2)
                           GoTo PAKSIW
                        Else
                           If CSMIOS_PMS - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                              CSMIOS_PMS = Round(CSMIOS_PMS - INSURANCE_DIRECT_EXPENSE_LABOR, 2)
                           Else
                              INSURANCE_DIRECT_EXPENSE_LABOR = Round(INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_PMS, 2)
                              CSMIOS_PMS = 0
                           End If
                        End If
PAKSIW:                 INSURANCE_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSLABOR), 2)
                   Else
                       CSMIOS_LABOR = 0
                       CSMIOS_SUBLET = 0
                       CSMIOS_TINSPAINT = 0
                       CSMIOS_PMS = 0
                   End If
                   If CSMIOS_PARTS > 0 Then
                      CSMIOS_PARTS = Round(CSMIOS_PARTS - INSURANCE_DIRECT_EXPENSE_SPAREPARTS, 2)
                   End If
                   If CSMIOS_MATERIALS > 0 Then
                      CSMIOS_MATERIALS = Round(CSMIOS_MATERIALS - INSURANCE_DIRECT_EXPENSE_GOL, 2)
                   End If
                   If CSMIOS_ACCESSORIES > 0 Then
                      CSMIOS_ACCESSORIES = Round(CSMIOS_ACCESSORIES - INSURANCE_DIRECT_EXPENSE_ACCESSORIES, 2)
                   End If
                   
                   TOTAL_INSURANCE_AMOUNT = Round(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES, 2)
                End If
                '====================================================================================================================================================================================
                
                '=======================================================================================================================================================================
                
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
    
                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                
                J_JDATE = N2Date2Null(CSMIOS_DTE_REL)
                J_VOUCHERNO = N2Str2Null(GetSJVoucherNo())
                J_JTYPE = "'SJ'"
                J_REMARKS = "NULL"
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)
    
                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0
    
                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = Round(NumericVal(CSMIOS_RO_AMOUNT - TOTAL_INSURANCE_AMOUNT), 2)
                J_BALANCE = Round(NumericVal(CSMIOS_RO_AMOUNT - TOTAL_INSURANCE_AMOUNT), 2)
                J_AMOUNTPAID = 0
    
                J_STATUS = "'N'"
    
                J_INVOICEDATE = N2Date2Null(CSMIOS_DTE_REL)
                J_INVOICENO = N2Str2Null(CSMIOS_INVOICE)
                J_CHECKNO = "NULL"
                J_DUEDATE = N2Date2Null(CSMIOS_DTE_REL)
                J_PAYTYPE = "NULL"
                J_INVOICETYPE = "'SI'"
                J_CHECKDATE = "NULL"
                J_BANKCODE = "NULL"
                J_REFNO = N2Str2Null(CSMIOS_REP_OR)
                J_REFDATE = N2Date2Null(CSMIOS_DTE_REL)
                J_TERMS = N2Str2Null(CSMIOS_TERM)
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"
    
                If J_INVOICEAMT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_ACCT_CODE = N2Str2Null(COA_PRE_DELIVERY)
                        J_ACCT_NAME = N2Str2Null(Setacctname(COA_PRE_DELIVERY))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                    End If
                    J_DEBIT = Round(J_INVOICEAMT, 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_LABOR > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR")))
                    End If
                    J_GROSS = NumericVal(CSMIOS_LABOR)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_LABOR), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_LABOR), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_LABOR / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_LABOR) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_LABOR_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "LABOR"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "LABOR")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "LABOR"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "LABOR")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_LABOR_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_LABOR_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_LABOR_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_LABOR_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_LABOR_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_TINSPAINT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "BODY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "BODY")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "BODY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "BODY")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_TINSPAINT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_TINSPAINT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_TINSPAINT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_TINSPAINT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_TINSPAINT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_TINSPAINT_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "BODY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "BODY")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "BODY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "BODY")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_TINSPAINT_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_TINSPAINT_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_TINSPAINT_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_TINSPAINT_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_TINSPAINT_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_SUBLET > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "SUBLET"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "SUBLET")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "SUBLET"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "SUBLET")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_SUBLET), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_SUBLET), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_SUBLET), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_SUBLET / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_SUBLET) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                End If
                If CSMIOS_SUBLET_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "SUBLET"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "SUBLET")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "SUBLET"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "SUBLET")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_SUBLET_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_SUBLET_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_SUBLET_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_SUBLET_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_SUBLET_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_AIRCON > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "AIRCON"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "AIRCON")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "AIRCON"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "AIRCON")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_AIRCON), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_AIRCON), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_AIRCON), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_AIRCON / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_AIRCON) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_AIRCON_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "AIRCON"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "AIRCON")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "AIRCON"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "AIRCON")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_AIRCON_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_AIRCON_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_AIRCON_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_AIRCON_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_AIRCON_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                               
                If CSMIOS_PARTS > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                    End If
                    J_GROSS = NumericVal(CSMIOS_PARTS)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_PARTS), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_PARTS), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_PARTS / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_PARTS) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "PARTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "PARTS")))
                    End If
                    J_DEBIT = Round(CSMIOS_PARTS_COST, 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(CSMIOS_PARTS_COST, 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                
                If CSMIOS_PARTS_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "PARTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "PARTS")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_PARTS_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_MATERIALS > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_MATERIALS), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_MATERIALS), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_MATERIALS), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_MATERIALS / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_MATERIALS) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                     
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LUBRICANTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LUBRICANTS")))
                    End If
                    J_GROSS = 0
                    J_TAX = 0
                    J_NET = 0
                    J_DEBIT = Round(CSMIOS_MATERIALS_COST, 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("LUBRICANTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("LUBRICANTS")))
                    End If
                    J_GROSS = 0
                    J_TAX = 0
                    J_NET = 0
                    J_DEBIT = 0
                    J_CREDIT = Round(CSMIOS_MATERIALS_COST, 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_MATERIALS_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_MATERIALS_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
    
                'ACCESSORIES
                If CSMIOS_ACCESSORIES > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_ACCESSORIES / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_ACCESSORIES) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                     
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS", "ACCESSORIES")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS", "ACCESSORIES")))
                    End If
                    J_GROSS = 0
                    J_TAX = 0
                    J_NET = 0
                    J_DEBIT = Round(CSMIOS_ACCESSORIES_COST, 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                    End If
                    J_GROSS = 0
                    J_TAX = 0
                    J_NET = 0
                    J_DEBIT = 0
                    J_CREDIT = Round(CSMIOS_ACCESSORIES_COST, 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_ACCESSORIES_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_ACCESSORIES_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
    
                If J_INVOICEAMT > 0 Then
                    If CSMS_VAT_EXEMPT = False Then
                       ItemCnt = ItemCnt + 1
                       J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnOutPutTax())
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutPutTax()))
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(Round((J_INVOICEAMT / 1.12), 2) * 0.12), 2)
                       J_TAX = 0
                       J_GROSS = 0
                       J_NET = 0
                       J_STATUS = "'N'"
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
        
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    End If
                    
                    gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                                     " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                     " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                     ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                
                End If
                
                '==================================================================================================================================================================================
                'WARRANTY
                If WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES > 0 Then
             
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 2, "000000") & "'"
                    
                    WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetSJVoucherNo()) + 1, "000000"))
                    WARRANTY_ItemCnt = 0
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("WARRANTY"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("WARRANTY")))
                    J_DEBIT = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                    If WARRANTY_DIRECT_EXPENSE_LABOR > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR")))
                       J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR), 2)
                       J_TAX = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR) / 9.3333, 2)
                       J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
    
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    End If
                    If WARRANTY_DIRECT_EXPENSE_SPAREPARTS > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                       J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS), 2)
                       J_TAX = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS) / 9.3333, 2)
                       J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "PARTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "PARTS")))
                        End If
                        J_DEBIT = Round(WARRANTY_CSMIOS_PARTS_COST, 2)
                        J_CREDIT = 0
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                        End If
                        J_DEBIT = 0
                        J_CREDIT = Round(WARRANTY_CSMIOS_PARTS_COST, 2)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                    End If
                    If WARRANTY_DIRECT_EXPENSE_GOL > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                       J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL), 2)
                       J_TAX = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL) / 9.3333, 2)
                       J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = NumericVal(J_NET)
                       J_STATUS = "'N'"
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LUBRICANTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LUBRICANTS")))
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = Round(WARRANTY_CSMIOS_MATERIALS_COST, 2)
                        J_CREDIT = 0
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("LUBRICANTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("LUBRICANTS")))
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = 0
                        J_CREDIT = Round(WARRANTY_CSMIOS_MATERIALS_COST, 2)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        
                    End If
    
                    If WARRANTY_DIRECT_EXPENSE_ACCESSORIES > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                       J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_ACCESSORIES), 2)
                       J_TAX = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_ACCESSORIES) / 9.3333, 2)
                       J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_ACCESSORIES) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = NumericVal(J_NET)
                       J_STATUS = "'N'"
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS", "ACCESSORIES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS", "ACCESSORIES")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS", "ACCESSORIES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS", "ACCESSORIES")))
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = Round(WARRANTY_CSMIOS_ACCESSORIES_COST, 2)
                        J_CREDIT = 0
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = 0
                        J_CREDIT = Round(WARRANTY_CSMIOS_ACCESSORIES_COST, 2)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        
                    End If
    
                    If NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES) > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = COA_OUTPUT_TAX
                       J_ACCT_NAME = N2Str2Null(Setacctname(COA_OUTPUT_TAX))
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES) / 9.3333), 2)
                       J_TAX = 0
                       J_GROSS = 0
                       J_NET = 0
                       J_STATUS = "'N'"
                         
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    End If
                    
                    CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                    CSMIOS_ACCT_NO = Null2String("H00001")
                    
                    CSMIOS_PARTICIPAT = ""
        
                    CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!Plate_No)
                    CSMIOS_NIYM = Null2String(SetCustomerName("H00001"))
                    CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                    CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!DTE_COMP)
                    CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!Invoice)
                    
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                        J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    Else
                        J_JNO = "'000001'"
                    End If
        
                    CSMIOS_RO_AMOUNT = N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT)
                    
                    J_JDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_JTYPE = "'SJ'"
                    J_REMARKS = "NULL"
                    J_VENDORCODE = "'999999'"
                    J_CUSTOMERCODE = N2Str2Null("H00001")
        
                    J_DEBIT = 0
                    J_CREDIT = 0
                    J_TAX = 0
                    J_OUTBALANCE = 0
        
                    J_AMOUNTTOPAY = 0
                    J_INVOICEAMT = Round(CSMIOS_RO_AMOUNT, 2)
                    J_BALANCE = Round(CSMIOS_RO_AMOUNT, 2)
                    J_AMOUNTPAID = 0
        
                    J_STATUS = "'N'"
        
                    J_INVOICEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_INVOICENO = N2Str2Null(CSMIOS_INVOICE)
                    J_CHECKNO = "NULL"
                    J_DUEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_PAYTYPE = "NULL"
                    J_INVOICETYPE = "'SI'"
                    J_CHECKDATE = "NULL"
                    J_BANKCODE = "NULL"
                    J_REFNO = N2Str2Null(CSMIOS_REP_OR)
                    J_REFDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_TERMS = N2Str2Null(CSMIOS_TERM)
                    J_DEALER = "NULL"
                    J_PAIDSTATUS = "'N'"
                    J_RECEIVESTATUS = "'N'"
                    
                    WARRANTY_J_AMOUNTTOPAY = 0
                    WARRANTY_J_INVOICEAMT = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES), 2)
                    WARRANTY_J_BALANCE = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES), 2)
                    WARRANTY_J_AMOUNTPAID = 0
                    gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                                     " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                     " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & WARRANTY_J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & Round(WARRANTY_J_AMOUNTTOPAY, 2) & "," & WARRANTY_J_BALANCE & "," & WARRANTY_J_AMOUNTPAID & _
                                     ", " & WARRANTY_JNO & ", " & Round(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES, 2) & ", " & Round(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES, 2) & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                End If
                '==================================================================================================================================================================================
                        
                '==================================================================================================================================================================================
                'INSURANCE
                
                If INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES > 0 Then
    
             
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 2, "000000") & "'"
                    
                    WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetSJVoucherNo()) + 1, "000000"))
                    WARRANTY_ItemCnt = 0
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("INSURANCE"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("INSURANCE")))
                    J_DEBIT = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                    If INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR")))
                       J_GROSS = NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR)
                       J_TAX = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR) / 9.3333, 2)
                       J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
    
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    End If
                    If INSURANCE_DIRECT_EXPENSE_SPAREPARTS > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                       J_GROSS = NumericVal(INSURANCE_DIRECT_EXPENSE_SPAREPARTS)
                       J_TAX = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_SPAREPARTS) / 9.3333, 2)
                       J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_SPAREPARTS) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        
                    
                    End If
                    If INSURANCE_DIRECT_EXPENSE_GOL > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                       J_GROSS = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_GOL), 2)
                       J_TAX = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_GOL) / 9.3333, 2)
                       J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_GOL) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        
                    End If
    
                    If INSURANCE_DIRECT_EXPENSE_ACCESSORIES > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                       J_GROSS = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                       J_TAX = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_ACCESSORIES) / 9.3333, 2)
                       J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_ACCESSORIES) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        
                    End If
    
                    If NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES) > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = COA_OUTPUT_TAX
                       J_ACCT_NAME = N2Str2Null(Setacctname(COA_OUTPUT_TAX))
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES) / 9.3333), 2)
                       J_TAX = 0
                       J_GROSS = 0
                       J_NET = 0
                       J_STATUS = "'N'"
                         
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    End If
                    
                    CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                    CSMIOS_ACCT_NO = Null2String("H00001")
                        
                    CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!Plate_No)
                    CSMIOS_NIYM = Null2String(SetCustomerName(CSMIOS_PARTICIPAT))
                    CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                    CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!DTE_COMP)
                    CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!Invoice)
                    
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                        J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    Else
                        J_JNO = "'000001'"
                    End If
        
                    CSMIOS_RO_AMOUNT = N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT)
                    
                    J_JDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                    J_JTYPE = "'SJ'"
                    J_REMARKS = "NULL"
                    J_VENDORCODE = "'999999'"
        
                    J_DEBIT = 0
                    J_CREDIT = 0
                    J_TAX = 0
                    J_OUTBALANCE = 0
        
                    J_AMOUNTTOPAY = 0
                    J_INVOICEAMT = CSMIOS_RO_AMOUNT
                    J_BALANCE = CSMIOS_RO_AMOUNT
                    J_AMOUNTPAID = 0
        
                    J_STATUS = "'N'"
        
                    J_INVOICEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_INVOICENO = N2Str2Null(CSMIOS_INVOICE)
                    J_CHECKNO = "NULL"
                    J_DUEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_PAYTYPE = "NULL"
                    J_INVOICETYPE = "'SI'"
                    J_CHECKDATE = "NULL"
                    J_BANKCODE = "NULL"
                    J_REFNO = N2Str2Null(CSMIOS_REP_OR)
                    J_REFDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_TERMS = N2Str2Null(CSMIOS_TERM)
                    J_DEALER = "NULL"
                    J_PAIDSTATUS = "'N'"
                    J_RECEIVESTATUS = "'N'"
                    
                    WARRANTY_J_AMOUNTTOPAY = 0
                    WARRANTY_J_INVOICEAMT = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                    WARRANTY_J_BALANCE = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                    WARRANTY_J_AMOUNTPAID = 0
                    gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                                     " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                     " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & N2Str2Null(CSMIOS_PARTICIPAT) & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & WARRANTY_J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & WARRANTY_J_AMOUNTTOPAY & "," & WARRANTY_J_BALANCE & "," & WARRANTY_J_AMOUNTPAID & _
                                     ", " & WARRANTY_JNO & ", " & INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL & ", " & INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                End If
                '==================================================================================================================================================================================
                
                Grid3.Cell(GridImport, 1).Text = 1
            End If
        End If
NextGrid:
    Next
    Screen.MousePointer = 0: DoEvents
End Sub

Sub ImportPMISSales()
    Screen.MousePointer = 11
    Dim PMIOS_TRANTYPE                                 As String
    Dim PMIOS_TRANNO                                   As String
    Dim PMIOS_TRANDATE                                 As String
    Dim PMIOS_cuscde                                   As String
    Dim PMIOS_AcctName                                 As String
    Dim PMIOS_TTLINVAMT                                As Double
    Dim PMIOS_DS_AMT1                                  As Double
    Dim PMIOS_NETINVAMT                                As Double
    Dim PMIOS_NETCOST                                As Double

    Dim PMIOS_TYPE                                     As String
    
    
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            Set rsPMIOS_ORD_HD = New ADODB.Recordset
            If Grid3.Cell(GridImport, 2).Text = UCase("Accessories") Then
               Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'A' AND TranType = '" & Left(Grid3.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid3.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
            ElseIf Grid3.Cell(GridImport, 2).Text = UCase("Materials") Then
               Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'M' AND TranType = '" & Left(Grid3.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid3.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
            ElseIf Grid3.Cell(GridImport, 2).Text = UCase("Parts") Then
               Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'P' AND TranType = '" & Left(Grid3.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid3.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
            Else
               GoTo NextGrid
            End If
            
            If Not rsPMIOS_ORD_HD.EOF And Not rsPMIOS_ORD_HD.BOF Then
                PMIOS_TRANTYPE = Null2String(rsPMIOS_ORD_HD!TRANTYPE)
                PMIOS_TRANNO = Null2String(rsPMIOS_ORD_HD!tranno)
                PMIOS_TRANDATE = Null2String(rsPMIOS_ORD_HD!TRANDATE)
                PMIOS_cuscde = Null2String(rsPMIOS_ORD_HD!CUSTCODE)
                PMIOS_AcctName = SetCustomerName(rsPMIOS_ORD_HD!CUSTCODE)
                PMIOS_TTLINVAMT = Round(N2Str2Zero(rsPMIOS_ORD_HD!TTLINVAMT), 2)
                PMIOS_DS_AMT1 = Round(N2Str2Zero(rsPMIOS_ORD_HD!DS_AMT1), 2)
                PMIOS_NETINVAMT = Round(N2Str2Zero(rsPMIOS_ORD_HD!NetInvAmt), 2)
                PMIOS_NETCOST = Round(N2Str2Zero(rsPMIOS_ORD_HD!NETCOST), 2)
                PMIOS_TYPE = Null2String(rsPMIOS_ORD_HD!Type)
    
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If
    
                J_JDATE = N2Date2Null(PMIOS_TRANDATE)
                J_VOUCHERNO = N2Str2Null(GetSJVoucherNo())
                J_JTYPE = "'SJ'"
                J_REMARKS = "NULL"
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(PMIOS_cuscde)
    
                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0
    
                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = Round(PMIOS_NETINVAMT, 2)
                J_BALANCE = Round(PMIOS_NETINVAMT, 2)
                J_AMOUNTPAID = 0
    
                J_STATUS = "'N'"
    
                J_INVOICEDATE = N2Date2Null(PMIOS_TRANDATE)
                J_INVOICENO = N2Str2Null(PMIOS_TRANNO)
                J_CHECKNO = "NULL"
                J_DUEDATE = N2Date2Null(PMIOS_TRANDATE)
                J_PAYTYPE = "NULL"
                If PMIOS_TYPE = "P" Then
                   J_INVOICETYPE = "'PI'"
                End If
                If PMIOS_TYPE = "A" Then
                   J_INVOICETYPE = "'AI'"
                End If
                If PMIOS_TYPE = "M" Then
                   J_INVOICETYPE = "'MI'"
                End If
                J_CHECKDATE = "NULL"
                J_BANKCODE = "NULL"
                J_REFNO = "NULL"
                J_REFDATE = N2Date2Null(PMIOS_TRANDATE)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"
    
                J_JITEMNO = "'0001'"
                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                J_DEBIT = Round(NumericVal(PMIOS_NETINVAMT), 2)
                J_CREDIT = 0
                J_TAX = 0
                J_GROSS = 0
                J_NET = 0
                J_STATUS = "'N'"
                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                J_JITEMNO = "'0002'"
                If PMIOS_TYPE = "P" Then
                   J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS")))
                End If
                If PMIOS_TYPE = "A" Then
                   J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("ACCESSORIES"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("ACCESSORIES")))
                End If
                If PMIOS_TYPE = "M" Then
                   J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("MATERIALS"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("MATERIALS")))
                End If
                J_GROSS = Round(NumericVal(PMIOS_TTLINVAMT), 2)
                J_TAX = Round(NumericVal(Round((PMIOS_TTLINVAMT / 1.12), 2) * 0.12), 2)
                J_NET = Round(NumericVal(PMIOS_TTLINVAMT) - NumericVal(J_TAX), 2)
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(J_NET), 2)
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                If PMIOS_DS_AMT1 > 0 Then
                    J_JITEMNO = "'0003'"
                    If PMIOS_TYPE = "P" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS")))
                    End If
                    If PMIOS_TYPE = "A" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("ACCESSORIES")))
                    End If
                    If PMIOS_TYPE = "M" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("MATERIALS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("MATERIALS")))
                    End If
                    J_GROSS = Round(NumericVal(PMIOS_DS_AMT1), 2)
                    J_TAX = Round(NumericVal(Round((PMIOS_DS_AMT1 / 1.12), 2) * 0.12), 2)
                    J_NET = Round(NumericVal(PMIOS_DS_AMT1) - NumericVal(J_TAX), 2)
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    J_JITEMNO = "'0004'"
                Else
                    J_JITEMNO = "'0003'"
                End If
                J_ACCT_CODE = N2Str2Null(ReturnOutPutTax())
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutPutTax()))
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(Round((PMIOS_NETINVAMT / 1.12), 2) * 0.12), 2)
                J_TAX = 0
                J_GROSS = 0
                J_NET = 0
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
    
                If J_JITEMNO = "'0004'" Then
                   J_JITEMNO = "'0005'"
                Else
                   J_JITEMNO = "'0004'"
                End If
    
                If PMIOS_TYPE = "P" Then
                   J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS")))
                End If
                If PMIOS_TYPE = "A" Then
                   J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("ACCESSORIES"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("ACCESSORIES")))
                End If
                If PMIOS_TYPE = "M" Then
                   J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS", "LUBRICANTS"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS", "LUBRICANTS")))
                End If
                J_GROSS = 0
                J_TAX = 0
                J_NET = 0
                J_DEBIT = Round(NumericVal(PMIOS_NETCOST), 2)
                J_CREDIT = 0
                J_STATUS = "'N'"
                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                
                If J_JITEMNO = "'0004'" Then
                   J_JITEMNO = "'0005'"
                Else
                   J_JITEMNO = "'0006'"
                End If
                
                If PMIOS_TYPE = "P" Then
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                End If
                If PMIOS_TYPE = "A" Then
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA")))
                End If
                If PMIOS_TYPE = "M" Then
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVM"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVM")))
                End If
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(PMIOS_NETCOST), 2)
                J_TAX = 0
                J_GROSS = 0
                J_NET = 0
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
    
                gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                               " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                Grid3.Cell(GridImport, 1).Text = 1
            End If
        End If
NextGrid:
    Next
    Screen.MousePointer = 0: DoEvents
End Sub

Sub ImportSMISSales()
    Screen.MousePointer = 11
    Dim rsSMIS_PURCHAGREE                              As ADODB.Recordset
    Dim SMIS_VI_NO                                     As String
    Dim SMIS_DATERELEASED                              As String
    Dim SMIS_CODE                                      As String
    Dim SMIS_AcctName                                  As String
    Dim SMIS_NETSALESPRICE                             As Double
    Dim SMIS_OTHERS                                    As Double
    Dim SMIS_FOB                                       As Double
    Dim SMIS_TOTALCOST                                       As Double
    
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            If Grid3.Cell(GridImport, 2).Text = "SALES" Then
            Else
               GoTo NextGrid
            End If
            Set rsSMIS_PURCHAGREE = New ADODB.Recordset
            Set rsSMIS_PURCHAGREE = gconDMIS.Execute("Select * from SMIS_PurchAgree Where VI_NO = '" & Grid3.Cell(GridImport, 3).Text & "' AND dateRELEASED = '" & CDate(dtpTranDate) & "' order by VI_NO ASC")
            If Not rsSMIS_PURCHAGREE.EOF And Not rsSMIS_PURCHAGREE.BOF Then
                SMIS_VI_NO = Null2String(rsSMIS_PURCHAGREE!VI_NO)
                SMIS_DATERELEASED = Null2String(rsSMIS_PURCHAGREE!DATERELEASED)
                SMIS_CODE = Null2String(rsSMIS_PURCHAGREE!code)
                SMIS_FOB = N2Str2Zero(rsSMIS_PURCHAGREE!FREIGHT)
                SMIS_OTHERS = N2Str2Zero(rsSMIS_PURCHAGREE!OTHERS)
                SMIS_TOTALCOST = N2Str2Zero(rsSMIS_PURCHAGREE!TOTAL_COST)
                If Null2String(rsSMIS_PURCHAGREE!TERM) = "F" Then
                    SMIS_NETSALESPRICE = (N2Str2Zero(rsSMIS_PURCHAGREE!NETSALESPRICE) + SMIS_FOB) - SMIS_OTHERS
                Else
                    SMIS_NETSALESPRICE = N2Str2Zero(rsSMIS_PURCHAGREE!NETSALESPRICE)
                End If
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If
    
                J_JDATE = N2Date2Null(SMIS_DATERELEASED)
                J_VOUCHERNO = N2Str2Null(GetSJVoucherNo())
                J_JTYPE = "'SJ'"
                J_REMARKS = "NULL"
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(SMIS_CODE)
    
                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0
    
                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = Round(SMIS_NETSALESPRICE, 2)
                J_BALANCE = Round(SMIS_NETSALESPRICE, 2)
                J_AMOUNTPAID = 0
    
                J_STATUS = "'N'"
    
                J_INVOICEDATE = N2Date2Null(SMIS_DATERELEASED)
                J_INVOICENO = N2Str2Null(SMIS_VI_NO)
                J_CHECKNO = "NULL"
                J_DUEDATE = N2Date2Null(SMIS_DATERELEASED)
                J_PAYTYPE = "NULL"
                J_INVOICETYPE = "'VI'"
                J_CHECKDATE = "NULL"
                J_BANKCODE = "NULL"
                J_REFNO = "NULL"
                J_REFDATE = N2Date2Null(SMIS_DATERELEASED)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"
    
                J_JITEMNO = "'0001'"
                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                   J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                Else
                   J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                End If
                J_CREDIT = 0
                J_TAX = 0
                J_GROSS = 0
                J_NET = 0
                J_STATUS = "'N'"
                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                J_JITEMNO = "'0002'"
                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL)))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL))))
                If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                   J_GROSS = NumericVal(SMIS_NETSALESPRICE) + NumericVal(SMIS_OTHERS)
                   J_TAX = Round(NumericVal(J_GROSS) / 9.3333, 2)
                   J_NET = Round(NumericVal(J_GROSS) - NumericVal(J_TAX), 2)
                Else
                   J_GROSS = Round(NumericVal(SMIS_NETSALESPRICE) + NumericVal(SMIS_OTHERS), 2)
                   J_TAX = 0
                   J_NET = Round(NumericVal(SMIS_NETSALESPRICE) + NumericVal(SMIS_OTHERS), 2)
                End If
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(J_NET), 2)
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
    
                J_JITEMNO = "'0003'"
                If SMIS_OTHERS > 0 Then
                    J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL)))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL))))
                    J_GROSS = Round(NumericVal(SMIS_OTHERS), 2)
                    J_TAX = Round(NumericVal(J_GROSS) / 9.3333, 2)
                    J_NET = Round(NumericVal(J_GROSS) - NumericVal(J_TAX), 2)
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    J_JITEMNO = "'0004'"
                End If
                If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                    J_ACCT_CODE = N2Str2Null(ReturnOutPutTax())
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutPutTax()))
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(SMIS_NETSALESPRICE / 9.3333), 2)
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    If J_JITEMNO = "'0004'" Then
                        J_JITEMNO = "'0005'"
                    Else
                        J_JITEMNO = "'0004'"
                    End If
                Else
                    If J_JITEMNO = "'0003'" Then
                        J_JITEMNO = "'0004'"
                    Else
                        J_JITEMNO = "'0003'"
                    End If
                End If
                
                J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL)))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL))))
                J_GROSS = 0
                J_TAX = 0
                J_NET = 0
                J_DEBIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 9.3333), 2)
                J_CREDIT = 0
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                
                If J_JITEMNO = "'0005'" Then
                    J_JITEMNO = "'0006'"
                Else
                    J_JITEMNO = "'0005'"
                End If
    
                J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL)))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL))))
                J_GROSS = 0
                J_TAX = 0
                J_NET = 0
                J_DEBIT = 0
                J_CREDIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 9.3333), 2)
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
    
                gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                               " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                Grid3.Cell(GridImport, 1).Text = 1
            End If
        End If
NextGrid:
    Next
    Screen.MousePointer = 0: DoEvents
End Sub

Function ReturnAR_AccountCode(XXX As String) As String
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'AR' AND TRANTYPE1 = '" & XXX & "'")
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnAR_AccountCode = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function ReturnSales_AccountCode(INVTYPE As String, Optional OTHERTYPE As String) As String
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
If OTHERTYPE = "" Then
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & INVTYPE & "'")
Else
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
End If
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnSales_AccountCode = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function ReturnDiscount_AccountCode(INVTYPE As String, Optional OTHERTYPE As String) As String
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
If OTHERTYPE = "" Then
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'DISCOUNT' AND TRANTYPE2 = '" & INVTYPE & "'")
Else
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'DISCOUNT' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
End If
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnDiscount_AccountCode = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function ReturnCostOfSales(INVTYPE As String, Optional OTHERTYPE As String) As String
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
If OTHERTYPE = "" Then
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & INVTYPE & "'")
Else
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
End If
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnCostOfSales = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function ReturnInventory(INVTYPE As String, Optional OTHERTYPE As String) As String
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
If OTHERTYPE = "" Then
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & INVTYPE & "'")
Else
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
End If
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnInventory = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function ReturnOutPutTax()
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'OUTPUT TAX'")
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnOutPutTax = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function SetVendorName(VVV As Variant)
    Dim rsVENDOR As ADODB.Recordset
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!nameofvendor)
    Else
        SetVendorName = ""
    End If
    Set rsVENDOR = New ADODB.Recordset
End Function

Function CheckIfNonVatSup(SupplierCode As String) As Boolean
    Dim rsSupplierMaster                               As ADODB.Recordset
    Set rsSupplierMaster = New ADODB.Recordset
    rsSupplierMaster.Open "Select supcode,supname,NONVAT from PMIS_vw_Supplier where supcode = '" & SupplierCode & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplierMaster.EOF And Not rsSupplierMaster.BOF Then
        If Null2String(rsSupplierMaster!NONVAT) = "Y" Then
            CheckIfNonVatSup = True
        Else
            CheckIfNonVatSup = False
        End If
    Else
        CheckIfNonVatSup = False
    End If
End Function

Private Sub Option1_Click()
If Option1.Value = True Then
   Grid1.BackColor1 = &HFFFFFF
   Grid1.BackColor2 = &HFFFFFF
   Grid2.BackColor1 = &H8000000F
   Grid2.BackColor2 = &H8000000F
   Grid1.Enabled = True
   Grid2.Enabled = False
Else
   Grid1.BackColor1 = &H8000000F
   Grid1.BackColor2 = &H8000000F
   Grid2.BackColor1 = &HFFFFFF
   Grid2.BackColor2 = &HFFFFFF
   Grid1.Enabled = False
   Grid2.Enabled = True
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
   Grid1.BackColor1 = &H8000000F
   Grid1.BackColor2 = &H8000000F
   Grid2.BackColor1 = &HFFFFFF
   Grid2.BackColor2 = &HFFFFFF
   Grid1.Enabled = False
   Grid2.Enabled = True
Else
   Grid1.BackColor1 = &HFFFFFF
   Grid1.BackColor2 = &HFFFFFF
   Grid2.BackColor1 = &H8000000F
   Grid2.BackColor2 = &H8000000F
   Grid1.Enabled = True
   Grid2.Enabled = False
End If
End Sub

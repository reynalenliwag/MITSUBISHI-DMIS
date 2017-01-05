VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Begin VB.Form FrmGJImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Adjustment Import process"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9285
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
      Left            =   8520
      MouseIcon       =   "FrmGJImport.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "FrmGJImport.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Window"
      Top             =   6750
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MaterialAdjusment"
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
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   6390
      Width           =   4095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Parts Adjusment"
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
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   6120
      Value           =   -1  'True
      Width           =   2205
   End
   Begin VB.CommandButton cmdClearJournals 
      BackColor       =   &H0080FF80&
      Caption         =   "Clear Selected Date"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
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
      Left            =   3600
      MouseIcon       =   "FrmGJImport.frx":04B8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Process Import of SALES"
      Top             =   60
      Width           =   2010
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   4650
      TabIndex        =   4
      Top             =   6330
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   556
      Picture         =   "FrmGJImport.frx":060A
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "FrmGJImport.frx":0626
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
      Left            =   1800
      TabIndex        =   5
      Top             =   60
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
      Format          =   51445761
      CurrentDate     =   38216
   End
   Begin FlexCell.Grid Grid1 
      Height          =   4905
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColor2      =   16777152
      BackColorBkg    =   -2147483633
      BackColorSel    =   16777215
      Cols            =   7
      DefaultFontSize =   8.25
      GridColor       =   16777215
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin FlexCell.Grid Grid2 
      Height          =   4905
      Left            =   4650
      TabIndex        =   7
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColor2      =   16777152
      BackColorBkg    =   -2147483633
      Cols            =   7
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
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
      Left            =   7800
      MouseIcon       =   "FrmGJImport.frx":0642
      MousePointer    =   99  'Custom
      Picture         =   "FrmGJImport.frx":0794
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Process Importing of Cash Receipts "
      Top             =   6750
      Width           =   735
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
      Left            =   60
      TabIndex        =   12
      Top             =   180
      Width           =   1875
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Only Un-Imported Adjusment can be Imported"
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
      Left            =   90
      TabIndex        =   11
      Top             =   7200
      Width           =   7005
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PARTS ADJUSMENT"
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
      Left            =   0
      TabIndex        =   10
      Top             =   540
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MATERIAL ADJUSMENT"
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
      Left            =   4650
      TabIndex        =   9
      Top             =   525
      Width           =   4575
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
      Left            =   4680
      TabIndex        =   8
      Top             =   6060
      Width           =   5835
   End
End
Attribute VB_Name = "FrmGJImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TransactionID                                 As String
Function GetVoucherNo() As String
    Dim rsJournal_HD                              As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'GJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function ReturnAccountName(XXX As String) As String
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset
    SQL = "SELECT Description FROM AMIS_ChartAccount where acctcode=" & XXX & ""
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    If Not RS.EOF And Not RS.BOF Then
        ReturnAccountName = Null2String(RS!Description)
    End If
    Set RS = Nothing
End Function

Function CheckIfORIG(XXX As String) As Boolean
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    SQL = "SELECT GENUINE From PMIS_stockmas where Stockno='" & XXX & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        If Null2String(RS!genuine) = "Y" Then
            CheckIfORIG = True
        Else
            CheckIfORIG = False
        End If
    End If
    Set RS = Nothing
End Function

Function CheckGJifExist(XXX As String, YYY As String) As Boolean
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    SQL = "SELECT lastupdate,refno from AMIS_journal_HD where refno='" & XXX & "' and jdate='" & CDate(YYY) & "' "

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        CheckGJifExist = True
    Else
        CheckGJifExist = False
    End If
    Set RS = Nothing
End Function

Sub InitGrid()
    With Grid1
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "Part No"
        .Cell(0, 3).Text = "Description"
        .Cell(0, 4).Text = "Cost"
        .Cell(0, 5).Text = "(+)"
        .Cell(0, 6).Text = "(-)"

        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 200
        .Column(4).Width = 75
        .Column(5).Width = 30
        .Column(6).Width = 38

        .Column(1).CellType = cellCheckBox
        .Column(4).Alignment = cellRightGeneral
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter

        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True

    End With

    With Grid2
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "Material No"
        .Cell(0, 3).Text = "Description"
        .Cell(0, 4).Text = "Cost"
        .Cell(0, 5).Text = "(+)"
        .Cell(0, 6).Text = "(-)"

        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 200
        .Column(4).Width = 75
        .Column(5).Width = 30
        .Column(6).Width = 38

        .Column(1).CellType = cellCheckBox
        .Column(4).Alignment = cellRightGeneral
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter

        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True

    End With

End Sub

Function ImportGJ() As Boolean
    On Error GoTo Errorcode

    Dim rsJournal_HDDup                           As New ADODB.Recordset
    Dim RsPartsAdjust                             As New ADODB.Recordset
    Dim SQL                                       As String
    Dim GridImports                               As Integer
    Dim PMIS_AdjustDate                           As String
    Dim PMIS_PartsAdjust                          As String
    Dim J_JDATE, J_VOUCHERNO, J_JTYPE             As String
    Dim J_JNO, J_REMARKS, J_VENDORCODE, J_CUSTOMERCODE As String
    Dim J_CUSTOMERNAME                            As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
    Dim J_STATUS, J_JITEMNO, J_CHECKNO            As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE       As String
    Dim J_INVOICETYPE, J_INVOICENO                As String
    Dim J_CHECKDATE, J_BANKCODE                   As String
    Dim J_REFNO, J_REFDATE                        As String
    Dim J_TERMS, J_DEALER, J_ACCT_CODE, J_ACCT_NAME As String
    Dim J_ATC, J_RATE, J_TAXBASE                  As Double
    Dim I                                         As Integer
    Dim J_PAIDSTATUS, J_RECEIVESTATUS
    Dim thePlusAdjust                             As Integer
    Dim theMinusAdjust                            As Integer
    Dim Cost                                      As Double
    J_CUSTOMERNAME = "NULL"
    J_VENDORCODE = "'999999'"
    J_JTYPE = "'GJ'"
    Dim TOTAL_CREDIT                              As Double
    Dim TOTAL_DEBIT                               As Double
    Dim theAdjustedCount                          As String
    Dim theMinusAdjusted                          As String

    TOTAL_CREDIT = 0: TOTAL_DEBIT = 0
    I = 1
    For GridImports = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(I, 1).Text) = 0 Then

            SQL = "SELECT * from PMIS_adjust where partno='" & Grid1.Cell(GridImports, 2).Text & "' and type='P' and status='P' and lastupdate='" & CDate(dtpTranDate) & "'"
            Set RsPartsAdjust = New ADODB.Recordset
            Set RsPartsAdjust = gconDMIS.Execute(SQL)
            If Not RsPartsAdjust.EOF And Not RsPartsAdjust.BOF Then
                PMIS_AdjustDate = Null2String(RsPartsAdjust!lastupdate)
                PMIS_PartsAdjust = Null2String(RsPartsAdjust!PARTNO)
                J_JDATE = N2Date2Null(PMIS_AdjustDate)
                thePlusAdjust = Null2String(RsPartsAdjust!Add)
                theMinusAdjust = Null2String(RsPartsAdjust!minus)
                theAdjustedCount = thePlusAdjust
                theMinusAdjusted = theMinusAdjust

                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If

                J_CUSTOMERCODE = "'999999'"
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_DEBIT = 0: J_CREDIT = 0: J_TAX = 0: J_OUTBALANCE = 0
                J_INVOICEDATE = "NULL": J_BALANCE = 0: J_AMOUNTPAID = 0
                J_DUEDATE = N2Str2Null(PMIS_AdjustDate)
                J_PAYTYPE = "NULL": J_STATUS = "'N'"
                J_TERMS = "NULL": J_DEALER = "NULL"
                J_CHECKDATE = "NULL": J_BANKCODE = "NULL"
                J_INVOICETYPE = "NULL": J_INVOICENO = "NULL"
                J_INVOICEAMT = 0: J_REFNO = N2Str2Null(RsPartsAdjust!PARTNO)
                J_PAIDSTATUS = "'N'": J_RECEIVESTATUS = "'N'"
                J_CHECKNO = "NULL": J_REFDATE = "NULL"
                J_AMOUNTTOPAY = 0


                ' if the Parts inventory is ADD
                If thePlusAdjust <> 0 Then
                    J_REMARKS = N2Str2Null("To Record inventory adjusment with Part No:" + PMIS_PartsAdjust + " (Pcs-" + theAdjustedCount + ")")
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = "'71-63000-30'"
                    J_ACCT_NAME = N2Str2Null(ReturnAccountName("'71-63000-30'"))
                    J_DEBIT = 0
                    J_CREDIT = N2Str2Zero(RsPartsAdjust!Cost) * NumericVal(thePlusAdjust)
                    J_TAX = 0
                    J_ATC = 0
                    J_RATE = 0
                    J_TAXBASE = 0
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "ARNIE", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)

                    J_JITEMNO = "'0002'"
                    If CheckIfORIG(PMIS_PartsAdjust) = True Then
                        J_ACCT_CODE = "'11-05000-00'"
                        J_ACCT_NAME = N2Str2Null(ReturnAccountName("'11-05000-00'"))
                    Else
                        J_ACCT_CODE = "'11-05001-00'"
                        J_ACCT_NAME = N2Str2Null(ReturnAccountName("'11-05001-00'"))
                    End If
                    J_DEBIT = N2Str2Zero(RsPartsAdjust!Cost) * NumericVal(thePlusAdjust)
                    J_CREDIT = 0: J_TAX = 0: J_ATC = 0
                    J_RATE = 0: J_TAXBASE = 0
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "ARNIE", J_JTYPE, "Jtype"))
                    NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)


                End If

                'if the Parts inventory is minus
                If theMinusAdjust <> 0 Then
                    J_REMARKS = N2Str2Null("To Record inventory adjusment with Part No:" + PMIS_PartsAdjust + "(Pcs-" + theAdjustedCount + ")")
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = "'71-63000-30'"
                    J_ACCT_NAME = N2Str2Null(ReturnAccountName("'71-63000-30'"))
                    J_DEBIT = N2Str2Zero(RsPartsAdjust!Cost) * NumericVal(theMinusAdjust)
                    J_CREDIT = 0: J_TAX = 0: J_ATC = 0
                    J_RATE = 0: J_TAXBASE = 0
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"

                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "ARNIE", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)

                    J_JITEMNO = "'0002'"
                    If CheckIfORIG(PMIS_PartsAdjust) = True Then
                        J_ACCT_CODE = "'11-05000-00'"
                        J_ACCT_NAME = N2Str2Null(ReturnAccountName("'11-05000-00'"))
                    Else
                        J_ACCT_CODE = "'11-05001-00'"
                        J_ACCT_NAME = N2Str2Null(ReturnAccountName("'11-05001-00'"))
                    End If
                    J_DEBIT = 0
                    J_CREDIT = N2Str2Zero(RsPartsAdjust!Cost) * NumericVal(theMinusAdjust)
                    J_TAX = 0: J_ATC = 0: J_RATE = 0
                    J_TAXBASE = 0
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"

                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "ARNIE", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)


                End If

                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                " (jdate,voucherno,jtype,vendorcode,customercode,customername,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,USERCODE,LASTUPDATE) " & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & "," & J_CUSTOMERNAME & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                "," & J_JNO & "," & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ",'" & LOGCODE & "','" & LOGDATE & "')"

                gconDMIS.Execute SQL_STATEMENT
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "ARNIE", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)

            End If

        End If
        progCPB.Value = (I / (Grid1.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        Grid1.Cell(I, 1).Text = 1
        I = I + 1
        DoEvents

    Next

    ImportGJ = True
    Exit Function
Errorcode:
    ImportGJ = False
End Function

Function ImportMaterial() As Boolean
    On Error GoTo Errorcode

    Dim rsJournal_HDDup                           As New ADODB.Recordset
    Dim RsPartsAdjust                             As New ADODB.Recordset
    Dim SQL                                       As String
    Dim GridImports                               As Integer
    Dim PMIS_AdjustDate                           As String
    Dim PMIS_PartsAdjust                          As String
    Dim J_JDATE, J_VOUCHERNO, J_JTYPE             As String
    Dim J_JNO, J_REMARKS, J_VENDORCODE, J_CUSTOMERCODE As String
    Dim J_CUSTOMERNAME                            As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
    Dim J_STATUS, J_JITEMNO, J_CHECKNO            As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE       As String
    Dim J_INVOICETYPE, J_INVOICENO                As String
    Dim J_CHECKDATE, J_BANKCODE                   As String
    Dim J_REFNO, J_REFDATE                        As String
    Dim theAdjustedCount                          As String
    Dim theMinusAdjusted                          As String
    Dim J_TERMS, J_DEALER, J_ACCT_CODE, J_ACCT_NAME As String
    Dim J_ATC, J_RATE, J_TAXBASE                  As Double
    Dim I                                         As Integer
    Dim J_PAIDSTATUS, J_RECEIVESTATUS
    Dim thePlusAdjust                             As Integer
    Dim theMinusAdjust                            As Integer
    Dim Cost                                      As Double
    J_CUSTOMERNAME = "NULL"
    J_VENDORCODE = "'999999'"
    J_JTYPE = "'GJ'"
    Dim TOTAL_CREDIT                              As Double
    Dim TOTAL_DEBIT                               As Double


    TOTAL_CREDIT = 0: TOTAL_DEBIT = 0
    I = 1
    For GridImports = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(I, 1).Text) = 0 Then

            SQL = "SELECT * from PMIS_adjust where partno='" & Grid2.Cell(GridImports, 2).Text & "' and type='M' and status='P' and lastupdate='" & CDate(dtpTranDate) & "'"
            Set RsPartsAdjust = New ADODB.Recordset
            Set RsPartsAdjust = gconDMIS.Execute(SQL)
            If Not RsPartsAdjust.EOF And Not RsPartsAdjust.BOF Then
                PMIS_AdjustDate = Null2String(RsPartsAdjust!lastupdate)
                PMIS_PartsAdjust = Null2String(RsPartsAdjust!PARTNO)
                J_JDATE = N2Date2Null(PMIS_AdjustDate)
                thePlusAdjust = Null2String(RsPartsAdjust!Add)
                theMinusAdjust = Null2String(RsPartsAdjust!minus)
                theAdjustedCount = thePlusAdjust
                theMinusAdjusted = theMinusAdjust
                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If
                J_CUSTOMERCODE = "'999999'"
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_DEBIT = 0: J_CREDIT = 0: J_TAX = 0: J_OUTBALANCE = 0
                J_INVOICEDATE = "NULL": J_BALANCE = 0: J_AMOUNTPAID = 0
                J_DUEDATE = N2Str2Null(PMIS_AdjustDate)
                J_PAYTYPE = "NULL": J_STATUS = "'N'"
                J_TERMS = "NULL": J_DEALER = "NULL"
                J_CHECKDATE = "NULL": J_BANKCODE = "NULL"
                J_INVOICETYPE = "NULL": J_INVOICENO = "NULL"
                J_INVOICEAMT = 0: J_REFNO = N2Str2Null(RsPartsAdjust!PARTNO)
                J_PAIDSTATUS = "'N'": J_RECEIVESTATUS = "'N'"
                J_CHECKNO = "NULL": J_REFDATE = "NULL"
                J_AMOUNTTOPAY = 0

                ' if the material adjusted inventory is ADD
                If thePlusAdjust <> 0 Then
                    J_REMARKS = N2Str2Null("To Record inventory adjusment with Material No:" + PMIS_PartsAdjust + "(Pcs-" + theAdjustedCount + ")")
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = "'71-63001-30'"
                    J_ACCT_NAME = N2Str2Null(ReturnAccountName("'71-63001-30'"))
                    J_DEBIT = 0
                    J_CREDIT = N2Str2Zero(RsPartsAdjust!Cost) * NumericVal(thePlusAdjust)
                    J_TAX = 0
                    J_ATC = 0
                    J_RATE = 0
                    J_TAXBASE = 0
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"

                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "ARNIE", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)

                    J_JITEMNO = "'0002'"
                    If PMIS_PartsAdjust = "SVCMAT0068" Then
                        J_ACCT_CODE = "'71-61000-20'"
                        J_ACCT_NAME = N2Str2Null(ReturnAccountName("'71-61000-20'"))

                    Else
                        J_ACCT_CODE = "'11-05007-00'"
                        J_ACCT_NAME = N2Str2Null(ReturnAccountName("'11-05007-00'"))

                    End If

                    J_DEBIT = N2Str2Zero(RsPartsAdjust!Cost) * NumericVal(thePlusAdjust)
                    J_CREDIT = 0: J_TAX = 0: J_ATC = 0
                    J_RATE = 0: J_TAXBASE = 0
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "ARNIE", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)


                End If

                'if the material inventory is minus
                If theMinusAdjust <> 0 Then
                    J_REMARKS = N2Str2Null("To Record inventory adjusment with Material No:" + PMIS_PartsAdjust + "(Pcs-" + theMinusAdjusted + ")")
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = "'71-63001-30'"
                    J_ACCT_NAME = N2Str2Null(ReturnAccountName("'71-63001-30'"))
                    J_DEBIT = N2Str2Zero(RsPartsAdjust!Cost) * NumericVal(thePlusAdjust)
                    J_CREDIT = 0: J_TAX = 0: J_ATC = 0
                    J_RATE = 0: J_TAXBASE = 0
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE)" & _
                                     " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"


                    J_JITEMNO = "'0002'"
                    If PMIS_PartsAdjust = "SVCMAT0068" Then
                        J_ACCT_CODE = "'71-61000-20'"
                        J_ACCT_NAME = N2Str2Null(ReturnAccountName("'71-61000-20'"))

                    Else
                        J_ACCT_CODE = "'11-05007-00'"
                        J_ACCT_NAME = N2Str2Null(ReturnAccountName("'11-05007-00'"))

                    End If

                    J_DEBIT = 0
                    J_CREDIT = N2Str2Zero(RsPartsAdjust!Cost) * NumericVal(thePlusAdjust)
                    J_TAX = 0: J_ATC = 0: J_RATE = 0
                    J_TAXBASE = 0
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "ARNIE", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)

                End If
                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                " (jdate,voucherno,jtype,vendorcode,customercode,customername,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,USERCODE,LASTUPDATE) " & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & "," & J_CUSTOMERNAME & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                "," & J_JNO & "," & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ",'" & LOGCODE & "','" & LOGDATE & "')"
                gconDMIS.Execute SQL_STATEMENT

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "ARNIE", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)


            End If

        End If
        progCPB.Value = (I / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        Grid2.Cell(I, 1).Text = 1
        I = I + 1
        DoEvents

    Next

    ImportMaterial = True
    Exit Function
Errorcode:
    ImportMaterial = False
    'MsgBox "Import Successfully Completed!", vbInformation, "Finish"
End Function

Private Sub cmdCheck_Click()
'If Function_Access(LOGID, "Acess_Process", "IMPORT ADJUSMENT") = False Then Exit Sub
    Dim str_MSG                                   As String


    str_MSG = "Error Appear In During @UTX83912839123" & vbCrLf
    str_MSG = str_MSG & "Imported Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

    gconDMIS.BeginTrans
    If Option1(0).Value = True Then
        If ImportGJ = False Then
            str_MSG = Replace(str_MSG, "@UTX83912839123", "Import Adjustment")
            MsgBox str_MSG, vbCritical, "Importing Error"
            cmdExit.Enabled = True
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If
    Else
        If ImportMaterial = False Then
            str_MSG = Replace(str_MSG, "@UTX83912839123", "Import Adjustment")
            MsgBox str_MSG, vbCritical, "Importing Error"
            cmdExit.Enabled = True
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If

    gconDMIS.CommitTrans
    MsgBox "Import Successfully Completed!", vbInformation, "Finish"
End Sub

Private Sub cmdClearJournals_Click()
    Dim rsCHATCheckControlIfExistRecordInJournalHD As ADODB.Recordset

    Set rsCHATCheckControlIfExistRecordInJournalHD = New ADODB.Recordset
    Set rsCHATCheckControlIfExistRecordInJournalHD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'GJ' and Jdate = '" & CDate(dtpTranDate) & "'")
    If Not rsCHATCheckControlIfExistRecordInJournalHD.EOF And Not rsCHATCheckControlIfExistRecordInJournalHD.BOF Then
        Screen.MousePointer = 0
        If LOGLEVEL = "ADM" Then
            If MsgBox("Clear Unposted Data for this Particular Date?", vbQuestion + vbYesNo, "Purge Data") = vbYes Then
                Screen.MousePointer = 11
                gconDMIS.Execute ("delete from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'GJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                gconDMIS.Execute ("delete from AMIS_Journal_DET Where STATUS <> 'P' AND Jtype = 'GJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                cmdShowTrans.Value = True
                Screen.MousePointer = 0
                MsgBox "Existing Data Successfully deleted.", vbInformation, "Deleted"
                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdShowTrans_Click()
    Dim loveyouarnie                              As Byte
    Screen.MousePointer = 11
    Dim RsPartsAdjust                             As New ADODB.Recordset
    Dim SQL                                       As String
    Dim ARNIE                                     As Integer

    If Option1(0).Value = True Then

        SQL = "SELECT * from PMIS_adjust where status='P' and type='P'and lastupdate='" & CDate(dtpTranDate) & "'"

        Set RsPartsAdjust = New ADODB.Recordset
        Set RsPartsAdjust = gconDMIS.Execute(SQL)

        Grid1.Rows = 1
        If Not RsPartsAdjust.EOF And Not RsPartsAdjust.BOF Then
            RsPartsAdjust.MoveFirst: ARNIE = 0
            Grid1.AutoRedraw = False
            Do While Not RsPartsAdjust.EOF
                ARNIE = ARNIE + 1
                If CheckGJifExist(Null2String(RsPartsAdjust!PARTNO), dtpTranDate) = True Then
                    loveyouarnie = 1
                Else
                    loveyouarnie = 0
                End If
                Grid1.AddItem loveyouarnie & Chr(9) & Null2String(RsPartsAdjust!PARTNO) & Chr(9) & Null2String(RsPartsAdjust!PartDesc) & Chr(9) & ToDoubleNumber(N2Str2Zero(RsPartsAdjust!Cost)) & Chr(9) & N2Str2Zero(RsPartsAdjust!Add) & Chr(9) & N2Str2Zero(RsPartsAdjust!minus)
                RsPartsAdjust.MoveNext
            Loop
            Grid1.AutoRedraw = True
            Grid1.Refresh
            cmdCheck.Enabled = True
        Else
            MsgBox "No Such Record.", vbInformation, "Information"
        End If
        'If ARNIE = 1 Then Grid1.RemoveItem 1
        Grid1.AutoRedraw = True
        Grid1.Refresh
        Screen.MousePointer = 0

    Else

        SQL = "SELECT * from PMIS_adjust where status='P' and type='M'and lastupdate='" & CDate(dtpTranDate) & "'"

        Set RsPartsAdjust = New ADODB.Recordset
        Set RsPartsAdjust = gconDMIS.Execute(SQL)

        Grid2.Rows = 1
        If Not RsPartsAdjust.EOF And Not RsPartsAdjust.BOF Then
            RsPartsAdjust.MoveFirst: ARNIE = 0
            Grid2.AutoRedraw = False
            Do While Not RsPartsAdjust.EOF
                ARNIE = ARNIE + 1
                If CheckGJifExist(Null2String(RsPartsAdjust!PARTNO), dtpTranDate) = True Then
                    loveyouarnie = 1
                Else
                    loveyouarnie = 0
                End If
                Grid2.AddItem loveyouarnie & Chr(9) & Null2String(RsPartsAdjust!PARTNO) & Chr(9) & Null2String(RsPartsAdjust!PartDesc) & Chr(9) & ToDoubleNumber(N2Str2Zero(RsPartsAdjust!Cost)) & Chr(9) & N2Str2Zero(RsPartsAdjust!Add) & Chr(9) & N2Str2Zero(RsPartsAdjust!minus)
                RsPartsAdjust.MoveNext
            Loop
            Grid2.AutoRedraw = True
            Grid2.Refresh
            cmdCheck.Enabled = True
        Else
            MsgBox "No Such Record.", vbInformation, "Information"
        End If
        'If ARNIE = 1 Then Grid1.RemoveItem 1
        Grid2.AutoRedraw = True
        Grid2.Refresh
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitGrid
    dtpTranDate = LOGDATE
End Sub


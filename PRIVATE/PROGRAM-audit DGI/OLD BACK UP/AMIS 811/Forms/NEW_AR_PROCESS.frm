VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form NEW_AR_PROCESS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   3660
   Begin VB.PictureBox picAR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   120
      ScaleHeight     =   1125
      ScaleWidth      =   3405
      TabIndex        =   3
      Top             =   690
      Visible         =   0   'False
      Width           =   3435
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   60
         Picture         =   "NEW_AR_PROCESS.frx":0000
         ScaleHeight     =   405
         ScaleWidth      =   465
         TabIndex        =   5
         Top             =   30
         Width           =   465
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   405
         Left            =   60
         TabIndex        =   4
         Top             =   450
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "VoucehrNo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   60
         TabIndex        =   8
         Top             =   870
         Width           =   2475
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   7
         Top             =   30
         Width           =   3285
      End
      Begin VB.Label labPercent 
         BackStyle       =   0  'Transparent
         Caption         =   "Percent"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   540
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "AR SCHEDULE REPORT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      MaskColor       =   &H0080FFFF&
      TabIndex        =   1
      Top             =   660
      Width           =   3495
   End
   Begin VB.CommandButton cmdAR_PROCESS 
      BackColor       =   &H00FFFF80&
      Caption         =   "Process AR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2070
      MaskColor       =   &H0080FFFF&
      TabIndex        =   0
      Top             =   90
      Width           =   1515
   End
   Begin MSComCtl2.DTPicker dtprocess 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   661
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
End
Attribute VB_Name = "NEW_AR_PROCESS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''Dim xSJ_CustomerCode As String
''Private Sub cmdAR_PROCESS_Click()
''    TRANS_SLS_JOURNAL
''    AR_COMPUTE
''End Sub
''
''Sub TRANS_SLS_JOURNAL()
''    'DESCRIPTION: TRANSFER THE SALES JOURNAL FROM AMIS_JOURNAL_HD TO AMIS_AR_HD
''    Dim rsTRANS_SLS_JOURNAL         As ADODB.Recordset
''    Dim xVOUCHERNO                  As String
''    Dim xJdate                      As String
''    Dim xSTATUS                     As String
''    Dim xJTYPE                      As String
''    Dim xCUSTOMERCODE               As String
''    Dim xINVOICETYPE                As String
''    Dim xINVOICENO                  As String
''    Dim xInvoicedate                As String
''    Dim xINVOICE_AMT                As Double
''    Dim xAMOUNT_TO_PAY              As Double
''    Dim xAMOUNT_PAID                As Double
''    Dim xACCT_CODE                  As String
''    Dim xACCNT_NAME                 As String
''    Dim xDebit                      As Double
''
''    gconDMIS.Execute "TRUNCATE TABLE AMIS_AR_HD"
''
''        Set rsTRANS_SLS_JOURNAL = New ADODB.Recordset
''
''
''         rsTRANS_SLS_JOURNAL.Open "SELECT HD_DET.INVOICENO AS CDJ_NO,HD.VENDORCODE AS VEN_CODE,HD.VoucherNo as HD_VOUCHERNO,HD.jdate AS HD_JDATE,HD.Status AS HD_STATUS, HD.JType AS HD_JTYPE,HD.CustomerCode AS HD_CUST_CODE,HD.InvoiceType AS HD_INV_TYPE, " & _
 ''                                 "HD.InvoiceNo AS HD_INV_NO,HD.InvoiceDate AS HD_INV_DATE,HD.InvoiceAmt AS HD_INV_AMT,HD.AmountToPay AS HD_AMT_TO_PAY,HD.AmountPaid AS HD_AMT_PAID,HD_DET.Acct_Code AS DET_ACCT_CODE, " & _
 ''                                 "HD_DET.Acct_Name AS DET_ACCT_NAME,HD_DET.Debit AS DET_DEBIT FROM AMIS_Journal_HD HD LEFT OUTER JOIN AMIS_Journal_Det HD_DET ON HD.VoucherNo = HD_DET.VoucherNo AND HD.JType = HD_DET.JType " & _
 ''                                 "WHERE LEFT(HD_DET.Acct_Code,5) IN ('11-02','11-03') AND HD.JType IN('SJ','CDJ','COB','CRJ','APJ') AND HD.jdate <= " & N2Str2Null(dtprocess) & " AND HD.status ='P' AND (AR_BALANCE <> 0 OR AR_BALANCE IS NULL) ORDER BY HD.VoucherNo", gconDMIS, adOpenKeyset
''
''        'FOR DEBUGGING PURPOSE ONLY
''        'rsTRANS_SLS_JOURNAL.Open "SELECT HD_DET.INVOICENO AS CDJ_NO,HD.VENDORCODE AS VEN_CODE,HD.VoucherNo as HD_VOUCHERNO,HD.jdate AS HD_JDATE,HD.Status AS HD_STATUS, HD.JType AS HD_JTYPE,HD.CustomerCode AS HD_CUST_CODE,HD.InvoiceType AS HD_INV_TYPE, " & _
 ''                                  "HD.InvoiceNo AS HD_INV_NO,HD.InvoiceDate AS HD_INV_DATE,HD.InvoiceAmt AS HD_INV_AMT,HD.AmountToPay AS HD_AMT_TO_PAY,HD.AmountPaid AS HD_AMT_PAID,HD_DET.Acct_Code AS DET_ACCT_CODE, " & _
 ''                                  "HD_DET.Acct_Name AS DET_ACCT_NAME,HD_DET.Debit AS DET_DEBIT FROM AMIS_Journal_HD HD LEFT OUTER JOIN AMIS_Journal_Det HD_DET ON HD.VoucherNo = HD_DET.VoucherNo AND HD.JType = HD_DET.JType " & _
 ''                                  "WHERE HD_DET.Acct_Code = '11-02017-00' AND HD.JType IN('SJ','CDJ','COB','CRJ','APJ') AND HD.jdate <= " & N2Str2Null(dtprocess) & " AND HD.status ='P' AND (AR_BALANCE <> 0 OR AR_BALANCE IS NULL) ORDER BY HD.VoucherNo", gconDMIS, adOpenKeyset
''
''        If rsTRANS_SLS_JOURNAL.RecordCount = 0 Then Exit Sub
''
''        If Not rsTRANS_SLS_JOURNAL.EOF And Not rsTRANS_SLS_JOURNAL.BOF Then
''            ProgressBar2.Value = 0
''            ProgressBar2.Max = rsTRANS_SLS_JOURNAL.RecordCount
''            Label11.Caption = "Processing SLS Journal....."
''
''            Do While Not rsTRANS_SLS_JOURNAL.EOF
''                'FOR DEBUGGING PURPOSES
''                'If rsTRANS_SLS_JOURNAL!HD_VOUCHERNO = "000199" And RTrim(LTrim(rsTRANS_SLS_JOURNAL!HD_JTYPE)) = "COB" Then Stop
''
''
''                xVOUCHERNO = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_VOUCHERNO)
''                xJdate = N2Date2Null(rsTRANS_SLS_JOURNAL!HD_JDATE)
''                xSTATUS = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_STATUS)
''                xJTYPE = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_JTYPE)
''
''                If Null2String(rsTRANS_SLS_JOURNAL!HD_JTYPE) = "CDJ" Or Null2String(rsTRANS_SLS_JOURNAL!HD_JTYPE) = "APJ" Then
''                    xCUSTOMERCODE = N2Str2Null(rsTRANS_SLS_JOURNAL!VEN_CODE)
''                    xINVOICENO = N2Str2Null(rsTRANS_SLS_JOURNAL!CDJ_NO)
''                Else
''                    xCUSTOMERCODE = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_CUST_CODE)
''                    xINVOICENO = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_INV_NO)
''                End If
''
''                xINVOICETYPE = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_INV_TYPE)
''                xInvoicedate = N2Date2Null(rsTRANS_SLS_JOURNAL!HD_INV_DATE)
''                xINVOICE_AMT = NumericVal(rsTRANS_SLS_JOURNAL!HD_INV_AMT)
''
''                If Null2String(rsTRANS_SLS_JOURNAL!HD_JTYPE) = "COB" Then
''                    xAMOUNT_TO_PAY = NumericVal(rsTRANS_SLS_JOURNAL!HD_INV_AMT)
''                Else
''                    xAMOUNT_TO_PAY = NumericVal(rsTRANS_SLS_JOURNAL!HD_AMT_TO_PAY)
''                End If
''
''                xAMOUNT_PAID = NumericVal(rsTRANS_SLS_JOURNAL!HD_AMT_PAID)
''                xACCT_CODE = N2Str2Null(rsTRANS_SLS_JOURNAL!DET_ACCT_CODE)
''                xACCNT_NAME = N2Str2Null(rsTRANS_SLS_JOURNAL!DET_ACCT_NAME)
''                xDebit = NumericVal(rsTRANS_SLS_JOURNAL!DET_DEBIT)
''
''                gconDMIS.Execute "Insert into AMIS_AR_HD(VoucherNo,Jdate,Status,JType,SJ_CustomerCode,InvoiceType,InvoiceNo,InvoiceDate,InvoiceAmnt,AmountToPay,Acct_code,Debit)" & _
 ''                                 "VALUES(" & xVOUCHERNO & "," & xJdate & "," & xSTATUS & "," & xJTYPE & "," & xCUSTOMERCODE & "," & xINVOICETYPE & "," & xINVOICENO & "," & xInvoicedate & "," & xINVOICE_AMT & "," & xAMOUNT_TO_PAY & "," & xACCT_CODE & "," & xDebit & ")"
''
''                Label12.Caption = Null2String(rsTRANS_SLS_JOURNAL!HD_JTYPE) & "-" & Null2String(rsTRANS_SLS_JOURNAL!HD_VOUCHERNO)
''                ProgressBar2.Value = ProgressBar2.Value + 1
''                labPercent.Caption = Round((ProgressBar2.Value / ProgressBar2.Max) * 100, 0) & "%"
''                DoEvents
''                rsTRANS_SLS_JOURNAL.MoveNext
''            Loop
''        End If
''    Set rsTRANS_SLS_JOURNAL = Nothing
''End Sub
''
''Sub AR_COMPUTE()
''    'DESCRIPTION: COMPUTE THE AR FROM SJ BEING DEDUCTED BY THE CRJ PAYMENT AND COMPUTE IF THERE IS AN ADJUSTMENT
''    Dim rsAR_COMPUTE        As ADODB.Recordset
''    Dim xSJ_VOUCHERNO       As String
''    Dim xCRJ_VOUCHERNO      As String
''    Dim xINVOICETYPE        As String
''    Dim xINVOICENO          As String
''    Dim xAMOUNT_TOPAY       As Double
''    Dim xAMOUNT_PAID        As Double
''    Dim xBALANCE            As Double
''    Dim xACCT_CODE          As String
''    Dim xSYSTEM_REMARKS     As String
''    Dim xInvoicedate        As String
''    Dim xLASTUPDATED        As String
''    Dim xACCT_NAME          As String
''    Dim xCHECKER            As Integer
''
''        xCHECKER = 0
''
''        gconDMIS.Execute "TRUNCATE TABLE AMIS_AR"
''        gconDMIS.Execute "TRUNCATE TABLE AMIS_DETAIL"
''
''        Set rsAR_COMPUTE = New ADODB.Recordset
''              rsAR_COMPUTE.Open "SELECT DISTINCT VOUCHERNO,JTYPE,INVOICETYPE,INVOICENO,SJ_CUSTOMERCODE,AMOUNTTOPAY,ACCT_CODE,INVOICEDATE,DEBIT,JDATE FROM AMIS_AR_HD ORDER BY VOUCHERNO ASC", gconDMIS, adOpenKeyset
''
''             'FOR DEBUGGING PURPOSES ONLY
''             'rsAR_COMPUTE.Open "SELECT DISTINCT VOUCHERNO,JTYPE,INVOICETYPE,INVOICENO,SJ_CUSTOMERCODE,AMOUNTTOPAY,ACCT_CODE,INVOICEDATE,JDATE FROM AMIS_AR_HD WHERE VOUCHERNO = '000146' AND JTYPE = 'SJ' ORDER BY VOUCHERNO ASC", gconDMIS, adOpenKeyset
''
''            If rsAR_COMPUTE.RecordCount = 0 Then Exit Sub
''
''            If Not rsAR_COMPUTE.EOF And Not rsAR_COMPUTE.BOF Then
''                ProgressBar2.Value = 0
''                ProgressBar2.Max = rsAR_COMPUTE.RecordCount
''                Label11.Caption = "Processing AR... Please Wait.."
''                Do While Not rsAR_COMPUTE.EOF
''                    xSJ_VOUCHERNO = Null2String(LTrim(RTrim(rsAR_COMPUTE!jtype))) & "-" & Null2String(LTrim(RTrim(rsAR_COMPUTE!VOUCHERNO)))
''                    'If rsAR_COMPUTE!VOUCHERNO = "002171" And rsAR_COMPUTE!jtype = "CRJ" Then Stop
''
''
''                    Dim rsJNO As ADODB.Recordset
''                    Set rsJNO = New ADODB.Recordset
''                    rsJNO.Open "Select VOUCHERNO,JTYPE,JNO,DEBIT,CREDIT FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & rsAR_COMPUTE!VOUCHERNO & "' AND JTYPE = '" & rsAR_COMPUTE!jtype & "' AND (CUSTOMERCODE = '" & rsAR_COMPUTE!SJ_CustomerCode & "' OR VENDORCODE = '" & rsAR_COMPUTE!SJ_CustomerCode & "') AND " & _
 ''                               "JNO IN(SELECT JNO FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & rsAR_COMPUTE!VOUCHERNO & "' AND JTYPE = '" & rsAR_COMPUTE!jtype & "')", gconDMIS, adOpenKeyset
''
''                    If Not rsJNO.EOF And Not rsJNO.BOF Then
''
''                            If Null2String(rsAR_COMPUTE!jtype) = "CDJ" Then
''                                'xAMOUNT_TOPAY = GET_AR_CDJ_AMOUNT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!jtype), xSJ_CustomerCode)
''                                xAMOUNT_TOPAY = GET_AR_CDJ_AMOUNT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!jtype), xSJ_CustomerCode, Null2String(rsAR_COMPUTE!Acct_Code))
''                                xAMOUNT_PAID = GET_AR_CDJ_PAYENT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!jtype), Null2String(rsAR_COMPUTE!SJ_CustomerCode), Null2String(rsAR_COMPUTE!INVOICENO))
''                                xACCT_NAME = GET_VEN_NAME(LTrim(RTrim(xSJ_CustomerCode)))
''                            ElseIf Null2String(rsAR_COMPUTE!jtype) = "SJ" Then
''                                xAMOUNT_TOPAY = GET_SJ_DEBIT_AMOUNT(Null2String(rsAR_COMPUTE!VOUCHERNO), xINVOICENO, xINVOICETYPE, xSJ_CustomerCode, Null2String(rsAR_COMPUTE!Acct_Code))
''                                xAMOUNT_PAID = COMP_AMT_PAID(xINVOICENO, xINVOICETYPE, xSJ_CustomerCode, Null2String(rsAR_COMPUTE!Acct_Code), Null2String(rsAR_COMPUTE!VOUCHERNO), Null2Date(rsAR_COMPUTE!JDate))
''                                xACCT_NAME = GET_CUST_NAME(LTrim(RTrim(xSJ_CustomerCode)))
''                            ElseIf Null2String(rsAR_COMPUTE!jtype) = "COB" Then
''                                xAMOUNT_TOPAY = GET_COB_AMOUNT(rsAR_COMPUTE!VOUCHERNO, rsAR_COMPUTE!jtype, rsAR_COMPUTE!SJ_CustomerCode)
''                                xAMOUNT_PAID = COMP_AMT_PAID(xINVOICENO, xINVOICETYPE, xSJ_CustomerCode, Null2String(rsAR_COMPUTE!Acct_Code), Null2String(rsAR_COMPUTE!VOUCHERNO), Null2Date(rsAR_COMPUTE!JDate))
''                                xACCT_NAME = GET_CUST_NAME(LTrim(RTrim(xSJ_CustomerCode)))
''                            ElseIf Null2String(rsAR_COMPUTE!jtype) = "APJ" Then
''                                xAMOUNT_TOPAY = GET_APJ_AR(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!jtype), Null2String(rsAR_COMPUTE!Acct_Code))
''                                'xAMOUNT_PAID = COMP_AMT_PAID(xINVOICENO, xINVOICETYPE, xSJ_CustomerCode, Null2String(rsAR_COMPUTE!ACCT_CODE), Null2String(rsAR_COMPUTE!VOUCHERNO))
''                                xACCT_NAME = GET_VEN_NAME(LTrim(RTrim(xSJ_CustomerCode)))
''                            End If
''
''
''                            xBALANCE = Round(NumericVal(xAMOUNT_TOPAY) - NumericVal(xAMOUNT_PAID), 2)
''                            xACCT_CODE = Null2String(rsAR_COMPUTE!Acct_Code)
''                            xInvoicedate = Null2Date(rsAR_COMPUTE!invoicedate)
''                            xLASTUPDATED = LOGDATE
''
''                            'CHECK IF VOUCHERNO AND JTYPE IS ALREADY EXISTING IN AMIS_AR
''                            Dim rsVOUCHERNO_IN_AR As ADODB.Recordset
''                            Set rsVOUCHERNO_IN_AR = New ADODB.Recordset
''                                rsVOUCHERNO_IN_AR.Open "Select * from Amis_Ar where SJVOUCHERNO = '" & xSJ_VOUCHERNO & "'", gconDMIS, adOpenKeyset
''                                If Not rsVOUCHERNO_IN_AR.EOF And Not rsVOUCHERNO_IN_AR.BOF Then
''                                Else
''                                    If NumericVal(xAMOUNT_TOPAY) = 0 And NumericVal(xAMOUNT_PAID) = 0 And NumericVal(xBALANCE) = 0 Then
''                                        'DONT INSERT
''                                    Else
''                                        gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
 ''                                                         "VALUES('" & xSJ_VOUCHERNO & "','" & xINVOICETYPE & "','" & xINVOICENO & "','" & xSJ_CustomerCode & "','" & xACCT_NAME & "','" & xAMOUNT_TOPAY & "','" & xAMOUNT_PAID & "','" & xBALANCE & "','" & xACCT_CODE & "'," & N2Str2Null("") & ",'" & xInvoicedate & "','" & xLASTUPDATED & "')"
''
''                                        'THIS IS FOR AMIS_AR DETAIL
''
''                                        If NumericVal(xCHECKER) = 1 Then
''                                            If CHECK_DUPLICATE(N2Str2Null(INV_FROMAT), N2Str2Null(INV_TEMP), N2Str2Null(LTrim(RTrim(rsAR_COMPUTE!VOUCHERNO)))) = False Then
''                                                gconDMIS.Execute "INSERT INTO AMIS_DETAIL (INVOICENO,INVOICETYPE,INVOICEAMOUNT,CUSTOMERCODE,ACCT_CODE,JDATE,REMARKS,VOUCHERNO) " & _
 ''                                                                 "VALUES(" & N2Str2Null(INV_FROMAT) & ", " & N2Str2Null(INV_TEMP) & ", " & NumericVal(xAMOUNT_PAID) & ", " & N2Str2Null(xSJ_CustomerCode) & "," & N2Str2Null(rsAR_COMPUTE!Acct_Code) & "," & N2Date2Null(rsAR_COMPUTE!JDate) & ",'1','" & LTrim(RTrim(rsAR_COMPUTE!VOUCHERNO)) & "')"
''                                            End If
''                                        End If
''                                    End If
''                                End If
''                            Set rsVOUCHERNO_IN_AR = Nothing
'''
'''                            If NumericVal(xBALANCE) = 0 Then
'''                                gconDMIS.Execute "Update Amis_Journal_Hd set AR_DATEGEN = '" & xLASTUPDATED & "', AR_BALANCE = '" & xBALANCE & "' where Voucherno = '" & Null2String(rsAR_COMPUTE!VOUCHERNO) & "' and JTYPE = '" & Null2String(rsAR_COMPUTE!jtype) & "'"
'''                            Else
'''                                gconDMIS.Execute "Update Amis_Journal_Hd set AR_DATEGEN = '" & xLASTUPDATED & "', AR_BALANCE = '" & xBALANCE & "' where Voucherno = '" & Null2String(rsAR_COMPUTE!VOUCHERNO) & "' and JTYPE = '" & Null2String(rsAR_COMPUTE!jtype) & "'"
'''                            End If
''                    Else
''                        'DO NOTHING
''                    End If
''                    Set rsJNO = Nothing
''
''
''                    xAMOUNT_TOPAY = 0
''                    xAMOUNT_PAID = 0
''                    xBALANCE = 0
''                    xCHECKER = 0
''
''
''                    ProgressBar2.Value = ProgressBar2.Value + 1
''                    labPercent.Caption = Round((ProgressBar2.Value / ProgressBar2.Max) * 100, 0) & "%"
''                    Label12.Caption = xSJ_VOUCHERNO
''                    DoEvents
''                    rsAR_COMPUTE.MoveNext
''                Loop
''            End If
''
''    Set rsAR_COMPUTE = Nothing
''    'Set rsCHECK_IN_AR = Nothing
''End Sub
''
''Function COMP_AMT_PAID(xINVOICENO As String, xINVOICETYPE As String, xCUSTOMERCODE As String, xACCT_CODE As String, xSJ_VOUCHERNO As String) As Double
''    Dim rsCOMP_AMT_PAID As ADODB.Recordset
''    Dim SUM_CRJ As Double
''        SUM_CRJ = 0
''        SUM_ADJ = 0
''
''        'THIS IS TO SET ANOTHER INVOICETYPE DUE SERVICE INVOICE AND VEHICLE INVOICE IS DIFFERENT FROM VI AND SI
''        If RTrim(LTrim(xINVOICETYPE)) = "VI" Then
''            INVOICETYPE_TEMP = "VEHICLE INVOICE"
''        ElseIf RTrim(LTrim(xINVOICETYPE)) = "SI" Then
''            INVOICETYPE_TEMP = "SERVICE INVOICE"
''        ElseIf RTrim(LTrim(xINVOICETYPE)) = "SI" Then
''            INVOICETYPE_TEMP = "SI"
''        ElseIf RTrim(LTrim(xINVOICETYPE)) = "VI" Then
''            INVOICETYPE_TEMP = "VI"
''        End If
''
''        Set rsCOMP_AMT_PAID = New ADODB.Recordset
''            If IsNumeric(xINVOICENO) = True Then
''                rsCOMP_AMT_PAID.Open "SELECT * FROM AMIS_CRJ_DETAIL WHERE (INVOICETYPE = '" & xINVOICETYPE & "' OR INVOICETYPE = '" & INVOICETYPE_TEMP & "') and (INVOICENO = '" & Abs(xINVOICENO) & "' or INVOICENO = '" & xINVOICENO & "') AND SJ_VOUCHERNO = '" & xSJ_VOUCHERNO & "' AND CUSTOMERCODE = '" & xCUSTOMERCODE & "'", gconDMIS, adOpenKeyset
''            Else
''                rsCOMP_AMT_PAID.Open "SELECT * FROM AMIS_CRJ_DETAIL WHERE (INVOICETYPE = '" & xINVOICETYPE & "' OR INVOICETYPE = '" & INVOICETYPE_TEMP & "') and  INVOICENO = '" & xINVOICENO & "' AND SJ_VOUCHERNO = '" & xSJ_VOUCHERNO & "' AND CUSTOMERCODE = '" & xCUSTOMERCODE & "'", gconDMIS, adOpenKeyset
''            End If
''
''            If Not rsCOMP_AMT_PAID.EOF And Not rsCOMP_AMT_PAID.BOF Then
''                Do While Not rsCOMP_AMT_PAID.EOF
''                    SUM_CRJ = Round((NumericVal(SUM_CRJ) + NumericVal(rsCOMP_AMT_PAID!INVOICEAMOUNT)), 2)
''                    rsCOMP_AMT_PAID.MoveNext
''                Loop
''            End If
''
''        'GET THE ADJUSTED AMOUNT
''        Set rsGET_ADJ = New ADODB.Recordset
''        rsGET_ADJ.Open "Select Debit,Credit from Amis_journal_det where InvoiceNo = '" & xINVOICENO & "' and invoicetype = '" & xINVOICETYPE & "' and right(Entity,6) = '" & xSJ_CustomerCode & "' AND Left(Acct_Code,5) IN('11-02','11-03') AND STATUS = 'P' and JDATE< = '" & dtprocess & "'", gconDMIS, adOpenKeyset
''        If Not rsGET_ADJ.EOF And Not rsGET_ADJ.BOF Then
''            Do While Not rsGET_ADJ.EOF
''                If NumericVal(rsGET_ADJ!CREDIT) <> 0 Then
''                    SUM_ADJ = Round((SUM_ADJ + NumericVal(rsGET_ADJ!CREDIT)), 2)
''                End If
''                rsGET_ADJ.MoveNext
''            Loop
''        End If
''
''        COMP_AMT_PAID = Round((NumericVal(SUM_CRJ) + NumericVal(SUM_ADJ)), 2)
''    Set rsCOMP_AMT_PAID = Nothing
''    Set rsGET_ADJ = Nothing
''End Function

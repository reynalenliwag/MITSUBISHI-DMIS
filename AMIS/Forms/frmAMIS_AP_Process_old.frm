VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMIS_AP_Process_old 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AP Aging Report"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3525
   Icon            =   "frmAMIS_AP_Process_old.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1935
   ScaleWidth      =   3525
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   60
      ScaleHeight     =   1815
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   30
      Width           =   3405
      Begin VB.CommandButton Command2 
         Caption         =   "Print AP Schedule"
         Height          =   465
         Left            =   150
         TabIndex        =   10
         Top             =   1140
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print AP Aging Report"
         Height          =   465
         Left            =   150
         TabIndex        =   9
         Top             =   690
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker dtAsOF 
         Height          =   345
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
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
         Format          =   51773441
         CurrentDate     =   40031
      End
      Begin VB.CommandButton cmdProcessAP 
         Caption         =   "Process AP"
         Height          =   345
         Left            =   1800
         TabIndex        =   2
         Top             =   90
         Width           =   1485
      End
   End
   Begin Crystal.CrystalReport rptAP_Aging 
      Left            =   1410
      Top             =   2370
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   330
      Top             =   2070
   End
   Begin VB.PictureBox picAP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   3255
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   3285
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   60
         Picture         =   "frmAMIS_AP_Process_old.frx":08CA
         ScaleHeight     =   405
         ScaleWidth      =   465
         TabIndex        =   5
         Top             =   60
         Width           =   465
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   405
         Left            =   60
         TabIndex        =   4
         Top             =   510
         Width           =   3105
         _ExtentX        =   5477
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
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   930
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
         Top             =   60
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
         Top             =   270
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmAMIS_AP_Process_old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xREMARKS                                      As String
Dim xPAYMENT_TYPE                                 As String
Dim zJTYPE                                        As String
Dim xJType                                        As String

Private Sub cmdProcessAP_Click()
'DESCRIPTION: ACCOUNTS PAYABLE AGING PROCESS
    picAP.Visible = True
    picAP.ZOrder 0
    cmdProcessAP.Enabled = False
    TRANSFER_AP_JOURNAL
    DIRECT_DSBRSMENT
    DISBURSEMENT1
    DISBURSEMENT2
    DISBURSEMENT3
    DISBURSEMENT4
    DISBURSEMENT5
    DISBURSEMENT6
    DISBURSEMENT7
    DISBURSEMENT8
    DISBURSEMENT9
    DISBURSEMENT10
    COMPUTE_AP
    picAP.Visible = False
    picAP.ZOrder 1

    MsgBox "Processing AP Completed"
    cmdProcessAP.Enabled = True
End Sub

Sub COMPUTE_AP()
'DESCRIPTION: ITS GETS THE AMOUNT TO PAY FROM AMIS_AP_HD AND COMPUTE THE PAYMENT AMOUNT
    Dim rsCOMPUTE_AP                              As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xCDJ_VOUCHERNO                            As String
    Dim xVENDORCODE                               As String
    Dim xVENDORNAME                               As String
    Dim xPAYMENT_TYPE                             As String
    Dim xDUEDATE                                  As String
    Dim xINVOICENO                                As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xBALANCE                                  As Double
    Dim xACCT_CODE                                As String
    Dim xLASTUPDATE                               As String
    xAMOUNTPAID = 0
    xAMOUNT2PAY = 0
    Label11.Caption = "Computing AP... Please wait.."

    gconDMIS.Execute "TRUNCATE TABLE AMIS_AP"
    'gconDMIS.Execute "TRUNCATE TABLE AMIS_AP_DETAIL"

    Set rsCOMPUTE_AP = New ADODB.Recordset
    'Description: For Debugging
    'rsCOMPUTE_AP.Open "SELECT DISTINCT VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICENO,INVOICETYPE,INVOICEDATE,INVOICEAMT,AMOUNT2PAY,ACCT_CODE,DUEDATE FROM AMIS_AP_HD WHERE VOUCHERNO='001079' AND JTYPE='APJ'", gconDMIS, adOpenKeyset
    '==================

    rsCOMPUTE_AP.Open "SELECT DISTINCT VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICENO,INVOICETYPE,INVOICEDATE,INVOICEAMT,AMOUNT2PAY,ACCT_CODE,DUEDATE FROM AMIS_AP_HD", gconDMIS, adOpenKeyset
    If rsCOMPUTE_AP.RecordCount = 0 Then Exit Sub
    ProgressBar2.Value = 0
    ProgressBar2.Max = rsCOMPUTE_AP.RecordCount

    If Not rsCOMPUTE_AP.EOF And Not rsCOMPUTE_AP.BOF Then
        Do While Not rsCOMPUTE_AP.EOF
            'If Null2String(rsCOMPUTE_AP!VOUCHERNO) = "002362" And Null2String(rsCOMPUTE_AP!jtype) = "CDJ" Then Stop
            xJType = Null2String(RTrim(LTrim(rsCOMPUTE_AP!jtype)))
            xVOUCHERNO = N2Str2Null(Null2String(RTrim(LTrim(rsCOMPUTE_AP!jtype))) & "-" & Null2String(rsCOMPUTE_AP!VOUCHERNO))
            xCDJ_VOUCHERNO = N2Str2Null(GET_CDJNO(Null2String(rsCOMPUTE_AP!VOUCHERNO), Null2String(rsCOMPUTE_AP!jtype)))
            xVENDORCODE = N2Str2Null(rsCOMPUTE_AP!VENDOR_CODE)
            xVENDORNAME = N2Str2Null(GET_VEN_NAME(N2Str2Null(rsCOMPUTE_AP!VENDOR_CODE)))
            xDUEDATE = N2Date2Null(rsCOMPUTE_AP!duedate)
            xINVOICENO = N2Str2Null(rsCOMPUTE_AP!INVOICENO)
            xINVOICETYPE = N2Str2Null(rsCOMPUTE_AP!InvoiceType)
            xInvoicedate = N2Date2Null(rsCOMPUTE_AP!invoicedate)
            xAMOUNT2PAY = Round(NumericVal(rsCOMPUTE_AP!AMOUNT2PAY) + NumericVal(AMOUNT2PAY(Null2String(rsCOMPUTE_AP!VOUCHERNO), Null2String(rsCOMPUTE_AP!jtype), Null2String(rsCOMPUTE_AP!VENDOR_CODE), Null2String(rsCOMPUTE_AP!Acct_code), NumericVal(rsCOMPUTE_AP!AMOUNT2PAY))) + NumericVal(Amount2Pay2(Null2String(rsCOMPUTE_AP!VOUCHERNO), Null2String(rsCOMPUTE_AP!jtype), Null2String(rsCOMPUTE_AP!VENDOR_CODE), Null2String(rsCOMPUTE_AP!Acct_code))), 2)
            'xAMOUNT2PAY = Round(NumericVal(Amount2Pay(Null2String(rsCOMPUTE_AP!VOUCHERNO), Null2String(rsCOMPUTE_AP!jtype), Null2String(rsCOMPUTE_AP!VENDOR_CODE), Null2String(rsCOMPUTE_AP!Acct_Code))), 2)

            If Null2String(RTrim(LTrim(rsCOMPUTE_AP!jtype))) = "GJ" Or Null2String(RTrim(LTrim(rsCOMPUTE_AP!jtype))) = "APJ" Or Null2String(RTrim(LTrim(rsCOMPUTE_AP!jtype))) = "CRJ" Or Null2String(RTrim(LTrim(rsCOMPUTE_AP!jtype))) = "VPJ" Then
                xAMOUNTPAID = Round(NumericVal(COMP_AP_AMT_PAID(Null2String(rsCOMPUTE_AP!VOUCHERNO), Null2String(rsCOMPUTE_AP!jtype), Null2String(rsCOMPUTE_AP!VENDOR_CODE), Null2String(rsCOMPUTE_AP!Acct_code))) + COMP_AP_ADJ_DEBIT(xVENDORCODE, xINVOICENO, xINVOICETYPE, Null2String(LTrim(RTrim(rsCOMPUTE_AP!jtype)))), 2)
            ElseIf Null2String(RTrim(LTrim(rsCOMPUTE_AP!jtype))) = "SJ" And rsCOMPUTE_AP!AMOUNT2PAY = 0 Then
                xAMOUNTPAID = Round(NumericVal(COMP_AP_AMT_PAID(Null2String(rsCOMPUTE_AP!VOUCHERNO), Null2String(rsCOMPUTE_AP!jtype), Null2String(rsCOMPUTE_AP!VENDOR_CODE), Null2String(rsCOMPUTE_AP!Acct_code))) + COMP_AP_ADJ_DEBIT(xVENDORCODE, xINVOICENO, xINVOICETYPE, Null2String(LTrim(RTrim(rsCOMPUTE_AP!jtype)))), 2)
            Else
                xAMOUNTPAID = Round(NumericVal(COMP_AP_AMT_PAID2(Null2String(rsCOMPUTE_AP!VOUCHERNO), Null2String(rsCOMPUTE_AP!jtype), Null2String(rsCOMPUTE_AP!VENDOR_CODE), Null2String(rsCOMPUTE_AP!Acct_code))) + COMP_AP_ADJ_DEBIT(xVENDORCODE, xINVOICENO, xINVOICETYPE, Null2String(LTrim(RTrim(rsCOMPUTE_AP!jtype)))), 2)
            End If
            If Null2String(RTrim(LTrim(rsCOMPUTE_AP!jtype))) = "VDJ" Or Null2String(RTrim(LTrim(rsCOMPUTE_AP!jtype))) = "VCJ" Then
                xVENDORCODE = N2Str2Null("XXXXX")
                xVENDORNAME = N2Str2Null("XXXXX")
            End If
            xBALANCE = Round(NumericVal(xAMOUNT2PAY) - NumericVal(xAMOUNTPAID), 2)
            xACCT_CODE = N2Str2Null(rsCOMPUTE_AP!Acct_code)
            xLASTUPDATE = N2Date2Null(LOGDATE)

            gconDMIS.Execute "Insert Into AMIS_AP (VOUCHERNO,CDJ_VOUCHERNO,VENDOR_CODE,VENDOR_NAME,PAYMENT_TYPE,DUEDATE,INVOICEDATE,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,SYSTEMREMARK,LASTUPDATED)" & _
                             "VALUES(" & xVOUCHERNO & "," & xCDJ_VOUCHERNO & "," & xVENDORCODE & "," & xVENDORNAME & ", " & N2Str2Null(xPAYMENT_TYPE) & ", " & xDUEDATE & "," & xInvoicedate & "," & xAMOUNT2PAY & "," & xAMOUNTPAID & "," & xBALANCE & "," & xACCT_CODE & "," & N2Str2Null(xREMARKS) & "," & xLASTUPDATE & ")"

            'gconDMIS.Execute "Insert into AMIS_AP_DETAIL (VoucherNo , JDate, VendorCode, INVOICENO, InvoiceType, InvoiceAmount, Acct_Code, Remarks)" & _
             "Values(" & xVOUCHERNO & "," & xInvoicedate & "," & xVENDORCODE & "," & xINVOICENO & "," & xINVOICETYPE & "," & xAMOUNT2PAY & "," & xACCT_CODE & "," & N2Str2Null(xREMARKS) & ")"

            '                If xBALANCE <> 0 Then
            '                    gconDMIS.Execute "Update Amis_journal_hd set AR_BALANCE = " & xBALANCE & ", AR_DATEGEN = " & xLASTUPDATE & " WHERE VOUCHERNO = '" & Null2String(rsCOMPUTE_AP!VOUCHERNO) & "' AND JTYPE = '" & rsCOMPUTE_AP!jtype & "' "
            '                Else
            '                    gconDMIS.Execute "Update Amis_journal_hd set AR_BALANCE = " & xBALANCE & ", AR_DATEGEN = " & xLASTUPDATE & " WHERE VOUCHERNO = '" & Null2String(rsCOMPUTE_AP!VOUCHERNO) & "' AND JTYPE = '" & rsCOMPUTE_AP!jtype & "' "
            '                End If

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsCOMPUTE_AP!VOUCHERNO)
            DoEvents
            rsCOMPUTE_AP.MoveNext
        Loop
    End If
    Set rsCOMPUTE_AP = Nothing
End Sub

Function AMOUNT2PAY(xAP_VOUCHERnO As String, xJType As String, xVEN_CODE As String, xACCT_CODE As String, xAMOUNT2PAY) As Double
    Dim rsAmount2Pay                              As ADODB.Recordset
    Dim sum_Amount2Pay                            As Double
    Dim xSJCount                                  As Long
    Set rsAmount2Pay = New ADODB.Recordset
    rsAmount2Pay.Open "SELECT COUNT(VOUCHERNO) AS SJCount,VOUCHERNO,JTYPE,ACCT_CODE,AMOUNT2PAY,VENDOR_CODE FROM AMIS_AP_HD where VoucherNo = '" & xAP_VOUCHERnO & "' and JType = 'SJ' and Vendor_Code ='" & xVEN_CODE & "' and Left(Acct_Code,5) = '21-01' and Amount2Pay ='" & xAMOUNT2PAY & "' GROUP BY VOUCHERNO,JTYPE,ACCT_CODE,AMOUNT2PAY,VENDOR_CODE HAVING COUNT(VOUCHERNO) > 1 ", gconDMIS, adOpenKeyset
    If rsAmount2Pay.RecordCount = 0 Then Exit Function
    If Not rsAmount2Pay.EOF And Not rsAmount2Pay.BOF Then
        Do While Not rsAmount2Pay.EOF
            xSJCount = rsAmount2Pay!SJCount - 1
            Do While Not xSJCount = 0
                sum_Amount2Pay = sum_Amount2Pay + rsAmount2Pay!AMOUNT2PAY
                xSJCount = xSJCount - 1
            Loop
            rsAmount2Pay.MoveNext
        Loop
        AMOUNT2PAY = NumericVal(sum_Amount2Pay)
    End If
End Function

Function COMP_AP_AMT_PAID(xAP_VOUCHERnO As String, xJType As String, xVEN_CODE As String, xACCT_CODE As String) As Double
'DESCPRIPTION: COMPUTE THE AMOUNT PAID  AND VALIDATE THE VENDOR CODE IF CODE IS VALID SUM UP THE PAYMENT
    Dim rsCheckJTYPE                              As ADODB.Recordset
    Dim rsCOMP_AP_AMT_PAID                        As ADODB.Recordset
    Dim rsCOMP_AP_AMT_PAID2                       As ADODB.Recordset
    Dim rsCOMP_AP_AMT_PAID3                       As ADODB.Recordset
    Dim rsCOMP_AP_AMT_PAID4                       As ADODB.Recordset
    Dim sumAP_PAYMENT                             As Double
    sumAP_PAYMENT = 0
    Set rsCOMP_AP_AMT_PAID = New ADODB.Recordset
    rsCOMP_AP_AMT_PAID.Open "SELECT DISTINCT CD.PV_VOUCHERNO,HD.JTYPE FROM AMIS_VW_VLEDGER HD INNER JOIN AMIS_CV_DETAIL CD ON HD.VOUCHERNO=CD.VOUCHERNO AND HD.JTYPE=CD.CV_JTYPE WHERE CD.PV_VOUCHERNO = '" & xAP_VOUCHERnO & "' AND HD.JDATE < = '" & dtAsOF & "' AND HD.STATUS='P' AND HD.ACCT_CODE = '" & xACCT_CODE & "'", gconDMIS, adOpenKeyset
    If Not rsCOMP_AP_AMT_PAID.EOF And Not rsCOMP_AP_AMT_PAID.BOF Then
        Do While Not rsCOMP_AP_AMT_PAID.EOF
            'VALIDATE THE VENDOR CODE
            Call VAL_VEN_CODE(xAP_VOUCHERnO, xJType, xVEN_CODE)
            If xREMARKS = N2Str2Null("") Then
                Set rsCOMP_AP_AMT_PAID2 = New ADODB.Recordset
                rsCOMP_AP_AMT_PAID2.Open "SELECT DISTINCT HD.VOUCHERNO,CD.PV_VOUCHERNO,HD.JDATE,AMOUNT,HD.JTYPE FROM AMIS_VW_VLEDGER HD INNER JOIN AMIS_CV_DETAIL CD ON HD.VOUCHERNO=CD.VOUCHERNO AND HD.JTYPE=CD.CV_JTYPE WHERE CD.PV_VOUCHERNO = '" & xAP_VOUCHERnO & "' AND HD.JDATE < = '" & dtAsOF & "' AND HD.STATUS='P' AND HD.ACCT_CODE = '" & xACCT_CODE & "' ORDER BY HD.VOUCHERNO,HD.JDATE ASC", gconDMIS, adOpenKeyset
                If Not rsCOMP_AP_AMT_PAID2.EOF And Not rsCOMP_AP_AMT_PAID2.BOF Then
                    Do While Not rsCOMP_AP_AMT_PAID2.EOF
                        sumAP_PAYMENT = Round(sumAP_PAYMENT + NumericVal(rsCOMP_AP_AMT_PAID2!amount), 2)
                        rsCOMP_AP_AMT_PAID2.MoveNext
                    Loop
                End If
                Set rsCOMP_AP_AMT_PAID2 = Nothing
            Else
                'DON'T COMPUTE WRONG VENDOR CODE
                Set rsCOMP_AP_AMT_PAID3 = New ADODB.Recordset
                rsCOMP_AP_AMT_PAID3.Open "SELECT DISTINCT HD.VOUCHERNO,CD.PV_VOUCHERNO,HD.JDATE,AMOUNT,HD.JTYPE FROM AMIS_VW_VLEDGER HD INNER JOIN AMIS_CV_DETAIL CD ON HD.VOUCHERNO=CD.VOUCHERNO AND HD.JTYPE=CD.CV_JTYPE WHERE CD.PV_VOUCHERNO = '" & xAP_VOUCHERnO & "' AND HD.JDATE < = '" & dtAsOF & "' AND HD.STATUS='P' AND HD.ACCT_CODE = '" & xACCT_CODE & "' ORDER BY HD.VOUCHERNO,HD.JDATE ASC", gconDMIS, adOpenKeyset
                If Not rsCOMP_AP_AMT_PAID3.EOF And Not rsCOMP_AP_AMT_PAID3.BOF Then
                    Do While Not rsCOMP_AP_AMT_PAID3.EOF
                        sumAP_PAYMENT = Round(sumAP_PAYMENT + NumericVal(rsCOMP_AP_AMT_PAID3!amount), 2)
                        rsCOMP_AP_AMT_PAID3.MoveNext
                    Loop
                End If
                Set rsCOMP_AP_AMT_PAID3 = Nothing
            End If
            rsCOMP_AP_AMT_PAID.MoveNext
        Loop

    Else
        'THIS IS FOR DIRECT DISBURSEMENT MEANING IT HAS NO AP
        If xJType = "CDJ" Or xJType = "GJ" Or xJType = "SJ" Or xJType = "APJ" Then
            Dim rsNO_AP                           As ADODB.Recordset
            Set rsNO_AP = New ADODB.Recordset
            rsNO_AP.Open "SELECT DET.Debit AS DET_DEBIT, HD.VendorCode, HD.VoucherNo " & _
                         "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                         "WHERE HD.JTYPE = 'CDJ' AND DET.ACCT_CODE='" & xACCT_CODE & "' AND HD.VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE PV_VOUCHERNO = '" & xAP_VOUCHERnO & "') AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  VENDORCODE = '" & Trim(xVEN_CODE) & "' and Det.Debit <> 0 and HD.JTYPE = '" & xJType & "' and HD.STATUS='P' and HD.JDate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset
            If Not rsNO_AP.EOF And Not rsNO_AP.BOF Then
                Do While Not rsNO_AP.EOF
                    sumAP_PAYMENT = Round(sumAP_PAYMENT + NumericVal(rsNO_AP!DET_DEBIT), 2)
                    rsNO_AP.MoveNext
                Loop
            End If
            Set rsNO_AP = Nothing

            Dim rsNO_AP14                         As ADODB.Recordset
            Set rsNO_AP14 = New ADODB.Recordset
            rsNO_AP14.Open "SELECT DET.Debit AS DET_DEBIT, HD.VendorCode, HD.VoucherNo " & _
                           "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                           "WHERE (HD.JTYPE = 'APJ' or HD.JTYPE = 'GJ' or HD.JTYPE = 'SJ') AND DET.ACCT_CODE='" & xACCT_CODE & "' AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  VENDORCODE = '" & Trim(xVEN_CODE) & "' and Det.Debit <> 0 and HD.JTYPE = '" & xJType & "' and HD.STATUS='P' and HD.JDate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset
            If Not rsNO_AP14.EOF And Not rsNO_AP14.BOF Then
                Do While Not rsNO_AP14.EOF
                    sumAP_PAYMENT = Round(sumAP_PAYMENT + NumericVal(rsNO_AP14!DET_DEBIT), 2)
                    rsNO_AP14.MoveNext
                Loop
            End If
            Set rsNO_AP14 = Nothing

            Dim rsNO_AP3                          As ADODB.Recordset
            Set rsNO_AP3 = New ADODB.Recordset
            rsNO_AP3.Open "SELECT DET.Debit AS DET_DEBIT, HD.VendorCode, HD.VoucherNo " & _
                          "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                          "WHERE (HD.JTYPE = 'CDJ' or HD.JTYPE = 'GJ') AND DET.ACCT_CODE='" & xACCT_CODE & "' AND HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE DOCDATE > '" & dtAsOF & "' ) AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  VENDORCODE = '" & Trim(xVEN_CODE) & "' and HD.JDate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset
            If Not rsNO_AP3.EOF And Not rsNO_AP3.BOF Then
                Do While Not rsNO_AP3.EOF
                    sumAP_PAYMENT = Round(sumAP_PAYMENT + NumericVal(rsNO_AP3!DET_DEBIT), 2)
                    rsNO_AP3.MoveNext
                Loop
            End If
            Set rsNO_AP3 = Nothing
        End If
    End If
    COMP_AP_AMT_PAID = NumericVal(sumAP_PAYMENT)
    Set rsCOMP_AP_AMT_PAID = Nothing
    Set rsCOMP_AP_AMT_PAID2 = Nothing
End Function

Function COMP_AP_AMT_PAID2(xAP_VOUCHERnO As String, xJType As String, xVEN_CODE As String, xACCT_CODE) As Double
    Dim rsNO_AP4                                  As ADODB.Recordset
    Dim rsNO_AP5                                  As ADODB.Recordset
    Dim rsNO_AP6                                  As ADODB.Recordset
    Dim rsNO_AP7                                  As ADODB.Recordset
    Dim rsNO_AP8                                  As ADODB.Recordset
    Dim rsNO_AP9                                  As ADODB.Recordset
    Dim rsNO_AP10                                 As ADODB.Recordset
    Dim rsNO_AP11                                 As ADODB.Recordset
    Dim rsNO_AP12                                 As ADODB.Recordset
    Dim rsNO_AP13                                 As ADODB.Recordset
    Dim sumAP_PAYMENT                             As Double
    sumAP_PAYMENT = 0
    Set rsNO_AP4 = New ADODB.Recordset
    rsNO_AP4.Open "SELECT DET.Debit AS DET_DEBIT, HD.VendorCode, HD.VoucherNo " & _
                  "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                  "WHERE HD.JTYPE IN ('CDJ','GJ','CCM') AND DET.ACCT_CODE='" & xACCT_CODE & "' AND HD.VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  HD.VENDORCODE = '" & Trim(xVEN_CODE) & "' AND HD.STATUS='P' and HD.JDate < = '" & dtAsOF & "' AND DET.DEBIT <> 0", gconDMIS, adOpenKeyset
    If Not rsNO_AP4.EOF And Not rsNO_AP4.BOF Then
        Do While Not rsNO_AP4.EOF
            sumAP_PAYMENT = sumAP_PAYMENT + NumericVal(rsNO_AP4!DET_DEBIT)
            rsNO_AP4.MoveNext
        Loop
    End If

    Set rsNO_AP5 = New ADODB.Recordset
    rsNO_AP5.Open "SELECT HD.VOUCHERNO AS HD_VOUCHERNO,HD.JTYPE AS HD_JTYPE,HD.JDATE AS HD_JDATE,HD.STATUS AS HD_STATUS,HD.INVOICETYPE AS HD_INV_TYPE,HD.INVOICEDATE AS HD_INV_DATE, HD.VENDORCODE AS HD_VEN_CODE, HD.INVOICEAMT AS HD_INV_AMT,HD.AMOUNTTOPAY AS HD_AMT_TO_PAY,HD.AMOUNTPAID AS HD_AMT_PAID,DET.DEBIT AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.ACCT_CODE AS DET_ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE " & _
                  "WHERE (HD.JTYPE = 'CDJ' OR HD.JTYPE = 'GJ') AND DET.ACCT_CODE='" & xACCT_CODE & "' " & _
                  "AND HD.STATUS = 'P' AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  HD.VENDORCODE = '" & Trim(xVEN_CODE) & "' AND HD.JDATE < = '" & dtAsOF & "' AND HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE DOCDATE > '" & dtAsOF & "' AND DET.DEBIT=AMOUNT)", gconDMIS, adOpenKeyset
    If Not rsNO_AP5.EOF And Not rsNO_AP5.BOF Then
        Do While Not rsNO_AP5.EOF
            sumAP_PAYMENT = sumAP_PAYMENT + NumericVal(rsNO_AP5!DET_DEBIT)
            rsNO_AP5.MoveNext
        Loop
    End If

    Set rsNO_AP6 = New ADODB.Recordset
    rsNO_AP6.Open "SELECT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE FROM (" & _
                  "SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL ) AND DET.ACCT_CODE='" & xACCT_CODE & "' AND HD.status = 'P' AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  HD.VENDORCODE = '" & Trim(xVEN_CODE) & "' AND HD.jdate < = '" & dtAsOF & "'" & _
                  ")X WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE JTYPE ='APJ' AND ACCT_CODE='11-02000-00' AND DEBIT=0)", gconDMIS, adOpenKeyset
    If Not rsNO_AP6.EOF And Not rsNO_AP6.BOF Then
        Do While Not rsNO_AP6.EOF
            sumAP_PAYMENT = sumAP_PAYMENT + NumericVal(rsNO_AP6!DET_DEBIT)
            rsNO_AP6.MoveNext
        Loop
    End If

    Set rsNO_AP7 = New ADODB.Recordset
    rsNO_AP7.Open "SELECT HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
                  "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, " & _
                  "HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                  "WHERE (HD.JTYPE = 'CDJ') AND DET.ACCT_CODE='21-01008-00' AND HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE CV_JTYPE = 'CDJ' AND PV_VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_HD)) " & _
                  "AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  HD.VENDORCODE = '" & Trim(xVEN_CODE) & "' AND HD.status = 'P' and HD.jdate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset
    If Not rsNO_AP7.EOF And Not rsNO_AP7.BOF Then
        Do While Not rsNO_AP7.EOF
            sumAP_PAYMENT = sumAP_PAYMENT + NumericVal(rsNO_AP7!DET_DEBIT)
            rsNO_AP7.MoveNext
        Loop
    End If

    Set rsNO_AP8 = New ADODB.Recordset
    rsNO_AP8.Open "SELECT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE FROM (SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE " & _
                  "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL ) AND LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-07') AND DET.ACCT_CODE='" & xACCT_CODE & "' AND HD.status = 'P' AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  HD.VENDORCODE = '" & Trim(xVEN_CODE) & "' AND HD.jdate < = '" & dtAsOF & "')X " & _
                  "WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE JTYPE ='APJ' AND X.DET_DEBIT=DEBIT AND ACCT_CODE='21-01002-00' AND X.DET_DEBIT <> 0) ORDER BY AP_VOUCHERNO", gconDMIS, adOpenKeyset
    If Not rsNO_AP8.EOF And Not rsNO_AP8.BOF Then
        Do While Not rsNO_AP8.EOF
            sumAP_PAYMENT = sumAP_PAYMENT + NumericVal(rsNO_AP8!DET_DEBIT)
            rsNO_AP8.MoveNext
        Loop
    End If

    Set rsNO_AP9 = New ADODB.Recordset
    rsNO_AP9.Open "SELECT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE " & _
                  "FROM (SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
                  "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE " & _
                  "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND DET.ACCT_CODE = ('21-02000-00') AND HD.status = 'P' AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  HD.VENDORCODE = '" & Trim(xVEN_CODE) & "' AND HD.jdate < = '" & dtAsOF & "')X " & _
                  "WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE ACCT_CODE <> X.DET_ACCT_CODE AND DEBIT=0) ORDER BY AP_VOUCHERNO", gconDMIS, adOpenKeyset
    If Not rsNO_AP9.EOF And Not rsNO_AP9.BOF Then
        Do While Not rsNO_AP9.EOF
            sumAP_PAYMENT = sumAP_PAYMENT + NumericVal(rsNO_AP9!DET_DEBIT)
            rsNO_AP9.MoveNext
        Loop
    End If

    Set rsNO_AP10 = New ADODB.Recordset
    rsNO_AP10.Open "SELECT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE " & _
                   "FROM (SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
                   "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE " & _
                   "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN " & _
                   "(SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND DET.ACCT_CODE = ('21-02004-00') AND HD.status = 'P' AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  HD.VENDORCODE = '" & Trim(xVEN_CODE) & "' AND HD.jdate < = '" & dtAsOF & "')X " & _
                   "WHERE X.AP_VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE JTYPE='APJ') ORDER BY AP_VOUCHERNO", gconDMIS, adOpenKeyset
    If Not rsNO_AP10.EOF And Not rsNO_AP10.BOF Then
        Do While Not rsNO_AP10.EOF
            sumAP_PAYMENT = sumAP_PAYMENT + NumericVal(rsNO_AP10!DET_DEBIT)
            rsNO_AP10.MoveNext
        Loop
    End If

    Set rsNO_AP11 = New ADODB.Recordset
    rsNO_AP11.Open "SELECT DET.Debit AS DET_DEBIT, HD.VendorCode, HD.VoucherNo " & _
                   "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                   "WHERE (HD.JTYPE = 'VDJ' or HD.JTYPE = 'VCJ') AND DET.ACCT_CODE='" & xACCT_CODE & "' AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  HD.VENDORCODE = '" & Trim(xVEN_CODE) & "' AND HD.STATUS='P' and HD.JDate < = '" & dtAsOF & "' AND DET.DEBIT <> 0", gconDMIS, adOpenKeyset
    If Not rsNO_AP11.EOF And Not rsNO_AP11.BOF Then
        Do While Not rsNO_AP11.EOF
            sumAP_PAYMENT = sumAP_PAYMENT + NumericVal(rsNO_AP11!DET_DEBIT)
            rsNO_AP11.MoveNext
        Loop
    End If

    Set rsNO_AP12 = New ADODB.Recordset
    rsNO_AP12.Open "SELECT DISTINCT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_CREDIT,X.HD_DUEDATE " & _
                   "FROM (SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VOUCHERNO AS HD_VOUCHERNO,HD.JTYPE AS HD_JTYPE,HD.JDATE AS HD_JDATE,HD.STATUS AS HD_STATUS,HD.INVOICETYPE AS HD_INV_TYPE,HD.INVOICEDATE AS HD_INV_DATE, " & _
                   "HD.VENDORCODE AS HD_VEN_CODE, HD.INVOICEAMT AS HD_INV_AMT,HD.AMOUNTTOPAY AS HD_AMT_TO_PAY,HD.AMOUNTPAID AS HD_AMT_PAID,DET.DEBIT AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE, " & _
                   "DET.ACCT_CODE AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO " & _
                   "AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND DET.ACCT_CODE = ('21-01002-00') AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  HD.VENDORCODE = '" & Trim(xVEN_CODE) & "' AND HD.STATUS = 'P' AND HD.JDATE < = '" & dtAsOF & "')X " & _
                   "WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE ACCT_CODE = '21-01004-00')", gconDMIS, adOpenKeyset
    If Not rsNO_AP12.EOF And Not rsNO_AP12.BOF Then
        Do While Not rsNO_AP12.EOF
            Dim rsWRONGENTRY                      As ADODB.Recordset
            Set rsWRONGENTRY = New ADODB.Recordset
            rsWRONGENTRY.Open "SELECT VOUCHERNO,JTYPE,ACCT_CODE,DEBIT,CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO= " & N2Str2Null(rsNO_AP12!AP_VOUCHERNO) & " AND JTYPE='APJ' AND ACCT_CODE = '21-01004-00'", gconDMIS, adOpenKeyset
            If Not rsWRONGENTRY.EOF And Not rsWRONGENTRY.BOF Then
                Do While Not rsWRONGENTRY.EOF
                    sumAP_PAYMENT = sumAP_PAYMENT + NumericVal(rsWRONGENTRY!CREDIT)
                    rsWRONGENTRY.MoveNext
                Loop
            End If
            rsNO_AP12.MoveNext
        Loop
    End If

    Set rsNO_AP13 = New ADODB.Recordset
    rsNO_AP13.Open "SELECT DISTINCT X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE " & _
                   "FROM (SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
                   "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE " & _
                   "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE " & _
                   "WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND DET.ACCT_CODE = ('21-01002-00') AND HD.status = 'P' AND DET.ACCT_CODE='" & xACCT_CODE & "' AND HD.VOUCHERNO = '" & xAP_VOUCHERnO & "' AND  HD.VENDORCODE = '" & Trim(xVEN_CODE) & "' AND HD.jdate < = '" & dtAsOF & "')X " & _
                   "WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE JTYPE='APJ' AND ACCT_CODE='21-02004-00')", gconDMIS, adOpenKeyset
    If Not rsNO_AP13.EOF And Not rsNO_AP13.BOF Then
        Do While Not rsNO_AP13.EOF
            sumAP_PAYMENT = sumAP_PAYMENT + NumericVal(rsNO_AP13!DET_DEBIT)
            rsNO_AP13.MoveNext
        Loop
    End If

    COMP_AP_AMT_PAID2 = NumericVal(sumAP_PAYMENT)
    Set rsNO_AP4 = Nothing
    Set rsNO_AP5 = Nothing
    Set rsNO_AP6 = Nothing
    Set rsNO_AP7 = Nothing
    Set rsNO_AP8 = Nothing
    Set rsNO_AP9 = Nothing
    Set rsNO_AP10 = Nothing
    Set rsNO_AP11 = Nothing
    Set rsNO_AP12 = Nothing
    Set rsNO_AP13 = Nothing
End Function

Function COM_AP_ADJ_CREDIT(xEntity As String, xINVOICENO As String, Optional xINVOICETYPE As String, Optional xJType As String) As Double
'DESCRIPTION: COMPUTE THE ADJUSTMENT AS ADDITIONAL AMOUNT TO PAY TO ACCOUNTS PAYABLES
    Dim rsCOM_AP_ADJ_CREDIT                       As ADODB.Recordset
    Set rsCOM_AP_ADJ_CREDIT = New ADODB.Recordset

    If xJType = "'CDJ'" Then
        rsCOM_AP_ADJ_CREDIT.Open "SELECT ROUND(SUM(CREDIT),2) AS CREDIT FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('21-01','21-02','21-07') AND INVOICENO = " & xINVOICENO & " AND RIGHT(ENTITY,6) = " & xEntity & " AND ADJ_JTYPE = " & xJType & "", gconDMIS, adOpenKeyset
    Else
        rsCOM_AP_ADJ_CREDIT.Open "SELECT ROUND(SUM(CREDIT),2) AS CREDIT FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('21-01','21-02','21-07') AND INVOICENO = " & xINVOICENO & " AND INVOICETYPE = " & xINVOICETYPE & " AND RIGHT(ENTITY,6) = " & xEntity & "", gconDMIS, adOpenKeyset
    End If

    If Not rsCOM_AP_ADJ_CREDIT.EOF And Not rsCOM_AP_ADJ_CREDIT.BOF Then
        COM_AP_ADJ_CREDIT = NumericVal(rsCOM_AP_ADJ_CREDIT!CREDIT)
    Else
        COM_AP_ADJ_CREDIT = 0
    End If
    Set rsCOM_AP_ADJ_CREDIT = Nothing
End Function

Function COM_AP_AMT_TO_PAY(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'DESCRIPTION: COMPUTE THE AMOUNT TO PAY IN DETAIL WHICH IS SCHEDULE ACCOUNT
    Dim rsCOM_AP_AMT_TO_PAY                       As ADODB.Recordset
    Set rsCOM_AP_AMT_TO_PAY = New ADODB.Recordset
    If xJType = "'CDJ'" Then
        'rsCOM_AP_AMT_TO_PAY.Open "SELECT ROUND(SUM(DEBIT),2) AS DEBIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & xVOUCHERNO & " AND JTYPE = " & xJTYPE & " AND STATUS = 'P' AND LEFT(ACCT_CODE,5) IN('21-01','21-02') ", gconDMIS, adOpenKeyset
        rsCOM_AP_AMT_TO_PAY.Open "SELECT ROUND(SUM(CREDIT),2) AS CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & xVOUCHERNO & " AND JTYPE = " & xJType & " AND STATUS = 'P' AND ACCT_CODE= '" & xACCT_CODE & "' ", gconDMIS, adOpenKeyset
        If Not rsCOM_AP_AMT_TO_PAY.EOF And Not rsCOM_AP_AMT_TO_PAY.BOF Then
            COM_AP_AMT_TO_PAY = NumericVal(rsCOM_AP_AMT_TO_PAY!DEBIT)
        End If
    Else
        'rsCOM_AP_AMT_TO_PAY.Open "SELECT ROUND(SUM(CREDIT),2) AS CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & xVOUCHERNO & " AND JTYPE = " & xJTYPE & " AND STATUS = 'P' AND LEFT(ACCT_CODE,5) IN('21-01','21-02') ", gconDMIS, adOpenKeyset
        rsCOM_AP_AMT_TO_PAY.Open "SELECT ROUND(SUM(CREDIT),2) AS CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & xVOUCHERNO & " AND JTYPE = " & xJType & " AND STATUS = 'P' AND ACCT_CODE= " & xACCT_CODE & " ", gconDMIS, adOpenKeyset
        If Not rsCOM_AP_AMT_TO_PAY.EOF And Not rsCOM_AP_AMT_TO_PAY.BOF Then
            COM_AP_AMT_TO_PAY = NumericVal(rsCOM_AP_AMT_TO_PAY!CREDIT)
        End If
    End If

    Set rsCOM_AP_AMT_TO_PAY = Nothing
End Function

Function COMP_AP_ADJ_DEBIT(xEntity As String, xINVOICENO As String, Optional xINVOICETYPE As String, Optional xJType As String) As Double
'DESCRIPTION: COMPUTE THE AJUSTMENT AS ADDITIONAL PAYMENT TO THE ACCOUNT PAYABLE
    Dim rsCOMP_AP_ADJ_DEBIT                       As ADODB.Recordset
    Set rsCOMP_AP_ADJ_DEBIT = New ADODB.Recordset
    If xJType = "CDJ" Then
        rsCOMP_AP_ADJ_DEBIT.Open "SELECT ROUND(SUM(DEBIT),2) AS DEBIT FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('21-01','21-02','21-07') AND INVOICENO = " & xINVOICENO & " AND RIGHT(ENTITY,6) = " & xEntity & " AND ADJ_JTYPE = '" & xJType & "'", gconDMIS, adOpenKeyset
    Else
        rsCOMP_AP_ADJ_DEBIT.Open "SELECT ROUND(SUM(DEBIT),2) AS DEBIT FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('21-01','21-02','21-07') AND INVOICENO = " & xINVOICENO & " AND INVOICETYPE = " & xINVOICETYPE & " AND RIGHT(ENTITY,6) = " & xEntity & "", gconDMIS, adOpenKeyset
    End If

    If Not rsCOMP_AP_ADJ_DEBIT.EOF And Not rsCOMP_AP_ADJ_DEBIT.BOF Then
        COMP_AP_ADJ_DEBIT = NumericVal(rsCOMP_AP_ADJ_DEBIT!DEBIT)
    Else
        COMP_AP_ADJ_DEBIT = 0
    End If
    Set rsCOMP_AP_ADJ_DEBIT = Nothing
End Function

Sub DISBURSEMENT1()
'DESCRIPTION: THIS SELECT ALL DIRECT DISBURSEMENT WHICH IS NO LINK TO THE ACCOUNTS PAYABLES
'             AND DETAIL DATE GREATER THE JDATE
    Dim rsDISBURSEMENT1                           As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJType                                    As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xVENDORCODE                               As String
    Dim xDUEDATE                                  As String
    Dim xACCOUNTCODE                              As String

    Dim xINVOICEAMOUNT                            As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    xAMOUNTPAID = 0
    Set rsDISBURSEMENT1 = New ADODB.Recordset
    rsDISBURSEMENT1.Open "SELECT HD.VOUCHERNO AS HD_VOUCHERNO,HD.JTYPE AS HD_JTYPE,HD.JDATE AS HD_JDATE,HD.STATUS AS HD_STATUS,HD.INVOICETYPE AS HD_INV_TYPE,HD.INVOICEDATE AS HD_INV_DATE, HD.VENDORCODE AS HD_VEN_CODE, HD.INVOICEAMT AS HD_INV_AMT,HD.AMOUNTTOPAY AS HD_AMT_TO_PAY,HD.AMOUNTPAID AS HD_AMT_PAID,DET.DEBIT AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.ACCT_CODE AS DET_ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE " & _
                         "WHERE (HD.JTYPE = 'CDJ' OR HD.JTYPE = 'GJ') AND LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-07') " & _
                         "AND HD.STATUS = 'P' AND HD.JDATE < = '" & dtAsOF & "' AND HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE DOCDATE > '" & dtAsOF & "' AND DET.DEBIT=AMOUNT)", gconDMIS, adOpenKeyset
    Label11.Caption = " Processing CDJ... Please Wait.."

    If rsDISBURSEMENT1.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsDISBURSEMENT1.RecordCount

    If Not rsDISBURSEMENT1.EOF And Not rsDISBURSEMENT1.BOF Then
        Do While Not rsDISBURSEMENT1.EOF
            xVOUCHERNO = N2Str2Null(rsDISBURSEMENT1!HD_VOUCHERNO)
            xJType = N2Str2Null(rsDISBURSEMENT1!HD_JTYPE)
            xJdate = N2Date2Null(rsDISBURSEMENT1!HD_JDATE)
            xSTATUS = N2Str2Null(rsDISBURSEMENT1!HD_STATUS)
            xInvoicedate = N2Date2Null(rsDISBURSEMENT1!HD_INV_DATE)
            xVENDORCODE = N2Str2Null(rsDISBURSEMENT1!HD_VEN_CODE)
            xINVOICETYPE = N2Str2Null(rsDISBURSEMENT1!HD_INV_TYPE)
            xINVOICEAMOUNT = NumericVal(rsDISBURSEMENT1!HD_INV_AMT)
            xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT1!HD_AMT_TO_PAY) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJType, xVENDORCODE)), 2)
            xAMOUNTPAID = GET_DRCT_AMOUNT(xVOUCHERNO, xJType)
            'NumericVal(rsDISBURSEMENT1!HD_AMT_PAID)
            xdebit = NumericVal(rsDISBURSEMENT1!DET_DEBIT)
            xcredit = NumericVal(rsDISBURSEMENT1!DET_CREDIT)
            xDUEDATE = N2Date2Null(rsDISBURSEMENT1!HD_DUEDATE)
            xACCOUNTCODE = N2Str2Null(rsDISBURSEMENT1!DET_ACCT_CODE)

            gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & "," & xINVOICETYPE & "," & xInvoicedate & "," & xINVOICEAMOUNT & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCOUNTCODE & ", " & xDUEDATE & ")"

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsDISBURSEMENT1!HD_VOUCHERNO)
            DoEvents

            rsDISBURSEMENT1.MoveNext
        Loop
    End If
    Set rsDISBURSEMENT1 = Nothing
End Sub

Sub DISBURSEMENT2()
'DESCRIPTION: THIS SELECT ALL DIRECT DISBURSEMENT WHICH IS NO LINK TO THE ACCOUNTS PAYABLES
    Dim rsDISBURSEMENT2                           As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJType                                    As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xVENDORCODE                               As String
    Dim xDUEDATE                                  As String
    Dim xACCOUNTCODE                              As String

    Dim xINVOICEAMOUNT                            As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    xAMOUNTPAID = 0
    Set rsDISBURSEMENT2 = New ADODB.Recordset
    'FOR DEBUGGING ONLY
    'rsDISBURSEMENT2.Open "SELECT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE FROM (" & _
     "SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL ) AND DET.ACCT_CODE='21-01008-00' AND HD.status = 'P' AND HD.jdate < = '" & dtAsOF & "'" & _
     ")X WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE JTYPE ='APJ' AND ACCT_CODE='11-02000-00' AND DEBIT=0)", gconDMIS, adOpenKeyset
    '=======================

    rsDISBURSEMENT2.Open "SELECT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE FROM (" & _
                         "SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL ) AND LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-07') AND HD.status = 'P' AND HD.jdate < = '" & dtAsOF & "'" & _
                         ")X WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE JTYPE ='APJ' AND ACCT_CODE='11-02000-00' AND DEBIT=0)", gconDMIS, adOpenKeyset

    Label11.Caption = " Processing CDJ... Please Wait.."

    If rsDISBURSEMENT2.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsDISBURSEMENT2.RecordCount

    If Not rsDISBURSEMENT2.EOF And Not rsDISBURSEMENT2.BOF Then
        Do While Not rsDISBURSEMENT2.EOF
            xVOUCHERNO = N2Str2Null(rsDISBURSEMENT2!HD_VOUCHERNO)
            xJType = N2Str2Null(rsDISBURSEMENT2!HD_JTYPE)
            xJdate = N2Date2Null(rsDISBURSEMENT2!HD_JDATE)
            xSTATUS = N2Str2Null(rsDISBURSEMENT2!HD_STATUS)
            xInvoicedate = N2Date2Null(rsDISBURSEMENT2!HD_INV_DATE)
            xVENDORCODE = N2Str2Null(rsDISBURSEMENT2!HD_VEN_CODE)
            xINVOICETYPE = N2Str2Null(rsDISBURSEMENT2!HD_INV_TYPE)
            xINVOICEAMOUNT = NumericVal(rsDISBURSEMENT2!HD_INV_AMT)
            xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT2!HD_AMT_TO_PAY) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJType, xVENDORCODE)), 2)
            xAMOUNTPAID = GET_DRCT_AMOUNT(xVOUCHERNO, xJType)
            'NumericVal(rsDISBURSEMENT2!HD_AMT_PAID)
            xdebit = NumericVal(rsDISBURSEMENT2!DET_DEBIT)
            xcredit = NumericVal(rsDISBURSEMENT2!DET_CREDIT)
            xDUEDATE = N2Date2Null(rsDISBURSEMENT2!HD_DUEDATE)
            xACCOUNTCODE = N2Str2Null(rsDISBURSEMENT2!DET_ACCT_CODE)

            gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & "," & xINVOICETYPE & "," & xInvoicedate & "," & xINVOICEAMOUNT & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCOUNTCODE & ", " & xDUEDATE & ")"

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsDISBURSEMENT2!HD_VOUCHERNO)
            DoEvents

            rsDISBURSEMENT2.MoveNext
        Loop
    End If
    Set rsDISBURSEMENT2 = Nothing
End Sub

Sub DISBURSEMENT3()
    Dim rsDISBURSEMENT3                           As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJType                                    As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xVENDORCODE                               As String
    Dim xDUEDATE                                  As String
    Dim xACCOUNTCODE                              As String

    Dim xINVOICEAMOUNT                            As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    xAMOUNTPAID = 0
    Set rsDISBURSEMENT3 = New ADODB.Recordset
    'FOR DEBUGGING
    'rsDISBURSEMENT3.Open "SELECT HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
     "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, " & _
     "HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
     "WHERE HD.JTYPE = 'SJ' AND DET.ACCT_CODE = '21-01008-00' AND HD.status = 'P' and HD.jdate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset

    rsDISBURSEMENT3.Open "SELECT HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
                         "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, " & _
                         "HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                         "WHERE HD.JTYPE = 'SJ' AND LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-07') AND HD.status = 'P' and HD.jdate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset
    Label11.Caption = " Processing AP... Please Wait.."

    If rsDISBURSEMENT3.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsDISBURSEMENT3.RecordCount

    If Not rsDISBURSEMENT3.EOF And Not rsDISBURSEMENT3.BOF Then
        Do While Not rsDISBURSEMENT3.EOF
            xVOUCHERNO = N2Str2Null(rsDISBURSEMENT3!HD_VOUCHERNO)
            xJType = N2Str2Null(rsDISBURSEMENT3!HD_JTYPE)
            xJdate = N2Date2Null(rsDISBURSEMENT3!HD_JDATE)
            xSTATUS = N2Str2Null(rsDISBURSEMENT3!HD_STATUS)
            xInvoicedate = N2Date2Null(rsDISBURSEMENT3!HD_INV_DATE)
            xVENDORCODE = N2Str2Null(rsDISBURSEMENT3!HD_VEN_CODE)
            xINVOICETYPE = N2Str2Null(rsDISBURSEMENT3!HD_INV_TYPE)
            xINVOICEAMOUNT = NumericVal(rsDISBURSEMENT3!HD_INV_AMT)
            '            If xJTYPE = "'CDJ'" Then
            '                xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT3!HD_AMT_TO_PAY) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJTYPE, xVENDORCODE)), 2)
            '            Else
            xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT3!DET_CREDIT) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJType, xVENDORCODE)), 2)
            '            End If
            xAMOUNTPAID = GET_DRCT_AMOUNT(xVOUCHERNO, xJType)
            'NumericVal(rsDISBURSEMENT3!HD_AMT_PAID)
            xdebit = NumericVal(rsDISBURSEMENT3!DET_DEBIT)
            xcredit = NumericVal(rsDISBURSEMENT3!DET_CREDIT)
            xDUEDATE = N2Date2Null(rsDISBURSEMENT3!HD_DUEDATE)
            xACCOUNTCODE = N2Str2Null(rsDISBURSEMENT3!DET_ACCT_CODE)

            gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & "," & xINVOICETYPE & "," & xInvoicedate & "," & xINVOICEAMOUNT & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCOUNTCODE & ", " & xDUEDATE & ")"

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsDISBURSEMENT3!HD_VOUCHERNO)
            DoEvents

            rsDISBURSEMENT3.MoveNext
        Loop
    End If
    Set rsDISBURSEMENT3 = Nothing
End Sub

Sub DISBURSEMENT4()
    Dim rsDISBURSEMENT4                           As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJType                                    As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xVENDORCODE                               As String
    Dim xDUEDATE                                  As String
    Dim xACCOUNTCODE                              As String

    Dim xINVOICEAMOUNT                            As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    xAMOUNTPAID = 0
    Set rsDISBURSEMENT4 = New ADODB.Recordset

    'rsDISBURSEMENT4.Open "SELECT HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
     "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, " & _
     "HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
     "WHERE (HD.JTYPE = 'GJ') AND LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02') AND HD.VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE CV_JTYPE = 'CDJ') AND HD.status = 'P' and HD.jdate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset

    rsDISBURSEMENT4.Open "SELECT HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
                         "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, " & _
                         "HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                         "WHERE (HD.JTYPE = 'CDJ') AND DET.ACCT_CODE='21-01008-00' AND HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE CV_JTYPE = 'CDJ' AND PV_VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_HD)) " & _
                         "AND HD.status = 'P' and HD.jdate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset

    Label11.Caption = " Processing CDJ... Please Wait.."

    If rsDISBURSEMENT4.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsDISBURSEMENT4.RecordCount

    If Not rsDISBURSEMENT4.EOF And Not rsDISBURSEMENT4.BOF Then
        Do While Not rsDISBURSEMENT4.EOF
            xVOUCHERNO = N2Str2Null(rsDISBURSEMENT4!HD_VOUCHERNO)
            xJType = N2Str2Null(rsDISBURSEMENT4!HD_JTYPE)
            xJdate = N2Date2Null(rsDISBURSEMENT4!HD_JDATE)
            xSTATUS = N2Str2Null(rsDISBURSEMENT4!HD_STATUS)
            xInvoicedate = N2Date2Null(rsDISBURSEMENT4!HD_INV_DATE)
            xVENDORCODE = N2Str2Null(rsDISBURSEMENT4!HD_VEN_CODE)
            xINVOICETYPE = N2Str2Null(rsDISBURSEMENT4!HD_INV_TYPE)
            xINVOICEAMOUNT = NumericVal(rsDISBURSEMENT4!HD_INV_AMT)
            'If xJTYPE = "'CDJ'" Then
            xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT4!HD_AMT_TO_PAY) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJType, xVENDORCODE)), 2)
            'Else
            '    xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT4!DET_CREDIT) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJTYPE, xVENDORCODE)), 2)
            'End If
            xAMOUNTPAID = GET_DRCT_AMOUNT(xVOUCHERNO, xJType)
            'NumericVal(rsDISBURSEMENT4!HD_AMT_PAID)
            xdebit = NumericVal(rsDISBURSEMENT4!DET_DEBIT)
            xcredit = NumericVal(rsDISBURSEMENT4!DET_CREDIT)
            xDUEDATE = N2Date2Null(rsDISBURSEMENT4!HD_DUEDATE)
            xACCOUNTCODE = N2Str2Null(rsDISBURSEMENT4!DET_ACCT_CODE)

            gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & "," & xINVOICETYPE & "," & xInvoicedate & "," & xINVOICEAMOUNT & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCOUNTCODE & ", " & xDUEDATE & ")"

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsDISBURSEMENT4!HD_VOUCHERNO)
            DoEvents

            rsDISBURSEMENT4.MoveNext
        Loop
    End If
    Set rsDISBURSEMENT4 = Nothing
End Sub

Sub DISBURSEMENT5()
'DESCRIPTION: THIS SELECT ALL DIRECT DISBURSEMENT WHICH IS NO LINK TO THE ACCOUNTS PAYABLES
    Dim rsDISBURSEMENT5                           As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJType                                    As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xVENDORCODE                               As String
    Dim xDUEDATE                                  As String
    Dim xACCOUNTCODE                              As String

    Dim xINVOICEAMOUNT                            As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    xAMOUNTPAID = 0
    Set rsDISBURSEMENT5 = New ADODB.Recordset
    'FOR DEBUGGING ONLY
    'rsDISBURSEMENT5.Open "SELECT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE FROM (" & _
     "SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL ) AND DET.ACCT_CODE='21-01008-00' AND HD.status = 'P' AND HD.jdate < = '" & dtAsOF & "'" & _
     ")X WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE JTYPE ='APJ' AND ACCT_CODE='11-02000-00' AND DEBIT=0)", gconDMIS, adOpenKeyset
    '=======================

    rsDISBURSEMENT5.Open "SELECT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE FROM (SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE " & _
                         "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL ) AND LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-07') AND HD.status = 'P' AND HD.jdate < = '" & dtAsOF & "')X " & _
                         "WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE JTYPE ='APJ' AND X.DET_DEBIT=DEBIT AND ACCT_CODE='21-01002-00' AND X.DET_DEBIT <> 0) ORDER BY AP_VOUCHERNO", gconDMIS, adOpenKeyset

    Label11.Caption = " Processing CDJ... Please Wait.."

    If rsDISBURSEMENT5.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsDISBURSEMENT5.RecordCount

    If Not rsDISBURSEMENT5.EOF And Not rsDISBURSEMENT5.BOF Then
        Do While Not rsDISBURSEMENT5.EOF
            xVOUCHERNO = N2Str2Null(rsDISBURSEMENT5!HD_VOUCHERNO)
            xJType = N2Str2Null(rsDISBURSEMENT5!HD_JTYPE)
            xJdate = N2Date2Null(rsDISBURSEMENT5!HD_JDATE)
            xSTATUS = N2Str2Null(rsDISBURSEMENT5!HD_STATUS)
            xInvoicedate = N2Date2Null(rsDISBURSEMENT5!HD_INV_DATE)
            xVENDORCODE = N2Str2Null(rsDISBURSEMENT5!HD_VEN_CODE)
            xINVOICETYPE = N2Str2Null(rsDISBURSEMENT5!HD_INV_TYPE)
            xINVOICEAMOUNT = NumericVal(rsDISBURSEMENT5!HD_INV_AMT)
            xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT5!HD_AMT_TO_PAY) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJType, xVENDORCODE)), 2)
            xAMOUNTPAID = GET_DRCT_AMOUNT(xVOUCHERNO, xJType)
            'NumericVal(rsDISBURSEMENT5!HD_AMT_PAID)
            xdebit = NumericVal(rsDISBURSEMENT5!DET_DEBIT)
            xcredit = NumericVal(rsDISBURSEMENT5!DET_CREDIT)
            xDUEDATE = N2Date2Null(rsDISBURSEMENT5!HD_DUEDATE)
            xACCOUNTCODE = N2Str2Null(rsDISBURSEMENT5!DET_ACCT_CODE)

            gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & "," & xINVOICETYPE & "," & xInvoicedate & "," & xINVOICEAMOUNT & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCOUNTCODE & ", " & xDUEDATE & ")"

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsDISBURSEMENT5!HD_VOUCHERNO)
            DoEvents

            rsDISBURSEMENT5.MoveNext
        Loop
    End If
    Set rsDISBURSEMENT5 = Nothing
End Sub

Sub DISBURSEMENT6()
'DESCRIPTION: THIS SELECT ALL DIRECT DISBURSEMENT WHICH HAS LINK TO THE ACCOUNTS PAYABLES BUT DIFFERENT ACCOUNT CODES
    Dim rsDISBURSEMENT6                           As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJType                                    As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xVENDORCODE                               As String
    Dim xDUEDATE                                  As String
    Dim xACCOUNTCODE                              As String

    Dim xINVOICEAMOUNT                            As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    xAMOUNTPAID = 0
    Set rsDISBURSEMENT6 = New ADODB.Recordset
    'FOR DEBUGGING ONLY
    'rsDISBURSEMENT6.Open "SELECT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE FROM (" & _
     "SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL ) AND DET.ACCT_CODE='21-01008-00' AND HD.status = 'P' AND HD.jdate < = '" & dtAsOF & "'" & _
     ")X WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE JTYPE ='APJ' AND ACCT_CODE='11-02000-00' AND DEBIT=0)", gconDMIS, adOpenKeyset
    '=======================

    rsDISBURSEMENT6.Open "SELECT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE " & _
                         "FROM (SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND DET.ACCT_CODE = ('21-02000-00') AND HD.status = 'P' AND HD.jdate < = '" & dtAsOF & "')X " & _
                         "WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE ACCT_CODE <> X.DET_ACCT_CODE AND DEBIT=0) ORDER BY AP_VOUCHERNO", gconDMIS, adOpenKeyset

    Label11.Caption = " Processing CDJ... Please Wait.."

    If rsDISBURSEMENT6.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsDISBURSEMENT6.RecordCount

    If Not rsDISBURSEMENT6.EOF And Not rsDISBURSEMENT6.BOF Then
        Do While Not rsDISBURSEMENT6.EOF
            xVOUCHERNO = N2Str2Null(rsDISBURSEMENT6!HD_VOUCHERNO)
            xJType = N2Str2Null(rsDISBURSEMENT6!HD_JTYPE)
            xJdate = N2Date2Null(rsDISBURSEMENT6!HD_JDATE)
            xSTATUS = N2Str2Null(rsDISBURSEMENT6!HD_STATUS)
            xInvoicedate = N2Date2Null(rsDISBURSEMENT6!HD_INV_DATE)
            xVENDORCODE = N2Str2Null(rsDISBURSEMENT6!HD_VEN_CODE)
            xINVOICETYPE = N2Str2Null(rsDISBURSEMENT6!HD_INV_TYPE)
            xINVOICEAMOUNT = NumericVal(rsDISBURSEMENT6!HD_INV_AMT)
            xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT6!HD_AMT_TO_PAY) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJType, xVENDORCODE)), 2)
            xAMOUNTPAID = GET_DRCT_AMOUNT(xVOUCHERNO, xJType)
            'NumericVal(rsDISBURSEMENT6!HD_AMT_PAID)
            xdebit = NumericVal(rsDISBURSEMENT6!DET_DEBIT)
            xcredit = NumericVal(rsDISBURSEMENT6!DET_CREDIT)
            xDUEDATE = N2Date2Null(rsDISBURSEMENT6!HD_DUEDATE)
            xACCOUNTCODE = N2Str2Null(rsDISBURSEMENT6!DET_ACCT_CODE)

            gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & "," & xINVOICETYPE & "," & xInvoicedate & "," & xINVOICEAMOUNT & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCOUNTCODE & ", " & xDUEDATE & ")"

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsDISBURSEMENT6!HD_VOUCHERNO)
            DoEvents

            rsDISBURSEMENT6.MoveNext
        Loop
    End If
    Set rsDISBURSEMENT6 = Nothing
End Sub

Sub DISBURSEMENT7()
'DESCRIPTION: THIS SELECT ALL DIRECT DISBURSEMENT WHICH IS NO LINK TO THE ACCOUNTS PAYABLES
    Dim rsDISBURSEMENT7                           As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJType                                    As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xVENDORCODE                               As String
    Dim xDUEDATE                                  As String
    Dim xACCT_CODE                                As String

    Dim xINVOICEAMOUNT                            As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    xAMOUNTPAID = 0
    Set rsDISBURSEMENT7 = New ADODB.Recordset
    rsDISBURSEMENT7.Open "SELECT HD.VOUCHERNO AS HD_VOUCHERNO,HD.JTYPE AS HD_JTYPE,HD.JDATE AS HD_JDATE,HD.STATUS AS HD_STATUS,HD.INVOICETYPE AS HD_INV_TYPE,HD.INVOICEDATE AS HD_INV_DATE, HD.VENDORCODE AS HD_VEN_CODE, " & _
                         "HD.INVOICEAMT AS HD_INV_AMT,HD.AMOUNTTOPAY AS HD_AMT_TO_PAY,HD.AMOUNTPAID AS HD_AMT_PAID,DET.DEBIT AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.ACCT_CODE AS DET_ACCT_CODE " & _
                         "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE (HD.JTYPE = 'CDJ') AND LEFT(DET.ACCT_CODE,5) IN ('21-02') AND HD.STATUS = 'P' " & _
                         "AND HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND DET.CREDIT > 0 AND HD.JDATE < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset
    Label11.Caption = " Processing CDJ... Please Wait.."

    If rsDISBURSEMENT7.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsDISBURSEMENT7.RecordCount

    If Not rsDISBURSEMENT7.EOF And Not rsDISBURSEMENT7.BOF Then
        Do While Not rsDISBURSEMENT7.EOF
            xVOUCHERNO = N2Str2Null(rsDISBURSEMENT7!HD_VOUCHERNO)
            xJType = N2Str2Null(rsDISBURSEMENT7!HD_JTYPE)
            xJdate = N2Date2Null(rsDISBURSEMENT7!HD_JDATE)
            xSTATUS = N2Str2Null(rsDISBURSEMENT7!HD_STATUS)
            xInvoicedate = N2Date2Null(rsDISBURSEMENT7!HD_INV_DATE)
            xVENDORCODE = N2Str2Null(rsDISBURSEMENT7!HD_VEN_CODE)
            xINVOICETYPE = N2Str2Null(rsDISBURSEMENT7!HD_INV_TYPE)
            xINVOICEAMOUNT = NumericVal(rsDISBURSEMENT7!HD_INV_AMT)
            xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT7!HD_AMT_TO_PAY) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJType, xVENDORCODE)), 2)
            '+ CDJAmount2Pay(xVOUCHERNO, xJTYPE, xVENDORCODE, xACCT_CODE)
            xAMOUNTPAID = GET_DRCT_AMOUNT(xVOUCHERNO, xJType)
            'NumericVal(rsDISBURSEMENT7!HD_AMT_PAID)
            xdebit = NumericVal(rsDISBURSEMENT7!DET_DEBIT)
            xcredit = NumericVal(rsDISBURSEMENT7!DET_CREDIT)
            xDUEDATE = N2Date2Null(rsDISBURSEMENT7!HD_DUEDATE)
            xACCT_CODE = N2Str2Null(rsDISBURSEMENT7!DET_ACCT_CODE)

            gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & "," & xINVOICETYPE & "," & xInvoicedate & "," & xINVOICEAMOUNT & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCT_CODE & ", " & xDUEDATE & ")"

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsDISBURSEMENT7!HD_VOUCHERNO)
            DoEvents

            rsDISBURSEMENT7.MoveNext
        Loop
    End If
    Set rsDISBURSEMENT7 = Nothing
End Sub

Sub DISBURSEMENT8()
'DESCRIPTION: THIS SELECT ALL DIRECT DISBURSEMENT WHICH IS NO LINK TO THE ACCOUNTS PAYABLES
    Dim rsDISBURSEMENT8                           As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJType                                    As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xVENDORCODE                               As String
    Dim xDUEDATE                                  As String
    Dim xACCT_CODE                                As String

    Dim xINVOICEAMOUNT                            As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    xAMOUNTPAID = 0
    Set rsDISBURSEMENT8 = New ADODB.Recordset
    rsDISBURSEMENT8.Open "SELECT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE " & _
                         "FROM (SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
                         "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE " & _
                         "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN " & _
                         "(SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND DET.ACCT_CODE = ('21-02004-00') AND HD.status = 'P' AND HD.jdate < = '" & dtAsOF & "')X " & _
                         "WHERE X.AP_VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE JTYPE='APJ') ORDER BY AP_VOUCHERNO", gconDMIS, adOpenKeyset
    Label11.Caption = " Processing CDJ... Please Wait.."

    If rsDISBURSEMENT8.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsDISBURSEMENT8.RecordCount

    If Not rsDISBURSEMENT8.EOF And Not rsDISBURSEMENT8.BOF Then
        Do While Not rsDISBURSEMENT8.EOF
            xVOUCHERNO = N2Str2Null(rsDISBURSEMENT8!HD_VOUCHERNO)
            xJType = N2Str2Null(rsDISBURSEMENT8!HD_JTYPE)
            xJdate = N2Date2Null(rsDISBURSEMENT8!HD_JDATE)
            xSTATUS = N2Str2Null(rsDISBURSEMENT8!HD_STATUS)
            xInvoicedate = N2Date2Null(rsDISBURSEMENT8!HD_INV_DATE)
            xVENDORCODE = N2Str2Null(rsDISBURSEMENT8!HD_VEN_CODE)
            xINVOICETYPE = N2Str2Null(rsDISBURSEMENT8!HD_INV_TYPE)
            xINVOICEAMOUNT = NumericVal(rsDISBURSEMENT8!HD_INV_AMT)
            xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT8!HD_AMT_TO_PAY) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJType, xVENDORCODE)), 2)
            '+ CDJAmount2Pay(xVOUCHERNO, xJTYPE, xVENDORCODE, xACCT_CODE)
            xAMOUNTPAID = GET_DRCT_AMOUNT(xVOUCHERNO, xJType)
            'NumericVal(rsDISBURSEMENT8!HD_AMT_PAID)
            xdebit = NumericVal(rsDISBURSEMENT8!DET_DEBIT)
            xcredit = NumericVal(rsDISBURSEMENT8!DET_CREDIT)
            xDUEDATE = N2Date2Null(rsDISBURSEMENT8!HD_DUEDATE)
            xACCT_CODE = N2Str2Null(rsDISBURSEMENT8!DET_ACCT_CODE)

            gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & "," & xINVOICETYPE & "," & xInvoicedate & "," & xINVOICEAMOUNT & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCT_CODE & ", " & xDUEDATE & ")"

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsDISBURSEMENT8!HD_VOUCHERNO)
            DoEvents

            rsDISBURSEMENT8.MoveNext
        Loop
    End If
    Set rsDISBURSEMENT8 = Nothing
End Sub

Sub DISBURSEMENT9()
'DESCRIPTION: THIS SELECT ALL DIRECT DISBURSEMENT WHICH IS NO LINK TO THE ACCOUNTS PAYABLES
    Dim rsDISBURSEMENT9                           As ADODB.Recordset
    Dim rsWRONGENTRY                              As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJType                                    As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xVENDORCODE                               As String
    Dim xDUEDATE                                  As String
    Dim xACCT_CODE                                As String

    Dim xINVOICEAMOUNT                            As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    xAMOUNTPAID = 0
    Set rsDISBURSEMENT9 = New ADODB.Recordset
    rsDISBURSEMENT9.Open "SELECT DISTINCT X.AP_VOUCHERNO,X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_CREDIT,X.HD_DUEDATE " & _
                         "FROM (SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VOUCHERNO AS HD_VOUCHERNO,HD.JTYPE AS HD_JTYPE,HD.JDATE AS HD_JDATE,HD.STATUS AS HD_STATUS,HD.INVOICETYPE AS HD_INV_TYPE,HD.INVOICEDATE AS HD_INV_DATE, " & _
                         "HD.VENDORCODE AS HD_VEN_CODE, HD.INVOICEAMT AS HD_INV_AMT,HD.AMOUNTTOPAY AS HD_AMT_TO_PAY,HD.AMOUNTPAID AS HD_AMT_PAID,DET.DEBIT AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE, " & _
                         "DET.ACCT_CODE AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO " & _
                         "AND HD.JTYPE=CV.CV_JTYPE WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND DET.ACCT_CODE = ('21-01002-00') AND HD.STATUS = 'P' AND HD.JDATE < = '" & dtAsOF & "')X " & _
                         "WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE ACCT_CODE = '21-01004-00')", gconDMIS, adOpenKeyset
    Label11.Caption = " Processing CDJ... Please Wait.."

    If rsDISBURSEMENT9.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsDISBURSEMENT9.RecordCount

    If Not rsDISBURSEMENT9.EOF And Not rsDISBURSEMENT9.BOF Then
        Do While Not rsDISBURSEMENT9.EOF

            Set rsWRONGENTRY = New ADODB.Recordset
            rsWRONGENTRY.Open "SELECT VOUCHERNO,JTYPE,ACCT_CODE,DEBIT,CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO= " & N2Str2Null(rsDISBURSEMENT9!AP_VOUCHERNO) & " AND JTYPE='APJ' AND ACCT_CODE = '21-01004-00'", gconDMIS, adOpenKeyset
            If Not rsWRONGENTRY.EOF And Not rsWRONGENTRY.BOF Then
                Do While Not rsWRONGENTRY.EOF
                    ' MsgBox "Q" + N2Str2Null(rsWRONGENTRY!CREDIT)


                    xVOUCHERNO = N2Str2Null(rsDISBURSEMENT9!HD_VOUCHERNO)
                    xJType = N2Str2Null(rsDISBURSEMENT9!HD_JTYPE)
                    xJdate = N2Date2Null(rsDISBURSEMENT9!HD_JDATE)
                    xSTATUS = N2Str2Null(rsDISBURSEMENT9!HD_STATUS)
                    xInvoicedate = N2Date2Null(rsDISBURSEMENT9!HD_INV_DATE)
                    xVENDORCODE = N2Str2Null(rsDISBURSEMENT9!HD_VEN_CODE)
                    xINVOICETYPE = N2Str2Null(rsDISBURSEMENT9!HD_INV_TYPE)
                    xINVOICEAMOUNT = NumericVal(rsDISBURSEMENT9!HD_INV_AMT)
                    xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT9!HD_AMT_TO_PAY) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJType, xVENDORCODE)), 2)
                    '+ CDJAmount2Pay(xVOUCHERNO, xJTYPE, xVENDORCODE, xACCT_CODE)
                    xAMOUNTPAID = GET_DRCT_AMOUNT(xVOUCHERNO, xJType)
                    'NumericVal(rsDISBURSEMENT9!HD_AMT_PAID)
                    xdebit = NumericVal(rsWRONGENTRY!CREDIT)
                    xcredit = NumericVal(rsDISBURSEMENT9!DET_CREDIT)
                    xDUEDATE = N2Date2Null(rsDISBURSEMENT9!HD_DUEDATE)
                    xACCT_CODE = N2Str2Null(rsDISBURSEMENT9!DET_ACCT_CODE)

                    gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                                     "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & "," & xINVOICETYPE & "," & xInvoicedate & "," & xINVOICEAMOUNT & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCT_CODE & ", " & xDUEDATE & ")"

                    rsWRONGENTRY.MoveNext
                Loop
            End If

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsDISBURSEMENT9!HD_VOUCHERNO)
            DoEvents

            rsDISBURSEMENT9.MoveNext
        Loop
    End If
    Set rsDISBURSEMENT9 = Nothing
End Sub

Sub DISBURSEMENT10()
    Dim rsDISBURSEMENT10                          As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJType                                    As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xVENDORCODE                               As String
    Dim xDUEDATE                                  As String
    Dim xACCT_CODE                                As String

    Dim xINVOICEAMOUNT                            As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    xAMOUNTPAID = 0
    'DESCRIPTION: WITH DETAIL BUT DETAIL
    Set rsDISBURSEMENT10 = New ADODB.Recordset
    rsDISBURSEMENT10.Open "SELECT DISTINCT X.DET_ACCT_CODE,X.HD_VOUCHERNO,X.HD_JTYPE,X.HD_JDATE,X.HD_STATUS,X.HD_INV_DATE,X.HD_VEN_CODE,X.HD_INV_TYPE,X.HD_INV_AMT,X.HD_AMT_TO_PAY,X.DET_DEBIT,X.DET_CREDIT,X.HD_DUEDATE " & _
                          "FROM (SELECT CV.PV_VOUCHERNO AS AP_VOUCHERNO,HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
                          "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE " & _
                          "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CV_DETAIL CV ON HD.VOUCHERNO=CV.VOUCHERNO AND HD.JTYPE=CV.CV_JTYPE " & _
                          "WHERE HD.VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND DET.ACCT_CODE = ('21-01002-00') AND HD.status = 'P' AND HD.jdate < = '" & dtAsOF & "')X " & _
                          "WHERE X.AP_VOUCHERNO IN (SELECT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE JTYPE='APJ' AND ACCT_CODE='21-02004-00')", gconDMIS, adOpenKeyset
    Label11.Caption = " Processing CDJ... Please Wait.."

    If rsDISBURSEMENT10.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsDISBURSEMENT10.RecordCount

    If Not rsDISBURSEMENT10.EOF And Not rsDISBURSEMENT10.BOF Then
        Do While Not rsDISBURSEMENT10.EOF
            xVOUCHERNO = N2Str2Null(rsDISBURSEMENT10!HD_VOUCHERNO)
            xJType = N2Str2Null(rsDISBURSEMENT10!HD_JTYPE)
            xJdate = N2Date2Null(rsDISBURSEMENT10!HD_JDATE)
            xSTATUS = N2Str2Null(rsDISBURSEMENT10!HD_STATUS)
            xInvoicedate = N2Date2Null(rsDISBURSEMENT10!HD_INV_DATE)
            xVENDORCODE = N2Str2Null(rsDISBURSEMENT10!HD_VEN_CODE)
            xINVOICETYPE = N2Str2Null(rsDISBURSEMENT10!HD_INV_TYPE)
            xINVOICEAMOUNT = NumericVal(rsDISBURSEMENT10!HD_INV_AMT)
            xAMOUNT2PAY = Round((NumericVal(rsDISBURSEMENT10!HD_AMT_TO_PAY) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJType, xVENDORCODE)), 2)
            '+ CDJAmount2Pay(xVOUCHERNO, xJTYPE, xVENDORCODE, xACCT_CODE)
            xAMOUNTPAID = GET_DRCT_AMOUNT(xVOUCHERNO, xJType)
            'NumericVal(rsDISBURSEMENT10!HD_AMT_PAID)
            xdebit = NumericVal(rsDISBURSEMENT10!DET_DEBIT)
            xcredit = NumericVal(rsDISBURSEMENT10!DET_CREDIT)
            xDUEDATE = N2Date2Null(rsDISBURSEMENT10!HD_DUEDATE)
            xACCT_CODE = N2Str2Null(rsDISBURSEMENT10!DET_ACCT_CODE)

            gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & "," & xINVOICETYPE & "," & xInvoicedate & "," & xINVOICEAMOUNT & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCT_CODE & ", " & xDUEDATE & ")"

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsDISBURSEMENT10!HD_VOUCHERNO)
            DoEvents

            rsDISBURSEMENT10.MoveNext
        Loop
    End If
    Set rsDISBURSEMENT10 = Nothing
End Sub

Sub DIRECT_DSBRSMENT()
'DESCRIPTION: THIS SELECT ALL DIRECT DISBURSEMENT WHICH IS NO LINK TO THE ACCOUNTS PAYABLES / NO DETAIL
    Dim rsDIRECT_DSBRSMENT                        As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJType                                    As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim xVENDORCODE                               As String
    Dim xDUEDATE                                  As String
    Dim xACCOUNTCODE                              As String
    Dim xACCT_CODE                                As String
    Dim xINVOICEAMOUNT                            As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xAMOUNTPAID                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    xAMOUNTPAID = 0
    Set rsDIRECT_DSBRSMENT = New ADODB.Recordset
    'FOR DEBUGGNG ONLY
    'rsDIRECT_DSBRSMENT.Open "SELECT HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
     "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, " & _
     "HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
     "WHERE HD.JTYPE IN ('CDJ','CCM') AND DET.ACCT_CODE='21-01008-00' AND HD.VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND HD.status = 'P' and HD.jdate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset
    '==================
    rsDIRECT_DSBRSMENT.Open "SELECT HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE, " & _
                            "HD.VendorCode AS HD_VEN_CODE, HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT,DET.CREDIT AS DET_CREDIT, " & _
                            "HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE   FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET  ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                            "WHERE HD.JTYPE IN ('CDJ','CCM') AND LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-07') AND HD.VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL) AND HD.status = 'P' and HD.jdate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset

    Label11.Caption = " Processing CDJ... Please Wait.."

    If rsDIRECT_DSBRSMENT.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsDIRECT_DSBRSMENT.RecordCount

    If Not rsDIRECT_DSBRSMENT.EOF And Not rsDIRECT_DSBRSMENT.BOF Then
        Do While Not rsDIRECT_DSBRSMENT.EOF
            xVOUCHERNO = N2Str2Null(rsDIRECT_DSBRSMENT!HD_VOUCHERNO)
            xJType = N2Str2Null(rsDIRECT_DSBRSMENT!HD_JTYPE)
            xJdate = N2Date2Null(rsDIRECT_DSBRSMENT!HD_JDATE)
            xSTATUS = N2Str2Null(rsDIRECT_DSBRSMENT!HD_STATUS)
            xInvoicedate = N2Date2Null(rsDIRECT_DSBRSMENT!HD_INV_DATE)
            xVENDORCODE = N2Str2Null(rsDIRECT_DSBRSMENT!HD_VEN_CODE)
            xINVOICETYPE = N2Str2Null(rsDIRECT_DSBRSMENT!HD_INV_TYPE)
            xINVOICEAMOUNT = NumericVal(rsDIRECT_DSBRSMENT!HD_INV_AMT)
            xACCT_CODE = N2Str2Null(rsDIRECT_DSBRSMENT!DET_ACCT_CODE)
            'If xJTYPE = "'CDJ'" Then
            'If rsDIRECT_DSBRSMENT!HD_VOUCHERNO = "002362" Then Stop
            xAMOUNT2PAY = Round((NumericVal(rsDIRECT_DSBRSMENT!HD_AMT_TO_PAY) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJType, xVENDORCODE) + CDJAmount2Pay(xVOUCHERNO, xJType, xVENDORCODE, xACCT_CODE)), 2)
            'Else
            '    xAMOUNT2PAY = Round((NumericVal(rsDIRECT_DSBRSMENT!DET_CREDIT) + GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO, xJTYPE, xVENDORCODE)), 2)
            'End If
            xAMOUNTPAID = GET_DRCT_AMOUNT(xVOUCHERNO, xJType)
            'NumericVal(rsDIRECT_DSBRSMENT!HD_AMT_PAID)
            xdebit = NumericVal(rsDIRECT_DSBRSMENT!DET_DEBIT)
            xcredit = NumericVal(rsDIRECT_DSBRSMENT!DET_CREDIT)
            xDUEDATE = N2Date2Null(rsDIRECT_DSBRSMENT!HD_DUEDATE)
            xACCOUNTCODE = N2Str2Null(rsDIRECT_DSBRSMENT!DET_ACCT_CODE)

            gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & "," & xINVOICETYPE & "," & xInvoicedate & "," & xINVOICEAMOUNT & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCOUNTCODE & ", " & xDUEDATE & ")"

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsDIRECT_DSBRSMENT!HD_VOUCHERNO)
            DoEvents

            rsDIRECT_DSBRSMENT.MoveNext
        Loop
    End If
    Set rsDIRECT_DSBRSMENT = Nothing
End Sub

Function CDJAmount2Pay(xVOUCHERNO As String, xJType As String, xVENDORCODE As String, xACCT_CODE As String) As Double
    Dim rsCDJAmount2Pay                           As ADODB.Recordset
    Dim sum_CDJAmount2Pay                         As Double
    Set rsCDJAmount2Pay = New ADODB.Recordset
    '21-01008-00
    '21-02009-00
    rsCDJAmount2Pay.Open "SELECT DET.CREDIT AS CDJ_CREDIT FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE where HD.VoucherNo = " & xVOUCHERNO & " and HD.JType = 'CDJ' and HD.VendorCode =" & xVENDORCODE & " and DET.Acct_Code = " & xACCT_CODE & " AND DET.CREDIT > 0", gconDMIS, adOpenKeyset
    If rsCDJAmount2Pay.RecordCount = 0 Then Exit Function
    If xACCT_CODE = "'21-01008-00'" Then
        If Not rsCDJAmount2Pay.EOF And Not rsCDJAmount2Pay.BOF Then
            Do While Not rsCDJAmount2Pay.EOF
                sum_CDJAmount2Pay = sum_CDJAmount2Pay + rsCDJAmount2Pay!CDJ_CREDIT
                rsCDJAmount2Pay.MoveNext
            Loop
            CDJAmount2Pay = NumericVal(sum_CDJAmount2Pay)
        End If
    End If
    Set rsCDJAmount2Pay = Nothing
End Function

Function Amount2Pay2(xVOUCHERNO As String, xJType As String, xVENDORCODE As String, xACCT_CODE As String) As Double
    Dim rsAmount2Pay2                             As ADODB.Recordset
    Dim sum_Amount2Pay2                           As Double
    Set rsAmount2Pay2 = New ADODB.Recordset
    rsAmount2Pay2.Open "SELECT DET.CREDIT AS CDJ_CREDIT FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE where HD.VoucherNo = '" & xVOUCHERNO & "' and HD.JType = 'CDJ' and HD.VendorCode ='" & xVENDORCODE & "' and DET.Acct_Code = '" & xACCT_CODE & "' AND DET.CREDIT > 0", gconDMIS, adOpenKeyset
    '21-02009-00
    If rsAmount2Pay2.RecordCount = 0 Then Exit Function
    If Not rsAmount2Pay2.EOF And Not rsAmount2Pay2.BOF Then
        If xACCT_CODE = "21-02009-00" Then
            Do While Not rsAmount2Pay2.EOF
                sum_Amount2Pay2 = sum_Amount2Pay2 + rsAmount2Pay2!CDJ_CREDIT
                rsAmount2Pay2.MoveNext
            Loop
            Amount2Pay2 = NumericVal(sum_Amount2Pay2)
        ElseIf xACCT_CODE = "21-02000-00" Then
            Do While Not rsAmount2Pay2.EOF
                sum_Amount2Pay2 = sum_Amount2Pay2 + rsAmount2Pay2!CDJ_CREDIT
                rsAmount2Pay2.MoveNext
            Loop
            Amount2Pay2 = NumericVal(sum_Amount2Pay2)
        ElseIf xACCT_CODE = "21-07006-00" Then
            Do While Not rsAmount2Pay2.EOF
                sum_Amount2Pay2 = sum_Amount2Pay2 + rsAmount2Pay2!CDJ_CREDIT
                rsAmount2Pay2.MoveNext
            Loop
            Amount2Pay2 = NumericVal(sum_Amount2Pay2)
        End If
    End If
    Set rsAmount2Pay2 = Nothing
End Function

Function GET_CDJNO(xAP_VOUCHER As String, xJType As String) As String
'DESCRIPTION: GET THE CDJ NO
    Dim rsGET_CDJNO                               As ADODB.Recordset
    Set rsGET_CDJNO = New ADODB.Recordset
    rsGET_CDJNO.Open "Select VOUCHERNO from AMIS_CV_DETAIL WHERE  JTYPE = '" & xJType & "' AND PV_VOUCHERNO = '" & xAP_VOUCHER & "'", gconDMIS, adOpenKeyset
    If Not rsGET_CDJNO.EOF And Not rsGET_CDJNO.BOF Then
        GET_CDJNO = Null2String(rsGET_CDJNO!VOUCHERNO)
    Else
        'NOT FOUND
    End If
    Set rsGET_CDJNO = Nothing
End Function

Function GET_VEN_NAME(xVENDORCODE As String) As String
'DESCRIPTION: GET THE VEDOR NAME IN ALL_VENDOR_TABLE
    Dim rsGET_VEN_NAME                            As ADODB.Recordset
    Set rsGET_VEN_NAME = New ADODB.Recordset
    rsGET_VEN_NAME.Open "Select nameofvendor from all_vendor_table where code = " & xVENDORCODE & "", gconDMIS, adOpenKeyset
    If Not rsGET_VEN_NAME.EOF And Not rsGET_VEN_NAME.BOF Then
        GET_VEN_NAME = Null2String(rsGET_VEN_NAME!nameofvendor)
    Else
        'VENDOR NAME NOT FOUND
    End If
    Set rsGET_VEN_NAME = Nothing
End Function

Function GET_DRCT_AMOUNT(xVOUCHERNO As String, xJType As String) As Double
'DESCRIPTION: GET THE DEBIT AMOUNT OF THE DIRECT DISBURSEMENT AS AN ADVANCE PAYMENT
    Dim rsGET_DRCT_AMOUNT                         As ADODB.Recordset
    Set rsGET_DRCT_AMOUNT = New ADODB.Recordset
    rsGET_DRCT_AMOUNT.Open "SELECT ROUND(SUM(DEBIT),2) AS DEBIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & xVOUCHERNO & " AND JTYPE = " & xJType & " AND STATUS = 'P' AND LEFT(ACCT_CODE,5) IN('21-01','21-02','21-07') ", gconDMIS, adOpenKeyset
    If Not rsGET_DRCT_AMOUNT.EOF And Not rsGET_DRCT_AMOUNT.BOF Then
        GET_DRCT_AMOUNT = NumericVal(rsGET_DRCT_AMOUNT!DEBIT)
    Else
        GET_DRCT_AMOUNT = 0
    End If
    Set rsGET_DRCT_AMOUNT = Nothing
End Function

Function GET_ADJ_DRCT_AMT2PAY(xVOUCHERNO As String, xJType As String, xCUSCDE As String) As Double
'DESCRIPTION: GET THE ADJUSTMENT AS AMOUNT TO PAY TO THE DIRECT DISBURSEMENT
    Dim rsGET_ADJ_DRCT_AMT2PAY                    As ADODB.Recordset
    Set rsGET_ADJ_DRCT_AMT2PAY = New ADODB.Recordset
    rsGET_ADJ_DRCT_AMT2PAY.Open "SELECT  ROUND(SUM(DET.CREDIT),2) AS DET_CREDIT " & _
                                "FROM AMIS_JOURNAL_DET DET " & _
                                "WHERE DET.ADJ_JTYPE  = 'CDJ' AND LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-07') AND DET.VOUCHERNO = " & xVOUCHERNO & " AND RIGHT(ENTITY,6) = " & xCUSCDE & " AND DET.STATUS = 'P' AND DET.JTYPE = 'GJ'", gconDMIS, adOpenKeyset
    If Not rsGET_ADJ_DRCT_AMT2PAY.EOF And Not rsGET_ADJ_DRCT_AMT2PAY.BOF Then
        GET_ADJ_DRCT_AMT2PAY = NumericVal(rsGET_ADJ_DRCT_AMT2PAY!DET_CREDIT)
    End If
    Set rsGET_ADJ_DRCT_AMT2PAY = Nothing
End Function

Sub TRANSFER_AP_JOURNAL()
'DESCRIPTION: THIS GET ALL THE ACCOUNTS PAYABLES
    Dim rsTRANS                                   As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJdate                                    As String
    Dim xSTATUS                                   As String
    Dim xVENDORCODE                               As String
    Dim xInvoicedate                              As String
    Dim xINVOICENO                                As String
    Dim xINVOICETYPE                              As String
    Dim xACCT_CODE                                As String
    Dim xDUEDATE                                  As String
    Dim xJType                                    As String
    Dim xInvoiceAmnt                              As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    Dim xAMOUNT2PAY                               As Double

    Set rsTRANS = New ADODB.Recordset

    Timer1.Enabled = True
    gconDMIS.Execute "TRUNCATE TABLE AMIS_AP_HD"
    'FOR DEBUGGING ONLY
    'rsTRANS.Open "SELECT  HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.InvoiceNo as HD_INV_NO,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE,HD.VendorCode AS HD_VEN_CODE, " & _
     "HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT, DET.CREDIT AS DET_CREDIT,HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE " & _
     "FROM  AMIS_Journal_HD HD INNER JOIN AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.jtype = DET.jtype " & _
     "WHERE HD.JType IN ('APJ','VPJ','CRJ','GJ','VDJ','VCJ') AND DET.Acct_Code = '21-01008-00' and HD.status='P' and HD.jdate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset

    '=========================
    If COMPANY_CODE = "HGC" Then
        gconDMIS.Execute ("DELETE FROM AMIS_JOURNAL_HD WHERE ID = 10505")
        gconDMIS.Execute ("DELETE FROM AMIS_JOURNAL_HD WHERE ID = 11332")
        gconDMIS.Execute ("DELETE FROM AMIS_JOURNAL_HD WHERE ID = 10634")
    End If
    rsTRANS.Open "SELECT  HD.VoucherNo AS HD_VOUCHERNO,HD.JType AS HD_JTYPE,HD.jdate AS HD_JDATE,HD.status AS HD_STATUS,HD.InvoiceNo as HD_INV_NO,HD.invoicetype AS HD_INV_TYPE,HD.invoicedate AS HD_INV_DATE,HD.VendorCode AS HD_VEN_CODE, " & _
                 "HD.invoiceamt AS HD_INV_AMT,HD.amounttopay AS HD_AMT_TO_PAY,HD.amountpaid AS HD_AMT_PAID,DET.Debit AS DET_DEBIT, DET.CREDIT AS DET_CREDIT,HD.DUEDATE AS HD_DUEDATE,DET.Acct_Code AS DET_ACCT_CODE " & _
                 "FROM  AMIS_Journal_HD HD INNER JOIN AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.jtype = DET.jtype " & _
                 "WHERE HD.JType IN ('APJ','VPJ','CRJ','GJ','VDJ','VCJ') AND LEFT(DET.Acct_Code,5) IN ('21-01','21-02','21-07') and HD.status='P' and HD.jdate < = '" & dtAsOF & "'", gconDMIS, adOpenKeyset
    Label11.Caption = "Processing AP... Please Wait.."

    If rsTRANS.RecordCount = 0 Then Exit Sub

    ProgressBar2.Value = 0
    ProgressBar2.Max = rsTRANS.RecordCount

    If Not rsTRANS.EOF And Not rsTRANS.BOF Then
        Do While Not rsTRANS.EOF
            xVOUCHERNO = N2Str2Null(rsTRANS!HD_VOUCHERNO)
            xJdate = N2Date2Null(rsTRANS!HD_JDATE)
            xJType = N2Str2Null(rsTRANS!HD_JTYPE)
            xSTATUS = N2Str2Null(rsTRANS!HD_STATUS)
            xVENDORCODE = N2Str2Null(rsTRANS!HD_VEN_CODE)
            xINVOICENO = N2Str2Null(rsTRANS!HD_INV_NO)
            xINVOICETYPE = N2Str2Null(rsTRANS!HD_INV_TYPE)
            xInvoicedate = N2Date2Null(rsTRANS!HD_INV_DATE)
            xInvoiceAmnt = NumericVal(rsTRANS!HD_INV_AMT)
            xdebit = NumericVal(rsTRANS!DET_DEBIT)
            xcredit = NumericVal(rsTRANS!DET_CREDIT)
            xACCT_CODE = N2Str2Null(rsTRANS!DET_ACCT_CODE)
            xDUEDATE = N2Date2Null(rsTRANS!HD_DUEDATE)
            If xJType = "'VPJ'" Then
                xAMOUNT2PAY = Round((NumericVal(rsTRANS!HD_AMT_TO_PAY) + COM_AP_ADJ_CREDIT(xVENDORCODE, xINVOICENO, xINVOICETYPE, xJType)), 2)
            Else
                xAMOUNT2PAY = Round((COM_AP_AMT_TO_PAY(xVOUCHERNO, xJType, xACCT_CODE) + COM_AP_ADJ_CREDIT(xVENDORCODE, xINVOICENO, xINVOICETYPE, xJType)), 2)
            End If
            gconDMIS.Execute "Insert into AMIS_AP_HD (VOUCHERNO,JDATE,JTYPE,STATUS,VENDOR_CODE,INVOICENO,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,DUEDATE)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & ", " & xJType & "," & xSTATUS & "," & xVENDORCODE & ", " & xINVOICENO & "," & xINVOICETYPE & "," & xInvoicedate & "," & xInvoiceAmnt & "," & xdebit & "," & xcredit & "," & xAMOUNT2PAY & ", " & xACCT_CODE & ", " & xDUEDATE & ")"
            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round(((ProgressBar2.Value / ProgressBar2.Max) * 100), 0) & "%"
            Label12.Caption = Null2String(rsTRANS!HD_VOUCHERNO)
            DoEvents
            rsTRANS.MoveNext
        Loop
    End If
    Set rsTRANS = Nothing
End Sub

Function VAL_VEN_CODE(xAP_VOUCHERnO As String, xJType As String, xVEN_CODE As String)
'DESCRIPTION: VALIDATE THE VENDOR CODE OF AP AND CDJ
    Dim rsVAL_VEN_CODE                            As ADODB.Recordset
    '    Dim rsVAL_VEN_CODE2 As ADODB.Recordset
    Dim RSCDJ                                     As ADODB.Recordset
    '    Dim RSCDJ2 As ADODB.Recordset

    '    If zJTYPE <> "NULL" Then
    Set rsVAL_VEN_CODE = New ADODB.Recordset
    rsVAL_VEN_CODE.Open "Select VOUCHERNO,JTYPE FROM AMIS_CV_DETAIL WHERE PV_VOUCHERNO = '" & xAP_VOUCHERnO & "' AND JTYPE = '" & xJType & "'", gconDMIS, adOpenKeyset
    If Not rsVAL_VEN_CODE.EOF And Not rsVAL_VEN_CODE.BOF Then
        '            zJTYPE = N2Str2Null(rsVAL_VEN_CODE!jtype)
        Set RSCDJ = New ADODB.Recordset
        RSCDJ.Open "Select VendorCOde,PAYTYPE from AMIS_JOURNAL_HD WHERE JTYPE = 'CDJ' AND VOUCHERNO = '" & Null2String(rsVAL_VEN_CODE!VOUCHERNO) & "' AND  VENDORCODE = '" & xVEN_CODE & "'", gconDMIS, adOpenKeyset
        If Not RSCDJ.EOF And Not RSCDJ.BOF Then
            xPAYMENT_TYPE = Null2String(RSCDJ!paytype)
            xREMARKS = N2Str2Null("")
        Else
            xREMARKS = N2Str2Null("WRONG CODE")
        End If
    End If

    '    Else
    '        Set rsVAL_VEN_CODE2 = New ADODB.Recordset
    '        rsVAL_VEN_CODE2.Open "Select VOUCHERNO FROM AMIS_CV_DETAIL WHERE PV_VOUCHERNO = '" & xAP_VOUCHERnO & "'", gconDMIS, adOpenKeyset
    '        If Not rsVAL_VEN_CODE2.EOF And Not rsVAL_VEN_CODE2.BOF Then
    '            Set RSCDJ2 = New ADODB.Recordset
    '            RSCDJ2.Open "Select VendorCOde,PAYTYPE from AMIS_JOURNAL_HD WHERE JTYPE = 'CDJ' AND VOUCHERNO = '" & Null2String(rsVAL_VEN_CODE2!VOUCHERNO) & "' AND  VENDORCODE = '" & xVEN_CODE & "'", gconDMIS, adOpenKeyset
    '            If Not RSCDJ2.EOF And Not RSCDJ2.BOF Then
    '                xPAYMENT_TYPE = Null2String(RSCDJ2!paytype)
    '                xREMARKS = N2Str2Null("")
    '            Else
    '                xREMARKS = N2Str2Null("WRONG CODE")
    '            End If
    '        End If
    '        Set RSCDJ2 = Nothing
    '    End If
    Set rsVAL_VEN_CODE = Nothing
    '    Set rsVAL_VEN_CODE2 = Nothing
End Function

Private Sub Command1_Click()
'DESCRIPTION: PRINTING OF GENERATED ACCOUNT RECEIVABLE REPORTS (AP AGING REPORT)
    rptAP_Aging.WindowShowSearchBtn = True
    rptAP_Aging.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    rptAP_Aging.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptAP_Aging.WindowTitle = "Accounts Payable Aging Report  AS OF: " & dtAsOF
    rptAP_Aging.ReportTitle = "Accounts Payable Aging Report AS OF: " & dtAsOF
    PrintSQLReport rptAP_Aging, AMIS_REPORT_PATH & "DueReports\AP_AgingReport.Rpt", "", DMIS_REPORT_Connection, 1
    LogAudit "V", "ACCOUNTS PAYABLE AGING REPORT", "As of: " & dtAsOF
End Sub

Private Sub Command2_Click()
'DESCRIPTION: PRINTING OF GENERATED ACCOUNT RECEIVABLE REPORTS (AP AGING REPORT)
    rptAP_Aging.WindowShowSearchBtn = True
    rptAP_Aging.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    rptAP_Aging.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptAP_Aging.WindowTitle = "Accounts Payable Schedule Report  AS OF: " & dtAsOF
    rptAP_Aging.ReportTitle = "Accounts Payable Schedule Report AS OF: " & dtAsOF
    PrintSQLReport rptAP_Aging, AMIS_REPORT_PATH & "DueReports\AP_ScheduleReport.Rpt", "", DMIS_REPORT_Connection, 1
    LogAudit "V", "ACCOUNTS PAYABLE AGING REPORT", "As of: " & dtAsOF
End Sub

Private Sub Command3_Click()
    If COMPANY_CODE = "HGC" Then

    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Timer1.Enabled = False
    dtAsOF = LOGDATE
End Sub

Private Sub Timer1_Timer()
    If Label11.Caption <> "" Then
        If Picture3.Visible = True Then
            Picture3.Visible = False
        Else
            Picture3.Visible = True
        End If
    End If
End Sub



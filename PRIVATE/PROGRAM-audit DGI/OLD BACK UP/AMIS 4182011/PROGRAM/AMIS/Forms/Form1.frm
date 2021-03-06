VERSION 5.00
Begin VB.Form frm_TOOLS_ARRefresher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AMIS Fixer"
   ClientHeight    =   675
   ClientLeft      =   315
   ClientTop       =   765
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   2655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "AR Refresh"
      Height          =   555
      Left            =   3720
      TabIndex        =   10
      Top             =   0
      Width           =   2445
   End
   Begin VB.CommandButton cmdAPRefresh 
      Caption         =   "AP Refresh"
      Height          =   555
      Left            =   90
      TabIndex        =   9
      Top             =   5190
      Width           =   2445
   End
   Begin VB.CommandButton cmdRefreshInvTerm 
      Caption         =   "Refresh Invoice Term Type"
      Height          =   555
      Left            =   90
      TabIndex        =   8
      Top             =   4650
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Account with Invalid Title Code"
      Height          =   555
      Left            =   90
      TabIndex        =   7
      Top             =   4110
      Width           =   2445
   End
   Begin VB.CommandButton cmdInvalidAcctCode 
      Caption         =   "Invalid Account Code"
      Height          =   555
      Left            =   90
      TabIndex        =   6
      Top             =   3570
      Width           =   2445
   End
   Begin VB.CommandButton checkWOSTotals 
      Caption         =   "Check Acct W/O Acct Totals"
      Height          =   555
      Left            =   90
      TabIndex        =   5
      Top             =   3030
      Width           =   2445
   End
   Begin VB.CommandButton cmdShowErrorQuery 
      Caption         =   "Show Error Query"
      Height          =   555
      Left            =   90
      TabIndex        =   4
      Top             =   2490
      Width           =   2445
   End
   Begin VB.CommandButton cmdCheckErrorTrans 
      Caption         =   "Check Error Journals"
      Height          =   555
      Left            =   90
      TabIndex        =   3
      Top             =   1950
      Width           =   2445
   End
   Begin VB.CommandButton cmdRefreshJno 
      Caption         =   "Refresh Voucher No"
      Height          =   555
      Left            =   90
      TabIndex        =   2
      Top             =   1410
      Width           =   2445
   End
   Begin VB.CommandButton cmdARRefresh 
      Caption         =   "AR Refresh"
      Height          =   555
      Left            =   90
      TabIndex        =   1
      Top             =   870
      Width           =   2445
   End
   Begin VB.CommandButton cmdRefreshEntries 
      Caption         =   "AMIS Refresher "
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2445
   End
End
Attribute VB_Name = "frm_TOOLS_ARRefresher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function SetCRJVoucherNo(XXX As String, zzz As Integer) As String
    Dim rsCRJ_Journal_HD                          As ADODB.Recordset
    Set rsCRJ_Journal_HD = New ADODB.Recordset
    If zzz = 1 Then
        Set rsCRJ_Journal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where Jtype = 'CRJ' and InvoiceNo = '" & XXX & "'")
    Else
        Set rsCRJ_Journal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where Jtype = 'CRJ' and LEFT(InvoiceNo,2) = 'NV' AND RIGHT(InvoiceNo,6) = '" & XXX & "'")
    End If
    If Not rsCRJ_Journal_HD.EOF And Not rsCRJ_Journal_HD.BOF Then
        SetCRJVoucherNo = Null2String(rsCRJ_Journal_HD!VOUCHERNO)
    End If
End Function

Sub CheckAccountsNotInTitleCode()
    Dim rsChartOfAccounts                         As ADODB.Recordset
    Dim rsTitleCode                               As ADODB.Recordset

    Set rsChartOfAccounts = New ADODB.Recordset
    Set rsChartOfAccounts = gconDMIS.Execute("Select * from AMIS_ChartAccount Order by AcctCode asc")
    If Not rsChartOfAccounts.EOF And Not rsChartOfAccounts.BOF Then
        rsChartOfAccounts.MoveFirst
        Do While Not rsChartOfAccounts.EOF
            Set rsTitleCode = New ADODB.Recordset
            Set rsTitleCode = gconDMIS.Execute("Select * from AMIS_TitleCode Where Code = " & N2Str2Null(rsChartOfAccounts!Titles))
            If rsTitleCode.EOF And rsTitleCode.BOF Then
                MsgBox "Account Code : " & rsChartOfAccounts!ACCTCODE & vbCrLf & _
                       "Account Desc : " & rsChartOfAccounts!Description & vbCrLf & _
                       " Is not in Account Sub-Totals", vbCritical, "Invalid Account Found!"
            End If
            rsChartOfAccounts.MoveNext
        Loop
    End If
    MsgBox "Done"
End Sub

Private Sub checkWOSTotals_Click()
    CheckAccountsNotInTitleCode
End Sub

Private Sub cmdAPRefresh_Click()
    Dim rsJournal_HD                              As ADODB.Recordset
    Dim rsJournal_Det                             As ADODB.Recordset

    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select ID,InvoiceType,InvoiceNo,VoucherNo,Jtype from AMIS_Journal_HD where jtype = 'APJ' order by id asc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst: TotalAmountToPay = 0
        Do While Not rsJournal_HD.EOF
            Set rsJournal_Det = New ADODB.Recordset
            Set rsJournal_Det = gconDMIS.Execute("Select SUM(CREDIT) AS TOTAL_AP from AMIS_Journal_Det Where (left(Acct_Code,5) = '21-01' OR left(Acct_Code,5)='21-02') AND Jtype = 'APJ' and VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO))
            If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
                gconDMIS.Execute ("Update AMIS_Journal_HD set AmountToPay = " & N2Str2Zero(rsJournal_Det!TOTAL_AP) & " Where ID = " & rsJournal_HD!ID)
            End If
            Me.Caption = Null2String(rsJournal_HD!jtype) & " " & Null2String(rsJournal_HD!VOUCHERNO): DoEvents
            Me.Caption = Me.Caption & " [" & Round((rsJournal_HD.AbsolutePosition / rsJournal_HD.RecordCount) * 100, 0) & "%]"
            rsJournal_HD.MoveNext
        Loop
    End If


    Dim rsCRJ_Detail                              As ADODB.Recordset
    'gconDMIS.Execute ("Update AMIS_Journal_HD set AmountPaid=0,Balance=AmountToPay where AmountToPay > 0 and Balance <> 0 and Jtype = 'SJ'")
    gconDMIS.Execute ("Update AMIS_Journal_HD set AmountPaid=0,Balance=AmountToPay where AmountToPay > 0 and Jtype = 'APJ'")
    Set rsCRJ_Detail = New ADODB.Recordset
    Set rsCRJ_Detail = gconDMIS.Execute("Select * from AMIS_CV_Detail Order by id asc")
    If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
        rsCRJ_Detail.MoveFirst
        Do While Not rsCRJ_Detail.EOF
            gconDMIS.Execute ("Update AMIS_Journal_HD Set AmountPaid=AmountPaid+" & rsCRJ_Detail!amount & ",BALANCE = BALANCE - " & rsCRJ_Detail!amount & " WHERE VoucherNo = '" & Null2String(rsCRJ_Detail!pv_voucherno) & "'")
            Me.Caption = Null2String(rsCRJ_Detail!pv_voucherno): DoEvents
            Me.Caption = Me.Caption & " [" & Round((rsCRJ_Detail.AbsolutePosition / rsCRJ_Detail.RecordCount) * 100, 0) & "%]"
            rsCRJ_Detail.MoveNext
        Loop
    Else
    End If



    MsgBox "Done"

End Sub

Private Sub cmdARRefresh_Click()

    Dim rsJournal_HD                              As ADODB.Recordset
    Dim rsJournal_Det                             As ADODB.Recordset


    Dim rsCheckCRJExist                           As ADODB.Recordset

    Dim PV_MRRNO, PV_INVNO, PV_PRODNO             As String
    Dim J_JVOUCHERNO                              As String
    Dim J_JDATE                                   As String
    Dim PV_AMOUNT                                 As Double
    Dim PV_STATUS, PV_ITEMNO                      As String

    Dim rsOFF_HD                                  As ADODB.Recordset
    Dim rsOFF_DT                                  As ADODB.Recordset

    Dim PayTranType                               As String
    Dim PayInvoiceNo                              As String
    Dim PayAmount                                 As Double


    Set rsOFF_HD = New ADODB.Recordset
    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD Order by ID ASC")
    If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
        rsOFF_HD.MoveFirst: I = 0
        Do While Not rsOFF_HD.EOF
            J_JDATE = N2Date2Null(rsOFF_HD!OR_DATE)
            Set rsOFF_DT = New ADODB.Recordset
            Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE (TRANTYPE <> 'OTH' OR PAIDFOR = 'SII' OR PAIDFOR = 'VII') AND OR_NUM = " & N2Str2Null(rsOFF_HD!OR_NUM))
            If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                rsOFF_DT.MoveFirst
                Do While Not rsOFF_DT.EOF
                    If Null2String(rsOFF_DT!TRANTYPE) = "OTH" Then
                        If Null2String(rsOFF_DT!PAIDFOR) = "SII" Then
                            PayTranType = "SI"
                            If Left(Null2String(rsOFF_DT!DESCRIPT), 3) = "ROB" Then
                                PayInvoiceNo = Format(Mid(Null2String(rsOFF_DT!DESCRIPT), 4, Len(Null2String(rsOFF_DT!DESCRIPT)) - 3), "000000")
                            End If
                        ElseIf Null2String(rsOFF_DT!PAIDFOR) = "VII" Then
                            PayTranType = "VI"
                            PayInvoiceNo = Null2String(rsOFF_DT!INVOICENO)
                        Else
                            MsgBox Null2String(rsOFF_DT!PAIDFOR)
                            PayTranType = Null2String(rsOFF_DT!PAIDFOR)
                            PayInvoiceNo = Format(Right(Null2String(rsOFF_DT!DESCRIPT), 4), "000000")
                        End If
                    Else
                        PayTranType = Null2String(rsOFF_DT!TRANTYPE)
                        PayInvoiceNo = Null2String(rsOFF_DT!INVOICENO)
                    End If

                    PayAmount = N2Str2Zero(rsOFF_DT!payment)

                    Set rsJournal_HD = New ADODB.Recordset
                    If Null2Bool(rsOFF_DT!VAT) = 1 Then
                        Set rsJournal_HD = gconDMIS.Execute("Select VoucherNo from AMIS_JOURNAL_HD where JTYPE = 'CRJ' and INVOICENO = " & N2Str2Null(rsOFF_DT!OR_NUM))
                    Else
                        Set rsJournal_HD = gconDMIS.Execute("Select VoucherNo from AMIS_JOURNAL_HD where JTYPE = 'CRJ' and INVOICENO = " & N2Str2Null(Right(rsOFF_DT!OR_NUM, 6)))
                    End If
                    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
                        Set rsCheckCRJExist = New ADODB.Recordset
                        Set rsCheckCRJExist = gconDMIS.Execute("Select * from AMIS_CRJ_Detail where VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO))
                        If rsCheckCRJExist.EOF And rsCheckCRJExist.BOF Then
                            Set rsSJ_DATA = New ADODB.Recordset
                            Set rsSJ_DATA = gconDMIS.Execute("Select * from AMIS_Journal_HD Where jtype = 'SJ' and invoicetype = " & N2Str2Null(PayTranType) & " and invoiceno = " & N2Str2Null(PayInvoiceNo))
                            If Not rsSJ_DATA.EOF And Not rsSJ_DATA.BOF Then
                                rsSJ_DATA.MoveFirst
                                Do While Not rsSJ_DATA.EOF
                                    J_JVOUCHERNO = "'" & SetCRJVoucherNo(Null2String(rsOFF_HD!OR_NUM), Null2String(rsOFF_HD!VAT)) & "'"
                                    PV_ITEMNO = N2Str2Null(Format(SJ_PV_ITEMNO, "0000"))
                                    PV_MRRNO = N2Str2Null(rsOFF_DT!TRANTYPE)
                                    PV_INVNO = N2Str2Null(rsOFF_DT!INVOICENO)
                                    PV_PRODNO = N2Date2Null(rsSJ_DATA!invoicedate)
                                    PV_AMOUNT = PayAmount  'N2Str2Zero(rsSJ_DATA!InvoiceAmt)
                                    PV_STATUS = "'N'"
                                    'INSERT IT

                                    gconDMIS.Execute "Delete from AMIS_CRJ_Detail Where VoucherNo = " & J_JVOUCHERNO & " AND JDate = " & J_JDATE & _
                                                     " AND INVOICETYPE = " & PV_MRRNO & _
                                                     " AND INVOICENO = " & PV_INVNO
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
                                        Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'COB'")
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
                        End If
                    End If

                    rsOFF_DT.MoveNext
                Loop
            End If
            I = I + 1
            Me.Caption = Round((rsOFF_HD.AbsolutePosition / rsOFF_HD.RecordCount) * 100, 0) & "%"
            DoEvents
            rsOFF_HD.MoveNext
        Loop
    End If
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select ID,InvoiceType,InvoiceNo,VoucherNo from AMIS_Journal_HD where jtype = 'SJ' order by id asc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst: TotalAmountToPay = 0
        Do While Not rsJournal_HD.EOF
            Set rsJournal_Det = New ADODB.Recordset
            Set rsJournal_Det = gconDMIS.Execute("Select SUM(DEBIT) AS TOTAL_AR from AMIS_Journal_Det Where (left(Acct_Code,5) = '11-02' OR left(Acct_Code,5)='11-03') AND Jtype = 'SJ' and VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO))
            If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
                gconDMIS.Execute ("Update AMIS_Journal_HD set AmountToPay = " & N2Str2Zero(rsJournal_Det!TOTAL_AR) & " Where ID = " & rsJournal_HD!ID)
            End If
            Me.Caption = Null2String(rsJournal_HD!InvoiceType) & " " & Null2String(rsJournal_HD!INVOICENO): DoEvents
            Me.Caption = Me.Caption & " [" & Round((rsJournal_HD.AbsolutePosition / rsJournal_HD.RecordCount) * 100, 0) & "%]"
            rsJournal_HD.MoveNext
        Loop
    End If


    Dim rsCRJ_Detail                              As ADODB.Recordset
    gconDMIS.Execute ("Update AMIS_Journal_HD set AmountPaid=0,Balance=AmountToPay where AmountToPay > 0 and Balance <> 0 and Jtype = 'SJ'")
    Set rsCRJ_Detail = New ADODB.Recordset
    Set rsCRJ_Detail = gconDMIS.Execute("Select * from AMIS_CRJ_Detail Order by id asc")
    If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
        rsCRJ_Detail.MoveFirst
        Do While Not rsCRJ_Detail.EOF
            gconDMIS.Execute ("Update AMIS_Journal_HD Set AmountPaid=AmountPaid+" & rsCRJ_Detail!invoiceamount & ",BALANCE = BALANCE - " & rsCRJ_Detail!invoiceamount & " WHERE InvoiceType = '" & Null2String(rsCRJ_Detail!InvoiceType) & "' AND InvoiceNo = '" & Null2String(rsCRJ_Detail!INVOICENO) & "'")
            Me.Caption = Null2String(rsCRJ_Detail!InvoiceType) & " " & Null2String(rsCRJ_Detail!INVOICENO): DoEvents
            Me.Caption = Me.Caption & " [" & Round((rsCRJ_Detail.AbsolutePosition / rsCRJ_Detail.RecordCount) * 100, 0) & "%]"
            rsCRJ_Detail.MoveNext
        Loop
    Else
    End If
    MsgBox "AR Refresh Tapos"

    Exit Sub


    'UnBalanceCRJ:
    '
    'Dim rsCRJ_HD As ADODB.Recordset
    'Dim rsCRJ_Det As ADODB.Recordset
    '
    'Set rsCRJ_HD = New ADODB.Recordset
    'Set rsCRJ_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where JTYPE = 'CRJ' Order by id asc")
    'If Not rsCRJ_HD.EOF And Not rsCRJ_HD.BOF Then
    '   rsCRJ_HD.MoveFirst
    '   Do While Not rsCRJ_HD.EOF
    '      Set rsCRJ_Det = New ADODB.Recordset
    '      Set rsCRJ_Det = gconDMIS.Execute("Select SUM(INVOICEAMOUNT) as TOTALPAYMENT from AMIS_CRJ_Detail where VoucherNo = " & N2Str2Null(rsCRJ_HD!VoucherNo))
    '      If Not rsCRJ_Det.EOF And Not rsCRJ_Det.BOF Then
    '         If N2Str2Zero(rsCRJ_Det!TOTALPAYMENT) <> N2Str2Zero(rsCRJ_HD!InvoiceAmt) And N2Str2Zero(rsCRJ_Det!TOTALPAYMENT) > 0 Then
    '            gconDMIS.Execute "Update AMIS_Journal_HD Set Dealer = 'CRJ' Where id = " & rsCRJ_HD!ID
    '         End If
    '      End If
    '      rsCRJ_HD.MoveNext
    '   Loop
    'End If
    'MsgBox "Ok na"
    'Exit Sub
End Sub

Private Sub cmdCheckErrorTrans_Click()
    frmAMISCheckDupTrans.Show
End Sub

Private Sub cmdInvalidAcctCode_Click()
    Dim rsJournal_Det                             As ADODB.Recordset
    Dim rsChartAccount                            As ADODB.Recordset
    Screen.MousePointer = 11
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_Det WHERE STATUS = 'P' order by JDate asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Do While Not rsJournal_Det.EOF
            Set rsChartAccount = New ADODB.Recordset
            Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where AcctCode = " & N2Str2Null(rsJournal_Det!Acct_code))
            If rsChartAccount.EOF And rsChartAccount.BOF Then
                MsgBox "Invalid Account! " & vbCrLf & _
                       "Account Code : " & rsJournal_Det!Acct_code & vbCrLf & _
                       "Account Desc : " & rsJournal_Det!acct_Name
            End If
            Me.Caption = rsJournal_Det!jtype & " " & rsJournal_Det!VOUCHERNO: DoEvents
            rsJournal_Det.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
    MsgBox "Done"
End Sub

Private Sub cmdRefreshEntries_Click()
    Dim rsJournal_Det                             As ADODB.Recordset
    Dim rsJournal_HD                              As ADODB.Recordset
    Dim X                                         As Double
    gconDMIS.Execute ("DELETE from AMIS_Journal_Det WHERE VOUCHERNO IS NULL")
    gconDMIS.Execute ("DELETE from AMIS_Journal_HD WHERE VOUCHERNO IS NULL")

    'Check Un-Balance
    Set rsJournal_HD = New ADODB.Recordset
    'Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where jdate = '9/9/2008' and jtype <> 'OPB' AND STATUS = 'P' Order by id asc")
    Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where jtype <> 'OPB' AND STATUS = 'P' and jdate <= '3/1/2009'  Order by id asc")
    'Set rsJournal_hd = gconDMIS.Execute("Select * from AMIS_Journal_HD where jtype <> 'OPB' AND STATUS = 'P' Order by id asc")

    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_HD.EOF
            Set rsJournal_Det = New ADODB.Recordset
            Set rsJournal_Det = gconDMIS.Execute("Select SUM(DEBIT) AS TotalDebit,SUM(CREDIT) AS TotalCredit from AMIS_Journal_Det Where JType = " & N2Str2Null(rsJournal_HD!jtype) & " And VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO) & " AND STATUS = 'P'")

            If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
                If Round(N2Str2Zero(rsJournal_Det!TotalDebit), 2) <> Round(N2Str2Zero(rsJournal_Det!Totalcredit), 2) Then
                    gconDMIS.Execute "update AMIS_Journal_HD set " & _
                                     " Debit = " & Round(N2Str2Zero(rsJournal_Det!TotalDebit), 2) & "," & _
                                     " Credit = " & Round(N2Str2Zero(rsJournal_Det!Totalcredit), 2) & "," & _
                                     " Status = 'N'" & _
                                     " Where JType = " & N2Str2Null(rsJournal_HD!jtype) & " AND VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO)
                    gconDMIS.Execute "update AMIS_Journal_Det set " & _
                                     " Status = 'N'" & _
                                     " Where JType = " & N2Str2Null(rsJournal_HD!jtype) & " AND VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO)
                Else
                    gconDMIS.Execute "update AMIS_Journal_HD set " & _
                                     " Debit = " & Round(N2Str2Zero(rsJournal_Det!TotalDebit), 2) & "," & _
                                     " Credit = " & Round(N2Str2Zero(rsJournal_Det!Totalcredit), 2) & "," & _
                                     " Status = 'P'" & _
                                     " Where JType = " & N2Str2Null(rsJournal_HD!jtype) & " AND VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO)
                    gconDMIS.Execute "update AMIS_Journal_Det set " & _
                                     " Status = 'P'" & _
                                     " Where JType = " & N2Str2Null(rsJournal_HD!jtype) & " AND VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO)
                End If
            Else
                gconDMIS.Execute ("Update AMIS_Journal_HD Set Status = 'N' where id = " & rsJournal_HD!ID)
            End If
            Me.Caption = rsJournal_HD!jtype & " " & rsJournal_HD!VOUCHERNO: DoEvents
            rsJournal_HD.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    gconDMIS.Execute ("Update AMIS_Journal_HD Set STATUS = 'N' WHERE STATUS <> 'P' AND STATUS <> 'C'")
    gconDMIS.Execute ("Update AMIS_Journal_Det Set STATUS = 'N' WHERE STATUS <> 'P' AND STATUS <> 'C'")

    'Check Lost
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_Det where jdate <= '3/1/2009' order by JNO ASC")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Do While Not rsJournal_Det.EOF
            X = X + 1
            Set rsJournal_HD = New ADODB.Recordset
            Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_journal_HD where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & "  and VoucherNo = " & N2Str2Null(rsJournal_Det!VOUCHERNO))
            If rsJournal_HD.EOF And rsJournal_HD.BOF Then
                gconDMIS.Execute ("Delete from AMIS_Journal_Det where ID = " & rsJournal_Det!ID)
            End If
            Me.Caption = rsJournal_Det!jtype & rsJournal_Det!VOUCHERNO & "-" & X: DoEvents
            rsJournal_Det.MoveNext
        Loop
    End If

    'Refresh Amounts
    Dim rsJournal_Det_Trans                       As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_Det where jdate <= '5/31/2008' Group By Jtype,VoucherNo Having (SUM(Debit) <> SUM(Credit)) Order by Jtype,VoucherNo")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        X = 0
        Do While Not rsJournal_Det.EOF
            X = X + 1
            Set rsJournal_Det_Trans = New ADODB.Recordset
            Set rsJournal_Det_Trans = gconDMIS.Execute("Select * from AMIS_Journal_Det Where JType = " & N2Str2Null(rsJournal_Det!jtype) & " AND VoucherNo = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " Order by JItemNo asc")
            If Not rsJournal_Det_Trans.EOF And Not rsJournal_Det_Trans.BOF Then
                rsJournal_Det_Trans.MoveFirst
                Do While Not rsJournal_Det_Trans.EOF
                    gconDMIS.Execute ("Update AMIS_Journal_Det set Debit = " & Round(NumericVal(rsJournal_Det_Trans!DEBIT), 2) & ", Credit = " & Round(NumericVal(rsJournal_Det_Trans!CREDIT), 2) & " where id = " & rsJournal_Det_Trans!ID)
                    rsJournal_Det_Trans.MoveNext
                Loop
            End If
            DoEvents
            Me.Caption = rsJournal_Det!VOUCHERNO & "-" & X
            rsJournal_Det.MoveNext
        Loop
    End If

    MsgBox "Completed"
End Sub

Private Sub cmdRefreshTemplates_Click()
    Dim rsChartAccount                            As ADODB.Recordset
    Dim rsTemplate_Header                         As ADODB.Recordset
    Dim rsTemplate_Details                        As ADODB.Recordset
Repeat:     Set rsTemplate_Details = New ADODB.Recordset
    Set rsTemplate_Details = gconDMIS.Execute("Select * from AMIS_Template_Details")
    If Not rsTemplate_Details.EOF And Not rsTemplate_Details.BOF Then
        rsTemplate_Details.MoveFirst
        'MsgBox "Poon"
        Do While Not rsTemplate_Details.EOF
            Set rsTemplate_Header = New ADODB.Recordset
            Set rsTemplate_Header = gconDMIS.Execute("Select * from AMIS_Template_Header Where TemplateCode = " & N2Str2Null(rsTemplate_Details!TemplateCode))
            If Not rsTemplate_Header.EOF And Not rsTemplate_Header.BOF Then
                Set rsChartAccount = New ADODB.Recordset
                Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount where AcctCode = " & N2Str2Null(rsTemplate_Details!AccountCode))
                If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
                    gconDMIS.Execute ("update AMIS_Template_Details Set Description = " & UCase(N2Str2Null(rsChartAccount!Description)) & " Where Code = " & rsTemplate_Details!code)
                Else
                    'MsgBox "Invalid Account Code for this Template... this Template will be Automatically deleted after this message"
                    gconDMIS.Execute ("Delete from AMIS_Template_Details where Code = " & rsTemplate_Details!code)
                    'gconDMIS.Execute ("Delete from AMIS_Template_Header where TemplateCode = " & rsTemplate_Details!TemplateCode)
                    'GoTo Repeat
                    'gconDMIS.Execute ("update AMIS_Template_Details Set Remarks = 'Invalid' Where Code = " & rsTemplate_Details!Code)
                End If
            Else
                'MsgBox "Invalid Template Code for this Template... this Template will be Automatically deleted after this message"
                gconDMIS.Execute ("Delete from AMIS_Template_Details where Code = " & rsTemplate_Details!code)
                'GoTo Repeat
            End If
            rsTemplate_Details.MoveNext
        Loop
        MsgBox "Tapos"
    End If
End Sub

Private Sub UpdateJournalNo()
    Dim rsJournal_Det                             As ADODB.Recordset
    Dim rsChartAccounts                           As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            gconDMIS.Execute ("Update AMIS_Journal_Det set JNo = " & N2Str2Null(rsJournal_Det!JNo) & ", Jdate = " & N2Date2Null(rsJournal_Det!JDate) & ", status = " & N2Str2Null(rsJournal_Det!Status) & " where VoucherNo = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and Jtype = " & N2Str2Null(rsJournal_Det!jtype))
            Me.Caption = rsJournal_Det!JNo: DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    MsgBox "ok"
    Exit Sub

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD where status = 'P' order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccounts = New ADODB.Recordset
            Set rsChartAccounts = gconDMIS.Execute("Select * from AMIS_Journal_det where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & " and voucherno = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and status <> 'P'")
            If Not rsChartAccounts.EOF And Not rsChartAccounts.BOF Then
                gconDMIS.Execute ("update AMIS_Journal_Det set status = 'P' where id = " & rsChartAccounts!ID)
                'MsgBox "Invalid Posted Transaction" & vbCrLf & _
                 Null2String(rsChartAccounts!acct_code) & " " & Null2String(rsChartAccounts!acct_Name) & vbCrLf & _
                 Null2String(rsJOURNAL_DET!Jtype) & "-" & Null2String(rsJOURNAL_DET!voucherno)
            End If
            Me.Caption = Null2String(rsJournal_Det!JNo): DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD where status = 'N' order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccounts = New ADODB.Recordset
            Set rsChartAccounts = gconDMIS.Execute("Select * from AMIS_Journal_det where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & " and voucherno = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and status <> 'N'")
            If Not rsChartAccounts.EOF And Not rsChartAccounts.BOF Then
                gconDMIS.Execute ("update AMIS_Journal_Det set status = 'N' where id = " & rsChartAccounts!ID)
                'MsgBox "Invalid Un-Posted Transaction" & vbCrLf & _
                 Null2String(rsChartAccounts!acct_code) & " " & Null2String(rsChartAccounts!acct_Name) & vbCrLf & _
                 Null2String(rsJOURNAL_DET!Jtype) & "-" & Null2String(rsJOURNAL_DET!voucherno)
            End If
            Me.Caption = Null2String(rsJournal_Det!JNo): DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD where status = 'C' order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccounts = New ADODB.Recordset
            Set rsChartAccounts = gconDMIS.Execute("Select * from AMIS_Journal_det where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & " and voucherno = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and status <> 'C'")
            If Not rsChartAccounts.EOF And Not rsChartAccounts.BOF Then
                gconDMIS.Execute ("update AMIS_Journal_Det set status = 'C' where id = " & rsChartAccounts!ID)
                'MsgBox "Invalid Cancelled Transaction" & vbCrLf & _
                 Null2String(rsChartAccounts!acct_code) & " " & Null2String(rsChartAccounts!acct_Name) & vbCrLf & _
                 Null2String(rsJOURNAL_DET!Jtype) & "-" & Null2String(rsJOURNAL_DET!voucherno)
            End If
            Me.Caption = Null2String(rsJournal_Det!JNo): DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If



    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            gconDMIS.Execute ("Update AMIS_Journal_Det set JNo = " & N2Str2Null(rsJournal_Det!JNo) & ", Jdate = " & N2Date2Null(rsJournal_Det!JDate) & ", status = " & N2Str2Null(rsJournal_Det!Status) & " where VoucherNo = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and Jtype = " & N2Str2Null(rsJournal_Det!jtype))
            Me.Caption = rsJournal_Det!JNo: DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    MsgBox "ok"
    Exit Sub

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD where status = 'P' order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccounts = New ADODB.Recordset
            Set rsChartAccounts = gconDMIS.Execute("Select * from AMIS_Journal_det where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & " and voucherno = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and status <> 'P'")
            If Not rsChartAccounts.EOF And Not rsChartAccounts.BOF Then
                gconDMIS.Execute ("update AMIS_Journal_Det set status = 'P' where id = " & rsChartAccounts!ID)
                'MsgBox "Invalid Posted Transaction" & vbCrLf & _
                 Null2String(rsChartAccounts!acct_code) & " " & Null2String(rsChartAccounts!acct_Name) & vbCrLf & _
                 Null2String(rsJOURNAL_DET!Jtype) & "-" & Null2String(rsJOURNAL_DET!voucherno)
            End If
            Me.Caption = Null2String(rsJournal_Det!JNo): DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD where status = 'N' order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccounts = New ADODB.Recordset
            Set rsChartAccounts = gconDMIS.Execute("Select * from AMIS_Journal_det where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & " and voucherno = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and status <> 'N'")
            If Not rsChartAccounts.EOF And Not rsChartAccounts.BOF Then
                gconDMIS.Execute ("update AMIS_Journal_Det set status = 'N' where id = " & rsChartAccounts!ID)
                'MsgBox "Invalid Un-Posted Transaction" & vbCrLf & _
                 Null2String(rsChartAccounts!acct_code) & " " & Null2String(rsChartAccounts!acct_Name) & vbCrLf & _
                 Null2String(rsJOURNAL_DET!Jtype) & "-" & Null2String(rsJOURNAL_DET!voucherno)
            End If
            Me.Caption = Null2String(rsJournal_Det!JNo): DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD where status = 'C' order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccounts = New ADODB.Recordset
            Set rsChartAccounts = gconDMIS.Execute("Select * from AMIS_Journal_det where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & " and voucherno = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and status <> 'C'")
            If Not rsChartAccounts.EOF And Not rsChartAccounts.BOF Then
                gconDMIS.Execute ("update AMIS_Journal_Det set status = 'C' where id = " & rsChartAccounts!ID)
                'MsgBox "Invalid Cancelled Transaction" & vbCrLf & _
                 Null2String(rsChartAccounts!acct_code) & " " & Null2String(rsChartAccounts!acct_Name) & vbCrLf & _
                 Null2String(rsJOURNAL_DET!Jtype) & "-" & Null2String(rsJOURNAL_DET!voucherno)
            End If
            Me.Caption = Null2String(rsJournal_Det!JNo): DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    MsgBox "Completed!"
End Sub

Private Sub Command10_Click()
    Dim rsJournal_HDDet                           As ADODB.Recordset
    Dim rsCHART_ACCOUNTS                          As ADODB.Recordset

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE             As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME           As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET  As Double
    Dim J_STATUS, J_JITEMNO                       As String
    Dim J_ATC                                     As String
    Dim J_RATE, J_TAXBASE                         As Double
    Dim ItemCount                                 As Integer
    J_JDATE = CDate("12/31/2004")
    Set rsCHART_ACCOUNTS = New ADODB.Recordset
    Set rsCHART_ACCOUNTS = gconDMIS.Execute("Select * from AMIS_ChartAccount Where (HeaderCode <> '1' and HeaderCode <> '2' and HeaderCode <> '3') order by acctcode asc")
    If Not rsCHART_ACCOUNTS.EOF And Not rsCHART_ACCOUNTS.BOF Then
        rsCHART_ACCOUNTS.MoveFirst: ItemCount = 0
        gconDMIS.Execute ("delete from AMIS_Journal_Det where jtype = 'CLO' and voucherno = '000002'")
        Do While Not rsCHART_ACCOUNTS.EOF
            Set rsJournal_HDDet = New ADODB.Recordset
            rsJournal_HDDet.Open "select SUM(DEBIT) AS TOTAL_DEBIT,SUM(CREDIT) AS TOTAL_CREDIT from vLEDGER where Jdate <= '" & J_JDATE & "' and Acct_Code = " & N2Str2Null(rsCHART_ACCOUNTS!ACCTCODE), gconDMIS
            If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
                'J_JDATE = "'" & J_JDATE & "'"
                J_VOUCHERNO = N2Str2Null("000002")
                J_JTYPE = N2Str2Null("CLO")
                J_JNO = N2Str2Null("020650")
                J_ACCT_CODE = N2Str2Null(Null2String(rsCHART_ACCOUNTS!ACCTCODE))
                J_ACCT_NAME = N2Str2Null(Null2String(rsCHART_ACCOUNTS!Description))
                If N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT) > 0 Then
                    J_DEBIT = 0
                    J_CREDIT = Abs(N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT))
                Else
                    J_DEBIT = Abs(N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT))
                    J_CREDIT = 0
                End If
                J_TAX = 0
                J_GROSS = 0
                J_NET = 0
                J_STATUS = "'N'"
                J_ATC = "NULL"
                J_RATE = 0
                J_TAXBASE = 0
                If J_DEBIT > 0 Or J_CREDIT > 0 Then
                    ItemCount = ItemCount + 1
                    J_JITEMNO = N2Str2Null(Format(ItemCount, "0000"))
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,USERCODE,LASTUPDATE,ATC,RATE,TAXBASE)" & _
                                     " values ('" & CDate("1/1/2005") & "', " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ",'" & LOGCODE & "','" & LOGDATE & "'," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"
                End If
            End If
            Me.Caption = Null2String(rsCHART_ACCOUNTS!Description): DoEvents
            rsCHART_ACCOUNTS.MoveNext
        Loop
    End If
    Set rsCHART_ACCOUNTS = Nothing
    Set rsJournal_HDDet = Nothing
End Sub

Private Sub Command11_Click()
    Dim rsJournal_Det                             As ADODB.Recordset
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select ID,Voucherno,jtype,acct_code,acct_name,Jno,status from AMIS_Journal_Det Order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        MsgBox "Poon"
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccount = New ADODB.Recordset
            Set rsChartAccount = gconDMIS.Execute("Select jno,status from AMIS_Journal_HD Where status <> " & N2Str2Null(rsJournal_Det!Status) & " and JNo = " & N2Str2Null(rsJournal_Det!JNo))
            If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
                If Null2String(rsChartAccount!Status) = "P" Then
                    gconDMIS.Execute ("update AMIS_Journal_Det SET STATUS = 'P' WHERE JNO = " & N2Str2Null(rsJournal_Det!JNo))
                Else
                    MsgBox "HEADER STATUS = (" & Null2String(rsChartAccount!Status) & ")" & vbCrLf & _
                           "DETAIL STATUS = (" & Null2String(rsJournal_Det!Status) & ")" & vbCrLf & _
                           Null2String(rsJournal_Det!jtype) & "-" & _
                           Null2String(rsJournal_Det!VOUCHERNO) & vbCrLf & _
                           Null2String(rsJournal_Det!acct_Name)
                End If
            End If
            Me.Caption = "[" & rsJournal_Det!ID & "] " & Null2String(rsJournal_Det!VOUCHERNO)
            DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
        MsgBox "Tapos"
    End If
End Sub

Private Sub Command13_Click()
    Dim rsAR_OPENING                              As ADODB.Recordset
    Set rsAR_OPENING = New ADODB.Recordset
    Dim J_JNO                                     As String
    Set rsAR_OPENING = gconDMIS.Execute("Select * from AR_OPENING Order by VoucherNo asc")
    If Not rsAR_OPENING.EOF And Not rsAR_OPENING.BOF Then
        rsAR_OPENING.MoveFirst
        Do While Not rsAR_OPENING.EOF
            Set rsJournal_HDDup = New ADODB.Recordset
            Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
            If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
            Else
                J_JNO = "'000001'"
            End If
            gconDMIS.Execute ("Insert into AMIS_Journal_HD (JNO,JDate,VoucherNo,Jtype,CustomerCode,CustomerName,InvoiceType,InvoiceNo,InvoiceDate,InvoiceAmt)" & _
                              " values (" & J_JNO & "," & N2Str2Null(rsAR_OPENING!JDate) & "," & N2Str2Null(Format(rsAR_OPENING!VOUCHERNO, "000000")) & "," & N2Str2Null(rsAR_OPENING!jtype) & _
                              ",NULL," & N2Str2Null(rsAR_OPENING!CUSTOMERNAME) & "," & N2Str2Null(rsAR_OPENING!InvoiceType) & "," & N2Str2Null(rsAR_OPENING!INVOICENO) & "," & N2Str2Null(rsAR_OPENING!invoicedate) & "," & N2Str2Null(rsAR_OPENING!InvoiceAmt) & ")")
            Me.Caption = rsAR_OPENING!VOUCHERNO
            rsAR_OPENING.MoveNext
            DoEvents
        Loop
    End If
    MsgBox "Tapos"
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Dim rsJournal_HD                              As ADODB.Recordset
    Dim rsCustomer                                As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Order by id asc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_HD.EOF
            Set rsCustomer = New ADODB.Recordset
            Set rsCustomer = gconDMIS.Execute("select * from AMIS_Customer where CustCode = " & N2Str2Null(rsJournal_HD!CustomerCode))
            If rsCustomer.EOF And rsCustomer.BOF Then
                MsgBox Null2String(rsJournal_HD!CustomerCode) & " is Invalid!"
            End If
            rsJournal_HD.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    MsgBox "Completed!"
End Sub

Private Sub Command5_Click()
    Dim rsJournal_HD                              As ADODB.Recordset
    Dim rsJournal_Det                             As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where jtype <> 'OPB' Order by id asc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_HD.EOF
            Set rsJournal_Det = New ADODB.Recordset
            Set rsJournal_Det = gconDMIS.Execute("Select SUM(DEBIT) AS TotalDebit,SUM(CREDIT) AS TotalCredit from AMIS_Journal_Det Where JType = " & N2Str2Null(rsJournal_HD!jtype) & " And VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO))
            If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
                If Round(N2Str2Zero(rsJournal_Det!TotalDebit), 2) <> Round(N2Str2Zero(rsJournal_Det!Totalcredit), 2) Then
                    gconDMIS.Execute "update AMIS_Journal_HD set " & _
                                     " Debit = " & N2Str2Zero(rsJournal_Det!TotalDebit) & "," & _
                                     " Credit = " & N2Str2Zero(rsJournal_Det!Totalcredit) & "," & _
                                     " Status = 'N'" & _
                                     " Where JType = " & N2Str2Null(rsJournal_HD!jtype) & " AND VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO)
                    MsgBox Null2String(rsJournal_HD!jtype) & "-" & Null2String(rsJournal_HD!VOUCHERNO) & " is not Balance."
                Else
                    gconDMIS.Execute "update AMIS_Journal_HD set " & _
                                     "Debit = " & N2Str2Zero(rsJournal_Det!TotalDebit) & "," & _
                                     "Credit = " & N2Str2Zero(rsJournal_Det!Totalcredit) & _
                                     " Where JType = " & N2Str2Null(rsJournal_HD!jtype) & " AND VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO)
                End If
            Else
                MsgBox Null2String(rsJournal_HD!jtype) & "-" & Null2String(rsJournal_HD!VOUCHERNO) & " has no Detail."
            End If
            rsJournal_HD.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    MsgBox "Completed!"
End Sub

Private Sub Command6_Click()
    Dim rsSTKSTAT                                 As ADODB.Recordset
    Dim rsRANKFLE                                 As ADODB.Recordset
    Set rsSTKSTAT = New ADODB.Recordset
    Set rsSTKSTAT = gconDMIS.Execute("Select * from STKSTAT where date_gen = '10/30/2004'")
    If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
        rsSTKSTAT.MoveFirst
        Do While Not rsSTKSTAT.EOF
            rsSTKSTAT.MoveNext
        Loop
    End If
End Sub

Private Sub Command7_Click()
    Dim rsSMIS_CUSTOMER                           As ADODB.Recordset
    Dim rsAMIS_CUSTOMER                           As ADODB.Recordset
    Dim rsCSMIOS_CUSTOMER                         As ADODB.Recordset

    Set rsCSMIOS_CUSTOMER = New ADODB.Recordset
    Set rsCSMIOS_CUSTOMER = gconDMIS.Execute("Select * from CUSMAS order by CUSCDE ASC")
    If Not rsCSMIOS_CUSTOMER.EOF And Not rsCSMIOS_CUSTOMER.BOF Then
        rsCSMIOS_CUSTOMER.MoveFirst
        Do While Not rsCSMIOS_CUSTOMER.EOF And Not rsCSMIOS_CUSTOMER.BOF
            Set rsSMIS_CUSTOMER = New ADODB.Recordset
            Set rsSMIS_CUSTOMER = gconDMIS.Execute("Select * from AMIS_Customer Where LastName = " & N2Str2Null(rsCSMIOS_CUSTOMER!lastname) & " and FirstName = " & N2Str2Null(rsCSMIOS_CUSTOMER!Firstname))
            If Not rsSMIS_CUSTOMER.EOF And Not rsSMIS_CUSTOMER.BOF Then
                gconDMIS.Execute ("Update CUSVEH Set CUSCDE = " & N2Str2Null(rsSMIS_CUSTOMER!CUSCDE))
            End If
            rsCSMIOS_CUSTOMER.MoveNext
        Loop
    End If
End Sub

Private Sub Command8_Click()
    Screen.MousePointer = 11
    Dim rsMYOBJournals                            As ADODB.Recordset
    Set rsMYOBJournals = New ADODB.Recordset
    Set rsMYOBJournals = gconDMIS.Execute("Select * from MYOBJournal order by id asc")
    If Not rsMYOBJournals.EOF And Not rsMYOBJournals.BOF Then
        rsMYOBJournals.MoveFirst
        MsgBox "Poon"
        Do While Not rsMYOBJournals.EOF
            gconDMIS.Execute ("Update MYOBJournal Set [Memo] = " & N2Str2Null(rsMYOBJournals!Memo) & " Where id = " & rsMYOBJournals!ID)
            rsMYOBJournals.MoveNext
        Loop
        MsgBox "Tapos"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Command9_Click()
    Dim rsJournal_Det                             As ADODB.Recordset
    Dim rsChartAccount                            As ADODB.Recordset
    MsgBox "Poon"
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select ID,Acct_Code,Acct_Name from AMIS_Journal_Det Order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccount = New ADODB.Recordset
            Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where AcctCode = " & N2Str2Null(rsJournal_Det!Acct_code))
            If rsChartAccount.EOF And rsChartAccount.BOF Then
                MsgBox (Null2String(rsJournal_Det!Acct_code) & vbCrLf & _
                        Null2String(rsJournal_Det!acct_Name))
                If MsgBox("Delete this Account?", vbQuestion + vbYesNo, "Delete..") = vbYes Then
                    gconDMIS.Execute ("Delete from AMIS_Journal_Det where id = " & rsJournal_Det!ID)
                End If
            End If
            Me.Caption = "[" & rsJournal_Det!ID & "] " & Null2String(rsJournal_Det!Acct_code)
            DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_Det Order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccount = New ADODB.Recordset
            Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_Journal_HD Where Jno = " & N2Str2Null(rsJournal_Det!JNo))
            If rsChartAccount.EOF And rsChartAccount.BOF Then
                MsgBox (Null2String(rsJournal_Det!jtype) & vbCrLf & _
                        Null2String(rsJournal_Det!VOUCHERNO))
                If MsgBox("Delete this Entry?", vbQuestion + vbYesNo, "Delete..") = vbYes Then
                    gconDMIS.Execute ("Delete from AMIS_Journal_Det where id = " & rsJournal_Det!ID)
                End If
            End If
            Me.Caption = "[" & rsJournal_Det!ID & "] " & Null2String(rsJournal_Det!Acct_code)
            DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    MsgBox "Tapos"
End Sub

Private Sub cmdRefreshInvTerm_Click()
    Dim rsJournal_HD                              As ADODB.Recordset

    Dim rsOrd_Hd                                  As ADODB.Recordset
    Screen.MousePointer = 11
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where (InvoiceType = 'PI' or InvoiceType = 'AI' or InvoiceType = 'MI') AND Jtype = 'SJ' order by id asc")
    'Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where (InvoiceType = 'PI') AND Jtype = 'SJ' order by id asc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst
        Do While Not rsJournal_HD.EOF
            Set rsOrd_Hd = New ADODB.Recordset
            If Null2String(rsJournal_HD!InvoiceType) = "PI" Then
                Set rsOrd_Hd = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY where (TRANTYPE = 'CSH' OR TRANTYPE = 'CHG') AND TYPE = 'P' and TRANNO = " & N2Str2Null(rsJournal_HD!INVOICENO))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    gconDMIS.Execute ("Update AMIS_Journal_HD Set TERMS = " & N2Str2Null(rsOrd_Hd!TRANTYPE) & " Where id = " & rsJournal_HD!ID)
                End If
            End If
            If Null2String(rsJournal_HD!InvoiceType) = "AI" Then
                Set rsOrd_Hd = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY where (TRANTYPE = 'CSH' OR TRANTYPE = 'CHG') AND TYPE = 'A' and TRANNO = " & N2Str2Null(rsJournal_HD!INVOICENO))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    gconDMIS.Execute ("Update AMIS_Journal_HD Set TERMS = " & N2Str2Null(rsOrd_Hd!TRANTYPE) & " Where id = " & rsJournal_HD!ID)
                End If
            End If
            If Null2String(rsJournal_HD!InvoiceType) = "MI" Then
                Set rsOrd_Hd = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY where (TRANTYPE = 'CSH' OR TRANTYPE = 'CHG') AND TYPE = 'M' and TRANNO = " & N2Str2Null(rsJournal_HD!INVOICENO))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    gconDMIS.Execute ("Update AMIS_Journal_HD Set TERMS = " & N2Str2Null(rsOrd_Hd!TRANTYPE) & " Where id = " & rsJournal_HD!ID)
                End If
            End If
            Me.Caption = rsJournal_HD!jtype & "-" & rsJournal_HD!VOUCHERNO: DoEvents
            rsJournal_HD.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
    MsgBox "Tapos Na"
End Sub

Private Sub cmdRefreshJno_Click()
    Dim rsdistinctJTYPE                           As ADODB.Recordset

    Dim rsJournal_HD                              As ADODB.Recordset
    Dim KIM                                       As Long
    Screen.MousePointer = 11
    Set rsdistinctJTYPE = New ADODB.Recordset
    Set rsdistinctJTYPE = gconDMIS.Execute("Select Distinct JTYPE FROM AMIS_Journal_HD WHERE JTYPE = 'COB' OR JTYPE = 'VPJ' Order by Jtype asc")
    If Not rsdistinctJTYPE.EOF And Not rsdistinctJTYPE.BOF Then
        rsdistinctJTYPE.MoveFirst
        Do While Not rsdistinctJTYPE.EOF
            Set rsJournal_HD = New ADODB.Recordset
            Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where JType = '" & rsdistinctJTYPE!jtype & "' Order by Jno asc")
            If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
                rsJournal_HD.MoveFirst: KIM = 0
                Do While Not rsJournal_HD.EOF
                    KIM = KIM + 1
                    gconDMIS.Execute ("Update AMIS_Journal_HD Set VoucherNo = '" & Format(KIM, "000000") & "' where Jno = " & N2Str2Null(rsJournal_HD!JNo))
                    gconDMIS.Execute ("Update AMIS_Journal_Det Set VoucherNo = '" & Format(KIM, "000000") & "' where Jno = " & N2Str2Null(rsJournal_HD!JNo))
                    If rsdistinctJTYPE!jtype = "CRJ" Then
                        gconDMIS.Execute ("Update AMIS_CRJ_Detail Set VoucherNo = '" & Format(KIM, "000000") & "' where VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO))
                    End If
                    If rsdistinctJTYPE!jtype = "CDJ" Then
                        gconDMIS.Execute ("Update AMIS_CV_Detail Set VoucherNo = '" & Format(KIM, "000000") & "' where VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO))
                    End If
                    If rsdistinctJTYPE!jtype = "APJ" Then
                        gconDMIS.Execute ("Update AMIS_PV_Detail Set VoucherNo = '" & Format(KIM, "000000") & "' where VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO))
                    End If
                    gconDMIS.Execute ("Update AMIS_Journal_Det Set VoucherNo = '" & Format(KIM, "000000") & "' where Jno = " & N2Str2Null(rsJournal_HD!JNo))
                    Me.Caption = rsdistinctJTYPE!jtype & " " & Format(KIM, "000000") & " Jno = " & rsJournal_HD!JNo: DoEvents
                    rsJournal_HD.MoveNext
                Loop
            End If
            rsdistinctJTYPE.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
    MsgBox "tapos"
End Sub

Private Sub cmdShowErrorQuery_Click()
    frmAMISErrorQuery.Show
End Sub

Private Sub Command1_Click()
    Dim rsChartAccount                            As ADODB.Recordset
    Dim rsTitleCode                               As ADODB.Recordset
    Dim AccountTitles                             As String

    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Order by AcctCode asc")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        rsChartAccount.MoveFirst
        Do While Not rsChartAccount.EOF
            AccountTitles = Left(rsChartAccount!ACCTCODE, 2) & Mid(rsChartAccount!ACCTCODE, 4, 2)
            Set rsTitleCode = New ADODB.Recordset
            Set rsTitleCode = gconDMIS.Execute("Select * from AMIS_TitleCode where Code = '" & AccountTitles & "'")
            If rsTitleCode.EOF And rsTitleCode.BOF Then
                MsgBox "Invalid Title Code for Account Code : " & rsChartAccount!ACCTCODE
            End If
            rsChartAccount.MoveNext
        Loop
    End If
    MsgBox "Done"
End Sub
'
Private Sub Command2_Click()
'    Dim rsx                                            As ADODB.Recordset
'    Set rsx = New ADODB.Recordset
'    rsx.Fields.Append "VOUCHERNO", adVarChar, 10
'    rsx.Fields.Append "JTYPE", adVarChar, 10
'
'    rsx.Fields.Append "DEBIT", adDecimal
'    rsx.Fields("DEBIT").Precision = 18
'    rsx.Fields("DEBIT").NumericScale = 2
'
'    rsx.Fields.Append "CREDIT", adDecimal
'    rsx.Fields("CREDIT").Precision = 18
'    rsx.Fields("CREDIT").NumericScale = 2
'
'
'
'
'
'    SQX = SQX & "SELECT VOUCHERNO,JTYPE,  ISNULL(SUM(DEBIT),0)  DEBIT, SUM(CREDIT) CREDIT FROM AMIS_JOURNAL_DET WHERE STATUS='P'" & vbCrLf
'    SQX = SQX & "GROUP BY JTYPE,VOUCHERNO" & vbCrLf
'     Dim RSDATA As ADODB.Recordset
'    Set RSDATA = gconDMIS.Execute(SQX)
'rsx.Open
'    While Not RSDATA.EOF
'    rsx.AddNew
'    rsx.Fields(0) = RSDATA("VOUCHERNO") & ""
'    rsx.Fields(1) = RSDATA("JTYPE") & ""
'    rsx.Fields(2) = RSDATA("DEBIT") & ""
'    rsx.Fields(3) = RSDATA("CREDIT") & ""
'
'
'    rsx.Update
'        'gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT=" & NumericVal(rsx!DEBIT) & ", CREDIT=" & NumericVal(rsx!CREDIT) & " WHERE VOUCHERNO=" & N2Str2Null(rsx!VOUCHERNO) & " AND JTYPE=" & N2Str2Null(rsx!jtype))
'        'Me.Caption = rsx!VOUCHERNO & ""
'        RSDATA.MoveNext
'    Wend

End Sub


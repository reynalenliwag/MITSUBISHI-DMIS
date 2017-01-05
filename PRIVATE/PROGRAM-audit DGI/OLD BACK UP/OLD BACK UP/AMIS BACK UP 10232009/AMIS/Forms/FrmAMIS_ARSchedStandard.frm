VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form FrmAMIS_ARSchedStandard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Recievable Schedule"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7245
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picloading 
      Height          =   2925
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   4125
      TabIndex        =   2
      Top             =   810
      Width           =   4185
      Begin VB.Frame Frame 
         Caption         =   "Progress"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   120
         TabIndex        =   3
         Top             =   510
         Width           =   3885
         Begin wizProgBar.Prg progress 
            Height          =   285
            Left            =   90
            TabIndex        =   4
            Top             =   330
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   503
            Picture         =   "FrmAMIS_ARSchedStandard.frx":0000
            ForeColor       =   0
            BarPicture      =   "FrmAMIS_ARSchedStandard.frx":001C
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
         Begin VB.Label Label6 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2370
            TabIndex        =   12
            Top             =   1860
            Width           =   1425
         End
         Begin VB.Label Label5 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2370
            TabIndex        =   11
            Top             =   1500
            Width           =   855
         End
         Begin VB.Label Label4 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2370
            TabIndex        =   10
            Top             =   1140
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Total Transaction:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   1500
            Width           =   1965
         End
         Begin VB.Label Label2 
            Caption         =   "Status:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   1860
            Width           =   1665
         End
         Begin VB.Label Label1 
            Caption         =   "Transaction completed:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1110
            Width           =   2265
         End
         Begin VB.Label Label 
            Caption         =   "Percent compete:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   780
            Width           =   1665
         End
         Begin VB.Label lblpercent 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2370
            TabIndex        =   5
            Top             =   780
            Width           =   855
         End
      End
      Begin VB.Label lblprocess 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Top             =   150
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command 
      Caption         =   "TEST"
      Height          =   435
      Index           =   0
      Left            =   2850
      TabIndex        =   1
      Top             =   150
      Width           =   915
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   556
      _Version        =   393216
      Format          =   48562177
      CurrentDate     =   39875
   End
End
Attribute VB_Name = "FrmAMIS_ARSchedStandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command_Click(Index As Integer)
GeneratePayment_HD
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
End Sub
Sub TransaferInvoice_HD()
    Dim HDjtype As String
    Dim HDVoucherno As String
    Dim HDjdate As String
    Dim HDstatus As String
    Dim HDInvoiceno As String
    Dim HDinvoicetype As String
    Dim HDcustomercode As String
    Dim HDBalance As Double
    Dim DTBALANCE As Double
    Dim HDAmount2pay As Double
    Dim Validate_detail As New ADODB.Recordset
    Dim TotalDebit As Double
    Dim tmp_amount As Double
    Dim Cnt As Double
    
    Dim rsJournal_hd As New ADODB.Recordset
    Dim rsJournal_dt As New ADODB.Recordset
    Set rsJournal_hd = gconDMIS.Execute("SELECT VOUCHERNO,JTYPE,JDATE,CUSTOMERCODE,INVOICENO,INVOICETYPE,STATUS,invoiceAmt FROM " & _
                                      "AMIS_JOURNAL_HD WHERE (JTYPE = 'SJ' OR JTYPE = 'COB') and status = 'P' AND JDATE <='" & dtDate & "' order by voucherno ")
    
    lblprocess.Caption = "Validating SJ and COB transaction.."
    If Not (rsJournal_hd.EOF And rsJournal_hd.BOF) Then
        progress.Value = 0
        progress.Max = rsJournal_hd.RecordCount
        Cnt = 0
        Do While Not rsJournal_hd.EOF
            Cnt = Cnt + 1
            HDjtype = rsJournal_hd!jtype
            HDVoucherno = rsJournal_hd!VOUCHERNO
            HDjdate = rsJournal_hd!JDATE
            HDstatus = rsJournal_hd!Status
            HDInvoiceno = rsJournal_hd!invoiceno
            HDinvoicetype = rsJournal_hd!InvoiceType
            HDcustomercode = rsJournal_hd!CustomerCode
            If HDjtype = "COB" Then
                HDAmount2pay = NumericVal(rsJournal_hd!INVOICEAMT)
              Else
                HDAmount2pay = 0
            End If
            Set rsJournal_dt = gconDMIS.Execute("SELECT VOUCHERNO,JTYPE,DEBIT,CREDIT,ACCT_CODE FROM " & _
                                                      "AMIS_JOURNAL_DET WHERE JTYPE ='" & HDjtype & _
                                                      "' AND VOUCHERNO ='" & HDVoucherno & "' and left(ACCT_CODE,5) = '11-02' AND STATUS = 'P'")
                  
                  
                  If Not (rsJournal_dt.EOF And rsJournal_dt.BOF) Then
                         'Header
                         gconDMIS.Execute ("INSERT INTO AMIS_INVOICE_HD(VOUCHERNO,JDATE,STATUS,JTYPE,SJ_CUSTOMERCODE,INVOICETYPE,INVOICENO,amount2pay) " & _
                                   "VALUES('" & HDVoucherno & "','" & HDjdate & "','" & HDstatus & "','" & HDjtype & "','" & HDcustomercode & _
                                   "','" & HDinvoicetype & "','" & HDInvoiceno & "','" & HDAmount2pay & "')")
                        
                        Do While Not rsJournal_dt.EOF
                            If rsJournal_dt!jtype = "SJ" Then
                                   If rsJournal_dt!DEBIT = 0 Then
                                        DTBALANCE = NumericVal(rsJournal_dt!CREDIT)
                                     Else
                                        DTBALANCE = NumericVal(rsJournal_dt!DEBIT)
                                   End If
                               Else ' COB
                                   DTBALANCE = HDBalance
                            End If
                            'Detail
                            Set Validate_detail = gconDMIS.Execute("SELECT COUNT(*) from AMIS_INVOICE_DT WHERE voucherno='" & HDVoucherno & "' AND JTYPE = '" & HDjtype & "' AND ACCT_CODE='" & rsJournal_dt!ACCT_CODE & "' and customercode ='" & HDcustomercode & "'")
                            If Validate_detail(0) = 1 Then
                                TotalDebit = NumericVal(tmp_amount + DTBALANCE)
                                gconDMIS.Execute ("UPDATE AMIS_INVOICE_DT SET DEBIT='" & TotalDebit & "' WHERE VOUCHERNO ='" & HDVoucherno & "' AND JTYPE = '" & HDjtype & "'")
                                Else
                                gconDMIS.Execute ("INSERT INTO AMIS_INVOICE_DT(VOUCHERNO,CUSTOMERCODE,JTYPE,INVOICETYPE,INVOICENO,ACCT_CODE,DEBIT,CREDIT,BALANCE) " & _
                                              "VALUES('" & HDVoucherno & "','" & HDcustomercode & "','" & HDjtype & "','" & HDinvoicetype & _
                                              "','" & HDInvoiceno & "','" & rsJournal_dt!ACCT_CODE & "','" & NumericVal(rsJournal_dt!DEBIT) & _
                                              "','" & NumericVal(rsJournal_dt!CREDIT) & "','" & DTBALANCE & "')")
                                 
                                tmp_amount = DTBALANCE
                            End If
                            rsJournal_dt.MoveNext
                        Loop
                        TotalDebit = 0
                        tmp_amount = 0
                        If HDjtype = "SJ" Then
                            Sum_account HDVoucherno, HDjtype
                        End If
                         
                        
                End If
                
                DoEvents
                progress.Text = HDjtype + "-" + HDVoucherno
                progress.Value = progress.Value + 1
                lblpercent = Round((progress.Value / progress.Max * 100), 0) & "%"
                Label4.Caption = Cnt
                Label5.Caption = rsJournal_hd.RecordCount
                Label6.Caption = "In progress"
            rsJournal_hd.MoveNext
        Loop
    End If
    MsgBox "tapos "
    Set TransaferInvoice = Nothing
End Sub
Sub Sum_account(XVOUCHERNO As String, xJtype As String)
    'Sum the AR detail
    Dim amounttopay As Double
    Dim rsSumIt As New ADODB.Recordset
    Set rsSumIt = gconDMIS.Execute("SELECT SUM(DEBIT)AS TOTALDEBIT, SUM(CREDIT) AS TOTALCREDIT FROM AMIS_INVOICE_DT WHERE VOUCHERNO = '" & XVOUCHERNO & "' AND JTYPE = '" & xJtype & "'")
    If Not (rsSumIt.EOF And rsSumIt.BOF) Then
        If xJtype = "SJ" Then
            amounttopay = NumericVal((rsSumIt!TotalDebit) + (rsSumIt!Totalcredit))
            gconDMIS.Execute ("UPDATE AMIS_INVOICE_HD SET BALANCE='" & amounttopay & "',amount2pay ='" & amounttopay & "' WHERE VOUCHERNO ='" & XVOUCHERNO & "' AND JTYPE ='" & xJtype & "'")
        End If
    End If
    Set rsSumIt = Nothing
End Sub
Sub GeneratePayment_HD()
        Dim rsPJournal_hd As New ADODB.Recordset
        Dim rsPJournal_dt As New ADODB.Recordset
        Dim rsdetail As New ADODB.Recordset
        Dim ValidateHD As New ADODB.Recordset
        Dim HDjtype As String
        Dim HDVoucherno As String
        Dim HDjdate As String
        Dim HDstatus As String
        Dim HDcustomercode As String
        Dim HDORNO As String
        Dim HDInvoiceAmt As Double
        Dim Cnt As Double
        Set rsPJournal_hd = gconDMIS.Execute("SELECT VOUCHERNO,JTYPE,JDATE,CUSTOMERCODE,INVOICENO,INVOICETYPE,STATUS,invoiceAmt,refno FROM " & _
                                      "AMIS_JOURNAL_HD WHERE JTYPE = 'CRJ' and status = 'P' AND JDATE <='" & dtDate & "' order by voucherno ")
        'Header
        lblprocess.Caption = "Validating CRJ trasaction.."
        If Not (rsPJournal_hd.EOF And rsPJournal_hd.BOF) Then
             progress.Value = 0
             progress.Max = rsPJournal_hd.RecordCount
             Cnt = 0
            Do While Not rsPJournal_hd.EOF
                    Cnt = Cnt + 1
                    HDjtype = rsPJournal_hd!jtype
                    HDVoucherno = rsPJournal_hd!VOUCHERNO
                    HDjdate = rsPJournal_hd!JDATE
                    HDstatus = rsPJournal_hd!Status
                    HDcustomercode = rsPJournal_hd!CustomerCode
                    HDORNO = Null2String(rsPJournal_hd!refno)
                    HDInvoiceAmt = NumericVal(rsPJournal_hd!INVOICEAMT)
                    
                    
                    
                    Set rsPJournal_dt = gconDMIS.Execute("SELECT VOUCHERNO,JTYPE,DEBIT,CREDIT,ACCT_CODE FROM " & _
                                                      "AMIS_JOURNAL_DET WHERE JTYPE ='" & HDjtype & _
                                                      "'AND VOUCHERNO ='" & HDVoucherno & "' and left(ACCT_CODE,5) = '11-02' AND STATUS = 'P'")
                    'Detail
                    If Not (rsPJournal_dt.EOF And rsPJournal_dt.BOF) Then
                        ' Insert the header
                        Do While Not rsPJournal_dt.EOF
                            Set ValidateHD = gconDMIS.Execute("SELECT COUNT(*) FROM AMIS_PAYMENT_HD WHERE CRJ_VOUCHERNO='" & HDVoucherno & "' AND ACCT_CODE='" & rsPJournal_dt!ACCT_CODE & "'")
                            If ValidateHD(0) = 0 Then ' if allready exist do not commit
                            gconDMIS.Execute ("INSERT INTO AMIS_PAYMENT_HD(CRJ_VOUCHERNO,CRJ_DATE,STATUS,CRJ_CUSTOMERCODE,OR_NO,OR_AMOUNT,ACCT_CODE,debit,credit) VALUES('" & HDVoucherno & _
                                                 "','" & HDjdate & "','" & HDstatus & "','" & HDcustomercode & "','" & HDORNO & "','" & HDInvoiceAmt & "','" & rsPJournal_dt!ACCT_CODE & _
                                                 "','" & NumericVal(rsPJournal_dt!DEBIT) & "','" & NumericVal(rsPJournal_dt!CREDIT) & "')")
                            End If
                              rsPJournal_dt.MoveNext
                         Loop ' loop in account code
                            Set rsdetail = gconDMIS.Execute("SELECT VOUCHERNO,J_CLASS,INVOICETYPE,INVOICENO,INVOICEAMOUNT,Status,ID from AMIS_CRJ_DETAIL where voucherno='" & HDVoucherno & "'")
                            Do While Not rsdetail.EOF
                                'insert the detail
                                gconDMIS.Execute ("INSERT INTO AMIS_PAYMENT_DT(VOUCHERNO,CRJ_CUSTOMERCODE,INVOICEDATE,INVOICENO,INVOICETYPE,INVOICE_AMOUNT,jclass,status,detailID) " & _
                                                  "VALUES('" & HDVoucherno & _
                                                  "','" & HDcustomercode & "','" & HDjdate & "'," & N2Str2Null(rsdetail!invoiceno) & "," & N2Str2Null(rsdetail!InvoiceType) & _
                                                  ",'" & rsdetail!invoiceamount & "'," & N2Str2Null(rsdetail!j_class) & ",'P','" & rsdetail!ID & "')")
                                rsdetail.MoveNext
                            Loop ' loop in CRJ_DETAIL
                
                    End If
                      DoEvents
                      progress.Text = HDjtype + "-" + HDVoucherno
                      progress.Value = progress.Value + 1
                      lblpercent = Round((progress.Value / progress.Max * 100), 0) & "%"
                      Label4.Caption = Cnt
                      Label5.Caption = rsPJournal_hd.RecordCount
                      Label6.Caption = "In progress"
                      rsPJournal_hd.MoveNext
            Loop
        End If
        Set rsJournal_hd = Nothing
        MsgBox "s"
End Sub

                    

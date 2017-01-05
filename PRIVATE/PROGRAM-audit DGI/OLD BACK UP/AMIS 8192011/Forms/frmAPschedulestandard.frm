VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmAPschedulestandard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Payable Report"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   Icon            =   "frmAPschedulestandard.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picloading 
      Height          =   2925
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   0
      Width           =   4095
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   510
         Width           =   3885
         Begin wizProgBar.Prg progress 
            Height          =   285
            Left            =   90
            TabIndex        =   3
            Top             =   330
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   503
            Picture         =   "frmAPschedulestandard.frx":058A
            ForeColor       =   0
            BarPicture      =   "frmAPschedulestandard.frx":05A6
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
            TabIndex        =   11
            Top             =   780
            Width           =   855
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
            TabIndex        =   10
            Top             =   780
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
            TabIndex        =   9
            Top             =   1110
            Width           =   2265
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
            TabIndex        =   7
            Top             =   1500
            Width           =   1965
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
            TabIndex        =   6
            Top             =   1140
            Width           =   855
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
            TabIndex        =   5
            Top             =   1500
            Width           =   855
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
            TabIndex        =   4
            Top             =   1860
            Width           =   1425
         End
      End
      Begin VB.Label lblpro 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   90
         Width           =   3975
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
         Left            =   210
         TabIndex        =   12
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label 
         Caption         =   "As of :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   270
         TabIndex        =   17
         Top             =   120
         Width           =   1155
      End
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
      Left            =   1950
      MouseIcon       =   "frmAPschedulestandard.frx":05C2
      MousePointer    =   99  'Custom
      Picture         =   "frmAPschedulestandard.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Close Window"
      Top             =   990
      Width           =   885
   End
   Begin VB.TextBox Text 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   14
      Text            =   "Last Date generated :"
      Top             =   2580
      Width           =   4095
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   210
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtprocess 
      Height          =   405
      Left            =   840
      TabIndex        =   0
      Top             =   60
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   714
      _Version        =   393216
      Format          =   51707905
      CurrentDate     =   39882
   End
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
      Left            =   1080
      MouseIcon       =   "frmAPschedulestandard.frx":0B5F
      MousePointer    =   99  'Custom
      Picture         =   "frmAPschedulestandard.frx":0CB1
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Print Report"
      Top             =   990
      Width           =   885
   End
   Begin VB.Label Label 
      Caption         =   "As of "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   150
      Width           =   825
   End
End
Attribute VB_Name = "frmAPschedulestandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CDJ_date                                      As String
Dim xACCT_CODE                                    As String
Dim VendorCode                                    As String
Dim duedate                                       As String
Dim invoicedate                                   As String

Sub TransferAPAccount()
    Dim RSAPHD                                    As New ADODB.Recordset
    Dim validateTransacation                      As New ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJdate                                    As String
    Dim xJType                                    As String
    Dim xSTATUS                                   As String
    Dim xVENDORCODE                               As String
    Dim xINVOICETYPE                              As String
    Dim xInvoicedate                              As String
    Dim XINVOICEAMT                               As Double
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    Dim xAMOUNT2PAY                               As Double
    Dim xACCT_CODE                                As String
    Dim xDUEDATE                                  As String
    Dim cnt                                       As Integer
    gconDMIS.Execute ("delete from AMIS_AP_HD")
    Set RSAPHD = gconDMIS.Execute("SELECT  AMIS_Journal_HD.VoucherNo, AMIS_Journal_HD.JType,AMIS_Journal_HD.jdate,AMIS_Journal_HD.status,AMIS_Journal_HD.invoicetype,AMIS_Journal_HD.invoicedate, AMIS_Journal_HD.VendorCode,AMIS_Journal_HD.invoiceamt,AMIS_Journal_HD.amounttopay,AMIS_Journal_HD.amountpaid,AMIS_Journal_Det.Debit, AMIS_Journal_Det.CREDIT,AMIS_Journal_HD.DUEDATE,AMIS_Journal_Det.Acct_Code " & _
                                  "FROM AMIS_Journal_HD INNER JOIN AMIS_Journal_Det " & _
                                  "ON AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo AND AMIS_Journal_HD.jtype = AMIS_Journal_Det.jtype " & _
                                  "WHERE (AMIS_Journal_HD.JType = 'APJ' OR AMIS_Journal_HD.JType = 'VPJ' or AMIS_Journal_HD.JType = 'VDJ' or AMIS_Journal_HD.JType = 'VCJ') AND (LEFT(AMIS_Journal_Det.Acct_Code, 5) = '21-02' OR LEFT(AMIS_Journal_Det.Acct_Code, 5) = '21-01') and AMIS_Journal_hd.status='P' and AMIS_Journal_hd.jdate < ='" & dtprocess.Value & "'")

    If Not (RSAPHD.EOF And RSAPHD.BOF) Then
        cnt = 0
        lblpro = "Validating Transaction as of " & ":" & dtprocess.Value
        progress.Value = 0
        progress.Max = RSAPHD.RecordCount

        Do While Not RSAPHD.EOF
            cnt = cnt + 1
            xVOUCHERNO = N2Str2Null(RSAPHD!VOUCHERNO)
            xJdate = N2Date2Null(RSAPHD!JDate)
            xJType = N2Str2Null(Trim(RSAPHD!jtype))
            xSTATUS = N2Str2Null(RSAPHD!Status)
            xVENDORCODE = N2Str2Null(RSAPHD!VendorCode)
            xINVOICETYPE = N2Str2Null(RSAPHD!InvoiceType)
            xInvoicedate = N2Date2Null(RSAPHD!invoicedate)
            XINVOICEAMT = NumericVal(RSAPHD!InvoiceAmt)
            xdebit = NumericVal(RSAPHD!DEBIT)
            xcredit = NumericVal(RSAPHD!CREDIT)
            xDUEDATE = N2Date2Null(RSAPHD!duedate)
            If RSAPHD!jtype = "VPJ" Or Trim(RSAPHD!jtype) = "VDJ" Then
                ' opening
                xAMOUNT2PAY = NumericVal(RSAPHD!amounttopay)
            ElseIf RSAPHD!jtype = "APJ" Then
                If xdebit = 0 Then
                    xAMOUNT2PAY = xcredit
                Else
                    xAMOUNT2PAY = xdebit
                End If
            ElseIf RSAPHD!jtype = "VCJ" Then
                xAMOUNT2PAY = NumericVal(RSAPHD!AMOUNTPAID)
            End If


            xACCT_CODE = N2Str2Null(RSAPHD!Acct_code)
            Set validateTransacation = gconDMIS.Execute("SELECT COUNT(*) FROM AMIS_AP_HD WHERE VOUCHERNO ='" & RSAPHD!VOUCHERNO & "' AND JTYPE = '" & RSAPHD!jtype & "'")
            If (CDate(RSAPHD!JDate) <= dtprocess.Value) Then
                If validateTransacation(0) = 0 Then        ' more than 1
                    gconDMIS.Execute ("INSERT INTO AMIS_AP_HD(VOUCHERNO,JDATE,jtype,STATUS,VENDOR_CODE,INVOICETYPE,INVOICEDATE,INVOICEAMT,DEBIT,CREDIT,AMOUNT2PAY,ACCT_CODE,duedate)" & _
                                      " VALUES(" & xVOUCHERNO & "," & xJdate & "," & Trim(xJType) & "," & xSTATUS & _
                                      "," & xVENDORCODE & "," & xINVOICETYPE & _
                                      "," & xInvoicedate & "," & XINVOICEAMT & _
                                      "," & xdebit & "," & xcredit & _
                                      "," & xAMOUNT2PAY & "," & xACCT_CODE & _
                                      "," & xDUEDATE & ")")
                End If
            End If

            DoEvents
            progress.Text = RSAPHD!jtype + "-" + RSAPHD!VOUCHERNO
            progress.Value = progress.Value + 1
            lblPercent = Round((progress.Value / progress.Max * 100), 0) & "%"
            Label4.Caption = cnt
            Label5.Caption = RSAPHD.RecordCount
            Label6.Caption = "In progress"
            RSAPHD.MoveNext
        Loop
    End If
    Set RSAPHD = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim thedate                                   As String
    Dim Ans                                       As String
    Dim RS                                        As New ADODB.Recordset
    Ans = MsgBox("Do you want to process Accounts Payable Aging report? ", vbQuestion + vbYesNo)
    If Ans = vbYes Then
        processAP
        Set RS = gconDMIS.Execute("Select lastupdated from AMIS_AP")
        If Not (RS.EOF And Not RS.BOF) Then
            thedate = Null2String(RS!lastupdated)
            Text.Text = Text.Text & thedate
        End If
        CrystalReport1.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
        CrystalReport1.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        CrystalReport1.WindowTitle = "Accounts Payable Aging Report  AS OF: " & thedate
        CrystalReport1.ReportTitle = "Accounts Payable Aging Report AS OF: " & thedate
        PrintSQLReport CrystalReport1, AMIS_REPORT_PATH & "DueReports\AP_AgingReport.Rpt", "", DMIS_REPORT_Connection, 1

        LogAudit "V", "ACCOUNTS PAYABLE AGING REPORT", "As of: " & thedate
    Else
        Set RS = gconDMIS.Execute("Select lastupdated from AMIS_AP")
        If Not (RS.EOF And Not RS.BOF) Then
            thedate = Null2String(RS!lastupdated)
            Text.Text = Text.Text & thedate
        End If
        CrystalReport1.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
        CrystalReport1.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        CrystalReport1.WindowTitle = "Accounts Payable Aging Report  AS OF: " & thedate
        CrystalReport1.ReportTitle = "Accounts Payable Aging Report AS OF: " & thedate
        PrintSQLReport CrystalReport1, AMIS_REPORT_PATH & "DueReports\AP_AgingReport.Rpt", "", DMIS_REPORT_Connection, 1
        LogAudit "V", "ACCOUNTS PAYABLE AGING REPORT", "As of: " & thedate
    End If
End Sub



Private Sub Form_Load()
    Dim Y                                         As New ADODB.Recordset
    Dim X                                         As String
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
    Set Y = gconDMIS.Execute("Select lastupdated from AMIS_AP")
    If Not (Y.EOF And Y.BOF) Then
        X = Null2String(Y!lastupdated)
        Text.Text = Text.Text & X
    End If
    picLoading.Visible = False
End Sub
Sub processAP()
    Dim RS_APHD                                   As New ADODB.Recordset
    Dim RSCVDETAIL                                As New ADODB.Recordset
    Dim totalpayment                              As Double
    Dim HDAmount2pay
    Dim BALANCE                                   As Double
    Dim HDVoucherno                               As String
    Dim CDJvoucherno                              As String
    Dim xysystemRemark                            As String
    Dim xsystemRemark                             As String
    Dim Reference                                 As String
    Dim xReference                                As String
    Dim cnt                                       As Double
    picLoading.Visible = True
    gconDMIS.Execute ("update AMIS_CV_DETAIL set status = 'P'")
    gconDMIS.Execute ("delete from AMIS_AP")
    TransferAPAccount
    Set RS_APHD = gconDMIS.Execute("select * from AMIS_AP_HD")
    If Not (RS_APHD.EOF And RS_APHD.BOF) Then
        lblpro = "Processing Transaction.."
        progress.Value = 0
        progress.Max = RS_APHD.RecordCount
        cnt = 0
        Do While Not RS_APHD.EOF
            cnt = cnt + 1
            HDAmount2pay = NumericVal(RS_APHD!AMOUNT2PAY)
            BALANCE = HDAmount2pay
            Reference = (Trim(RS_APHD!jtype) + "-" + RS_APHD!VOUCHERNO)

            If Trim(RS_APHD!jtype) = "VCJ" Or Trim(RS_APHD!jtype) = "VDJ" Then
                '
                totalpayment = 0
                If Trim(RS_APHD!jtype) = "VCJ" Then
                    'BALANCE = RS_APHD!AMOUNT2PAY * (-1)
                    BALANCE = 0
                Else
                    BALANCE = 0
                    'BALANCE = RS_APHD!AMOUNT2PAY * (-1)
                End If
                GoTo SaveAdjustment:
                'gconDMIS.Execute ("INSERT INTO AMIS_AP(VOUCHERNO,CDJ_VOUCHERNO,VENDOR_CODE,DUEDATE,INVOICEDATE,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,SYSTEMREMARK)" & _
                 '                               "VALUES('" & reference & _
                 '                               "','" & Null & _
                 '                               "','" & RS_APHD!vendor_code & _
                 '                               "','" & RS_APHD!duedate & _
                 '                               "','" & RS_APHD!Invoicedate & _
                 '                               "','" & HDAmount2pay & _
                 '                               "','" & TotalPayment & _
                 '                               "','" & BALANCE & _
                 '                               "','" & RS_APHD!ACCT_CODE & _
                 '                               "','" & xsystemRemark & "')")

            Else
                Set RSCVDETAIL = gconDMIS.Execute("SELECT VOUCHERNO,PV_VOUCHERNO,AMOUNT,STATUS,ID FROM AMIS_CV_DETAIL " & _
                                                  "WHERE PV_VOUCHERNO = '" & RS_APHD!VOUCHERNO & "' AND STATUS <>'Y'")

                CDJvoucherno = N2Str2Null("")
                If Not (RSCVDETAIL.EOF And RSCVDETAIL.BOF) Then
                    'there is a payment
                    CDJvoucherno = N2Str2Null(RSCVDETAIL!VOUCHERNO)

                    xsystemRemark = "Good"
                    Do While Not RSCVDETAIL.EOF
                        If ValidateCDJ(RSCVDETAIL!VOUCHERNO) = "P" And CDate(CDJ_date) <= dtprocess.Value Then
                            If Trim(RS_APHD!VENDOR_CODE) = GetPayeeCode(CDJvoucherno) Then
                                CDJvoucherno = N2Str2Null(RSCVDETAIL!VOUCHERNO)
                                totalpayment = totalpayment + NumericVal(RSCVDETAIL!amount)
                                gconDMIS.Execute ("Update AMIS_CV_DETAIL SET STATUS = 'Y' WHERE ID='" & RSCVDETAIL!ID & "'")
                            End If
                        End If
                        ' wrong vendor
                        If Trim(RS_APHD!VENDOR_CODE) <> GetPayeeCode(CDJvoucherno) And CDate(CDJ_date) <= dtprocess.Value Then
                            xysystemRemark = "WCC"
                            xReference = "CDJ-" + RSCVDETAIL!VOUCHERNO
                            totalpayment = 0
                            BALANCE = RS_APHD!AMOUNT2PAY * (-1)
                            gconDMIS.Execute ("INSERT INTO AMIS_AP(VOUCHERNO,CDJ_VOUCHERNO,VENDOR_CODE,DUEDATE,INVOICEDATE,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,SYSTEMREMARK)" & _
                                              "VALUES('" & xReference & _
                                              "'," & CDJvoucherno & _
                                              ",'" & GetPayeeCode(CDJvoucherno) & _
                                              "','" & RS_APHD!duedate & _
                                              "','" & RS_APHD!invoicedate & _
                                              "','" & HDAmount2pay & _
                                              "','" & totalpayment & _
                                              "','" & BALANCE & _
                                              "','" & RS_APHD!Acct_code & _
                                              "','" & xysystemRemark & "')")
                        End If
                        RSCVDETAIL.MoveNext
                    Loop

                    BALANCE = HDAmount2pay - totalpayment
                Else
                    ' no payment
                End If
            End If
            DoEvents
SaveAdjustment:

            gconDMIS.Execute ("INSERT INTO AMIS_AP(VOUCHERNO,CDJ_VOUCHERNO,VENDOR_CODE,DUEDATE,INVOICEDATE,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,SYSTEMREMARK)" & _
                              "VALUES('" & Reference & _
                              "'," & N2Str2Null(CDJvoucherno) & _
                              ",'" & RS_APHD!VENDOR_CODE & _
                              "','" & RS_APHD!duedate & _
                              "','" & RS_APHD!invoicedate & _
                              "','" & HDAmount2pay & _
                              "','" & totalpayment & _
                              "','" & BALANCE & _
                              "','" & RS_APHD!Acct_code & _
                              "'," & N2Str2Null(xsystemRemark) & ")")

            totalpayment = 0
            BALANCE = 0
            CDJ_date = ""
            progress.Text = Trim(RS_APHD!jtype) + "-" + Trim(RS_APHD!VOUCHERNO)
            progress.Value = progress.Value + 1
            lblPercent = Round((progress.Value / progress.Max * 100), 0) & "%"
            Label4.Caption = cnt
            Label5.Caption = RS_APHD.RecordCount
            Label6.Caption = "In progress"
            RS_APHD.MoveNext
        Loop
        CDJ_NOLINK
    End If
    DirectPayment
    gconDMIS.Execute ("Update AMIS_AP set lastupdated='" & dtprocess & "'")
    'MsgBox "You can now generate AP schedule..", vbInformation, "Process completed"
    picLoading.Visible = False
    Set RSCVDETAIL = Nothing
    Set RS_APHD = Nothing
End Sub
Function GetPayeeCode(xVOUCHERNO As String) As String
    Dim RsPayee                                   As New ADODB.Recordset
    Set RsPayee = gconDMIS.Execute("SELECT VENDORCODE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO=" & xVOUCHERNO & " and jtype = 'CDJ'")
    If Not (RsPayee.EOF And RsPayee.BOF) Then
        GetPayeeCode = N2String(RsPayee!VendorCode)
    Else
        GetPayeeCode = ""
    End If
    Set RsPayee = Nothing
End Function
Sub CDJ_NOLINK()
    Dim RSCDJ                                     As New ADODB.Recordset
    Dim RC_CV                                     As New ADODB.Recordset
    Dim HDAmount2pay                              As Double
    Dim totalpayment                              As Double
    Dim BALANCE                                   As Double
    Dim cnt                                       As Double
    Dim xVOUCHERNO                                As String
    Dim Reference                                 As String
    gconDMIS.Execute ("DELETE from AMIS_AP where systemremark ='NL'")
    Set RSCDJ = gconDMIS.Execute("SELECT  AMIS_Journal_HD.VoucherNo, AMIS_Journal_HD.JType,AMIS_Journal_HD.jdate,AMIS_Journal_HD.status,AMIS_Journal_HD.invoicetype,AMIS_Journal_HD.invoicedate, AMIS_Journal_HD.VendorCode,AMIS_Journal_HD.invoiceamt,AMIS_Journal_HD.amounttopay,AMIS_Journal_HD.amountpaid,AMIS_Journal_Det.Debit, AMIS_Journal_Det.CREDIT,AMIS_Journal_HD.DUEDATE,AMIS_Journal_Det.Acct_Code " & _
                                 "FROM AMIS_Journal_HD INNER JOIN AMIS_Journal_Det " & _
                                 "ON AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo AND AMIS_Journal_HD.jtype = AMIS_Journal_Det.jtype " & _
                                 "WHERE (AMIS_Journal_HD.JType = 'CDJ') AND (LEFT(AMIS_Journal_Det.Acct_Code, 5) = '21-02' OR LEFT(AMIS_Journal_Det.Acct_Code, 5) = '21-01') and AMIS_Journal_hd.status='P' and AMIS_Journal_hd.jdate < ='" & dtprocess.Value & "'")

    If Not (RSCDJ.EOF And RSCDJ.BOF) Then
        lblpro = "Processing CDJ with no link.."
        progress.Value = 0
        progress.Max = RSCDJ.RecordCount
        cnt = 0
        Do While Not RSCDJ.EOF
            cnt = cnt + 1
            xVOUCHERNO = RSCDJ!VOUCHERNO
            Reference = "CDJ-" + RSCDJ!VOUCHERNO
            HDAmount2pay = 0
            totalpayment = 0
            If RSCDJ!DEBIT = 0 Then
                BALANCE = RSCDJ!CREDIT
            Else
                BALANCE = RSCDJ!DEBIT * (-1)
            End If
            Set RC_CV = gconDMIS.Execute("SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE VOUCHERNO='" & xVOUCHERNO & "'")
            If RC_CV.EOF And RC_CV.BOF Then

                gconDMIS.Execute ("INSERT INTO AMIS_AP(VOUCHERNO,CDJ_VOUCHERNO,VENDOR_CODE,DUEDATE,INVOICEDATE,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,SYSTEMREMARK)" & _
                                  "VALUES('" & Reference & _
                                  "'," & RSCDJ!VOUCHERNO & _
                                  ",'" & RSCDJ!VendorCode & _
                                  "','" & RSCDJ!duedate & _
                                  "','" & RSCDJ!invoicedate & _
                                  "','" & HDAmount2pay & _
                                  "','" & totalpayment & _
                                  "','" & BALANCE & _
                                  "','" & RSCDJ!Acct_code & _
                                  "','NL')")
            End If
            DoEvents
            progress.Text = Trim(RSCDJ!jtype) + "-" + Trim(RSCDJ!VOUCHERNO)
            progress.Value = progress.Value + 1
            lblPercent = Round((progress.Value / progress.Max * 100), 0) & "%"
            Label4.Caption = cnt
            Label5.Caption = RSCDJ.RecordCount
            Label6.Caption = "In progress"
            RSCDJ.MoveNext
        Loop
    End If
    Set RSCDJ = Nothing
End Sub
Function ValidateCDJ(xVOUCHERNO As String)
    Dim CDJ                                       As New ADODB.Recordset
    Set CDJ = gconDMIS.Execute("SELECT STATUS,jdate,vendorcode,duedate,invoicedate FROM AMIS_JOURNAL_HD WHERE VOUCHERNO=" & xVOUCHERNO & " AND JTYPE = 'CDJ'")
    If Not (CDJ.EOF And CDJ.BOF) Then
        ValidateCDJ = Null2String(CDJ!Status)
        CDJ_date = (CDJ!JDate)
        VendorCode = Null2String(CDJ!VendorCode)
        duedate = Null2String(CDJ!duedate)
        invoicedate = Null2String(CDJ!invoicedate)
    End If
    Set CDJ = Nothing
End Function
Sub DirectPayment()
' WITH cdj BUT THE apj IS NOT YET CREATED
    Dim rsCV_Detail                               As New ADODB.Recordset
    Dim rsAPacct                                  As New ADODB.Recordset
    Dim rsHDCDJ                                   As New ADODB.Recordset
    Dim cnt                                       As Double
    Dim HDAmount2pay                              As Double
    Dim totalpayment                              As Double
    Dim Reference                                 As String
    Dim BALANCE                                   As Double
    Set rsCV_Detail = gconDMIS.Execute("SELECT VOUCHERNO,PV_VOUCHERNO,amount FROM AMIS_CV_DETAIL WHERE STATUS <>'Y'")
    lblpro.Caption = "Finalizing Data.."
    If Not (rsCV_Detail.EOF And rsCV_Detail.BOF) Then
        cnt = 0
        progress.Value = 0
        progress.Max = rsCV_Detail.RecordCount
        Reference = "CDJ-" & rsCV_Detail!VOUCHERNO
        Do While Not rsCV_Detail.EOF
            cnt = cnt + 1
            If ValidateCDJ(rsCV_Detail!VOUCHERNO) = "P" And CDate(CDJ_date) <= dtprocess.Value And isAPaccount(rsCV_Detail!VOUCHERNO) = True Then
                Set rsHDCDJ = gconDMIS.Execute("SELECT  AMIS_Journal_HD.VoucherNo, AMIS_Journal_HD.JType,AMIS_Journal_HD.jdate,AMIS_Journal_HD.status,AMIS_Journal_HD.invoicetype,AMIS_Journal_HD.invoicedate, AMIS_Journal_HD.VendorCode,AMIS_Journal_HD.invoiceamt,AMIS_Journal_HD.amounttopay,AMIS_Journal_HD.amountpaid,AMIS_Journal_Det.Debit, AMIS_Journal_Det.CREDIT,AMIS_Journal_HD.DUEDATE,AMIS_Journal_Det.Acct_Code " & _
                                               "FROM AMIS_Journal_HD INNER JOIN AMIS_Journal_Det " & _
                                               "ON AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo AND AMIS_Journal_HD.jtype = AMIS_Journal_Det.jtype " & _
                                               "WHERE (AMIS_Journal_HD.JType = 'APJ' OR AMIS_Journal_HD.JType = 'VPJ') AND (LEFT(AMIS_Journal_Det.Acct_Code, 5) = '21-02' OR LEFT(AMIS_Journal_Det.Acct_Code, 5) = '21-01') and AMIS_Journal_hd.status='P' and AMIS_Journal_hd.voucherno = '" & rsCV_Detail!pv_voucherno & "' and AMIS_Journal_hd.jdate < ='" & dtprocess.Value & "'")

                If (rsHDCDJ.EOF And rsHDCDJ.BOF) Then
                    ' not yet created
                    HDAmount2pay = 0
                    totalpayment = rsCV_Detail!amount
                    BALANCE = rsCV_Detail!amount * (-1)
                    gconDMIS.Execute ("INSERT INTO AMIS_AP(VOUCHERNO,CDJ_VOUCHERNO,VENDOR_CODE,DUEDATE,INVOICEDATE,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,SYSTEMREMARK)" & _
                                      "VALUES('" & Reference & _
                                      "'," & N2Str2Null("") & _
                                      ",'" & VendorCode & _
                                      "','" & duedate & _
                                      "','" & invoicedate & _
                                      "','" & HDAmount2pay & _
                                      "','" & totalpayment & _
                                      "','" & BALANCE & _
                                      "','" & xACCT_CODE & _
                                      "','ADP')")
                End If
            End If
            DoEvents
            progress.Text = Trim(rsCV_Detail!VOUCHERNO)
            progress.Value = progress.Value + 1
            lblPercent = Round((progress.Value / progress.Max * 100), 0) & "%"
            Label4.Caption = cnt
            Label5.Caption = rsCV_Detail.RecordCount
            Label6.Caption = "In progress"
            rsCV_Detail.MoveNext
        Loop
    End If
    MsgBox "Process completed..", vbInformation, "Information"
    Set rsCV_Detail = Nothing
End Sub
Function isAPaccount(xVOUCHERNO As String) As Boolean
    Dim rsCDJx                                    As New ADODB.Recordset
    'set rsCDJx = gconDMIS.Execute("SELECT  AMIS_Journal_HD.VoucherNo, AMIS_Journal_HD.JType,AMIS_Journal_HD.jdate,AMIS_Journal_HD.status,AMIS_Journal_HD.invoicetype,AMIS_Journal_HD.invoicedate, AMIS_Journal_HD.VendorCode,AMIS_Journal_HD.invoiceamt,AMIS_Journal_HD.amounttopay,AMIS_Journal_HD.amountpaid,AMIS_Journal_Det.Debit, AMIS_Journal_Det.CREDIT,AMIS_Journal_HD.DUEDATE,AMIS_Journal_Det.Acct_Code " & _
     '                               "FROM AMIS_Journal_HD INNER JOIN AMIS_Journal_Det " & _
     '                               "ON AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo AND AMIS_Journal_HD.jtype = AMIS_Journal_Det.jtype " & _
     '                               "WHERE (AMIS_Journal_HD.JType = 'CDJ') AND (LEFT(AMIS_Journal_Det.Acct_Code, 5) = '21-02' OR LEFT(AMIS_Journal_Det.Acct_Code, 5) = '21-01') and AMIS_Journal_hd.status='P' and AMIS_Journal_hd.voucherno = '" & XVOUCHERNO & "' and AMIS_Journal_hd.jdate < ='" & dtprocess.Value & "'")

    Set rsCDJx = gconDMIS.Execute("SELECT VOUCHERNO,ACCT_CODE FROM AMIS_JOURNAL_DET " & _
                                  "WHERE (JTYPE = 'CDJ') AND (LEFT(Acct_Code, 5) = '21-02' OR LEFT(Acct_Code, 5) = '21-01') and status='P' and voucherno = '" & xVOUCHERNO & "' and jdate < ='" & dtprocess.Value & "'")

    If Not (rsCDJx.EOF And rsCDJx.BOF) Then
        isAPaccount = True
        xACCT_CODE = rsCDJx!Acct_code
    Else
        xACCT_CODE = ""
        isAPaccount = False
    End If

    Set rsCDJx = Nothing
End Function

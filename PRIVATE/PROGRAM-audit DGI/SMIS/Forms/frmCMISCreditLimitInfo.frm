VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCMISCreditLimitInfo 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Limit Info"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "frmCMISCreditLimitInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCMISCreditLimitInfo.frx":39050
   ScaleHeight     =   2055
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   4320
      TabIndex        =   33
      Top             =   90
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox Picture1 
      Height          =   465
      Left            =   4320
      ScaleHeight     =   405
      ScaleWidth      =   2625
      TabIndex        =   32
      Top             =   90
      Width           =   2685
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "ViewUnpaid"
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdViewP 
      Caption         =   "ViewPastDue"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCIOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      Picture         =   "frmCMISCreditLimitInfo.frx":720A0
      TabIndex        =   13
      Top             =   3000
      Width           =   615
   End
   Begin VB.Frame picCrDetails 
      Height          =   2415
      Left            =   0
      TabIndex        =   11
      Top             =   3360
      Width           =   3855
      Begin MSComctlLib.ListView lstDetails 
         Height          =   2295
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Invoice no."
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Invoicedate"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Invoiceamount"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   2293
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7050
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCMISCreditLimitInfo.frx":B4A65
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame picOverride 
      Height          =   2295
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox cboUsers 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   23
         Text            =   "Combo1"
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Height          =   375
         Left            =   2040
         Picture         =   "frmCMISCreditLimitInfo.frx":B4CE3
         TabIndex        =   22
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2880
         TabIndex        =   21
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -360
         TabIndex        =   27
         Top             =   120
         Width           =   4215
      End
      Begin VB.Label lblId 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblUsercode 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.Frame PicCrInfo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox txtlevel 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   5160
         TabIndex        =   31
         Top             =   90
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.PictureBox Crypto1 
         Height          =   465
         Left            =   5280
         ScaleHeight     =   405
         ScaleWidth      =   2625
         TabIndex        =   30
         Top             =   90
         Width           =   2685
      End
      Begin VB.TextBox txtInvAmt 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   1440
         TabIndex        =   18
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtCreditAvail 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   1440
         TabIndex        =   16
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtCRTerm 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   1440
         TabIndex        =   9
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtCRAvailable 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtCrLimit 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtCrExpiry 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   2520
         Width           =   255
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4770
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCMISCreditLimitInfo.frx":C42C2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available Credit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Expiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Terms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unpaid amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Past Due"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmCMISCreditLimitInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xCUSTNAME As String
Dim XBALANCE As Double
Dim XMod_name As String
Dim xView As String
Dim override As String
Private Sub cmdCIOk_Click()
    Unload Me
'    frmSMIS_Trans_SalesOrder.cboSalesAE.SetFocus

End Sub

Private Sub cmdSelect_Click()
    Dim rsuserspass As New ADODB.Recordset
    Dim A As String
    With wizVar
    rsuserspass.Open "select * from all_rams_users where user_name = '" & cboUsers.Text & "'", gconDMIS
    If Not rsuserspass.EOF And Not rsuserspass.BOF Then
        If txtPass.Text = .DecryptAccess(rsuserspass!Password) Then
            lblUsercode.Caption = Null2String(rsuserspass!USERCODE)
            lblId.Caption = Null2String(rsuserspass!USERID)
        Else
            lblUsercode.Caption = ""
            lblId.Caption = 0
        End If
    End If
    End With
    Me.Hide
End Sub

Private Sub cmdview_Click()
    xView = "I"
    If picCrDetails.Visible = False Then
        picCrDetails.Visible = True
        Me.Height = 6195
        FillLstDet
    Else
        picCrDetails.Visible = False
        Me.Height = 3810
    End If
End Sub

Private Sub cmdViewP_Click()
    xView = "P"
    If picCrDetails.Visible = False Then
        picCrDetails.Visible = True
        Me.Height = 6195
        FillLstDet
    Else
        picCrDetails.Visible = False
        Me.Height = 3810
    End If
End Sub



Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If override = "Y" Then
        Me.Height = 2670
        Me.Caption = "Account Override verification."
        picoverride.Visible = True
        Call fillcbo
    Else
        frmCMISCreditLimitInfo.picoverride.Visible = False
        frmCMISCreditLimitInfo.Height = 3810
        frmCMISCreditLimitInfo.txtCRAvailable = ""
        frmCMISCreditLimitInfo.txtCrExpiry = ""
        frmCMISCreditLimitInfo.txtCrLimit = ""
        frmCMISCreditLimitInfo.Check1.Value = 0
        Dim RsCreditLimit As New ADODB.Recordset
        frmCMISCreditLimitInfo.PicCrInfo.Visible = True
        Dim xpayments As Double
        Dim rsRequirements As New ADODB.Recordset
        If XMod_name = "SMIS" Then ' If vehicle invoice
            Set rsRequirements = gconDMIS.Execute("SELECT *,DATEADD(month,+(terms_expiry_year * 12)+terms_expiry_month, credit_start) as expiry FROM ALL_CUSTOMER_TERMS WHERE customercode in (select cuscde from all_customer where acctname = '" & xCUSTNAME & "')")
            If Not (rsRequirements.EOF Or rsRequirements.BOF) Then
            '===========================================
            'credit limit
            Set RsCreditLimit = gconDMIS.Execute("SELECT *,DATEADD(month,+(terms_expiry_year * 12)+terms_expiry_month, credit_start) as expiry FROM ALL_CUSTOMER_TERMS WHERE customercode in (select cuscde from all_customer where acctname = '" & xCUSTNAME & "')")
            If Not (RsCreditLimit.EOF Or RsCreditLimit.BOF) Then
        '       frmCMISCreditLimitInfo.PicCrInfo.Visible = True
                frmCMISCreditLimitInfo.txtCrExpiry.Text = RsCreditLimit!expiry
                frmCMISCreditLimitInfo.txtCrLimit.Text = ToDoubleNumber(RsCreditLimit!CREDITLIMIT)
                frmCMISCreditLimitInfo.txtCRTerm.Text = RsCreditLimit!CREDITTERM
            End If
            'credit available
            Set RsCreditLimit = gconDMIS.Execute("select isnull(sum(baltofinanced),0)-(select isnull(sum(payment),0) from cmis_off_dt where trantype = 'VI' and paidna = 1 and cuscde = (select cuscde from all_customer where acctname ='" & xCUSTNAME & "')) available,(select isnull(sum(payment),0) from cmis_off_dt where trantype = 'VI' and paidna = 1 and cuscde = (select cuscde from all_customer where acctname ='" & xCUSTNAME & "')) payment  From SMIS_SalesOrder WHERE STATUS <> 'C' OR STATUS IS NULL and financingcode in (select cuscde from all_customer where acctname = '" & xCUSTNAME & "') and vi_no is not null")
            If RsCreditLimit!Available <> 0 Then
                'balance
                frmCMISCreditLimitInfo.txtCRAvailable.Text = ToDoubleNumber(RsCreditLimit!Available)
                xpayments = ToDoubleNumber(RsCreditLimit!PAYMENT)
            Else
                'balance
                frmCMISCreditLimitInfo.txtCRAvailable.Text = 0  'ToDoubleNumber(frmCMISCreditLimitInfo.txtCrLimit.Text)
            End If
            'past due
            Set RsCreditLimit = gconDMIS.Execute("select ar.customercode,ar.invoicetype,ar.invoiceno,ar.amount_topay,dt.invoiceamount from amis_ar ar inner join amis_detail dt on ar.InvoiceType = dt.InvoiceType " & _
            " And ar.INVOICENO = dt.INVOICENO And ar.CustomerCode = dt.CustomerCode And ar.account_code = dt.acct_code " & _
            " inner join all_customer_terms tr on ar.customercode=tr.customercode " & _
            " Where DateDiff(Day, ar.invoicedate, getdate()) > tr.creditterm " & _
            " and ar.amount_topay>dt.invoiceamount and ar.invoicetype = 'VI' " & _
            " and ar.customercode = (select top 1 cuscde from all_customer where acctname='" & xCUSTNAME & "') " & _
            " and ar.status='P' and dt.status='P' " & _
            " Union " & _
            " select ar.customercode,ar.invoicetype,invoiceno,amount_topay,0 as invoiceamount from amis_ar ar inner join all_customer_terms tr on ar.customercode=tr.customercode where ar.invoicetype = 'VI' and status = 'P' " & _
            " and invoiceno not in (select invoiceno from amis_detail where amis_detail.invoicetype = 'VI' and status = 'P'and customercode = (select top 1 cuscde from all_customer where acctname='" & xCUSTNAME & "') ) and ar.customercode = (select top 1 cuscde from all_customer where acctname='" & xCUSTNAME & "') and DateDiff(Day, ar.invoicedate, getdate()) > tr.creditterm")
           
            If Not (RsCreditLimit.EOF Or RsCreditLimit.BOF) Then
                frmCMISCreditLimitInfo.Check1.Value = 1
                frmCMISCreditLimitInfo.cmdViewP.Visible = True
            End If
            'expiry
            Set RsCreditLimit = gconDMIS.Execute("select DATEADD(month,(terms_expiry_year * 12)+terms_expiry_month, getdate()) as expiry,getdate() present from all_customer_terms where customercode = (select top 1 cuscde from all_customer where acctname='" & xCUSTNAME & "')")
            If Not (RsCreditLimit.EOF Or RsCreditLimit.BOF) Then
                If RsCreditLimit!expiry < RsCreditLimit!present Then
                    MessagePop InfoWarning, "Credit Limit Expired", "Credit Limit for this customer has expired! Please check your masterfile first.", 1000, 2
                End If
            End If
        
            'Limit
            Dim xxlimit As Double
            Dim xxbalance As Double
            xxlimit = frmCMISCreditLimitInfo.txtCrLimit.Text
            xxbalance = ToDoubleNumber(frmCMISCreditLimitInfo.txtCRAvailable.Text)
            txtCreditAvail.Text = xxlimit - ToDoubleNumber(xxbalance)
            txtCreditAvail.Text = ToDoubleNumber(NumericVal(txtCreditAvail.Text))
            txtInvAmt.Text = ToDoubleNumber(XBALANCE)
            
            frmCMISCreditLimitInfo.picCrDetails.Visible = False
            frmCMISCreditLimitInfo.PicCrInfo.Enabled = False
            If frmCMISCreditLimitInfo.Check1.Value = 1 Then
        '        frmCMISCreditLimitInfo.PicCrInfo.Enabled = True
            End If
            If N2Str2Zero(frmCMISCreditLimitInfo.txtCRAvailable.Text) + ToDoubleNumber(XBALANCE) > ToDoubleNumber(frmCMISCreditLimitInfo.txtCreditAvail.Text) Then
                frmCMISCreditLimitInfo.cmdView.Visible = True
        '        frmCMISCreditLimitInfo.PicCrInfo.Enabled = True
            End If
exiti:
        '    Unload Me
            Set RsCreditLimit = Nothing
            '======================================================
            Else
                frmCMISCreditLimitInfo.PicCrInfo.Enabled = False
                MessagePop InfoWarning, "Credit Limit Required", "Credit Limit for Financing or Bank PO is required! Please check your masterfile first.", 1000, 2
                Me.Visible = False
                
            End If
            
'If parts or service'======================================================
        Else
            Set rsRequirements = gconDMIS.Execute("SELECT *,DATEADD(month,(terms_expiry_year * 12)+terms_expiry_month, getdate()) as expiry FROM ALL_CUSTOMER_TERMS WHERE customercode = '" & xCUSTNAME & "'")
            If Not (rsRequirements.EOF Or rsRequirements.BOF) Then
            '===========================================
            'credit limit
            Set RsCreditLimit = gconDMIS.Execute("SELECT *,DATEADD(month,(terms_expiry_year * 12)+terms_expiry_month, getdate()) as expiry FROM ALL_CUSTOMER_TERMS WHERE customercode = '" & xCUSTNAME & "'")
            If Not (RsCreditLimit.EOF Or RsCreditLimit.BOF) Then
        '       frmCMISCreditLimitInfo.PicCrInfo.Visible = True
                frmCMISCreditLimitInfo.txtCrExpiry.Text = RsCreditLimit!expiry
                frmCMISCreditLimitInfo.txtCrLimit.Text = ToDoubleNumber(RsCreditLimit!CREDITLIMIT)
                frmCMISCreditLimitInfo.txtCRTerm.Text = RsCreditLimit!CREDITTERM
            End If
            'credit available
            Set RsCreditLimit = gconDMIS.Execute("SELECT SUM(TTLINV)-(SELECT ISNULL(SUM(PAYMENT),0) FROM CMIS_OFF_DT WHERE TRANTYPE IN ('PI','AI','MI','SI') AND PAIDNA = 1 AND CUSCDE = '" & xCUSTNAME & "')  as balanse,(SELECT ISNULL(SUM(PAYMENT),0) FROM CMIS_OFF_DT WHERE TRANTYPE IN ('PI','AI','MI','SI') AND PAIDNA = 1 AND CUSCDE = '" & xCUSTNAME & "') payment FROM " & _
                    " ( " & _
                    "    SELECT ISNULL(SUM(CAST(TTLINVAMT AS DECIMAL(18,2))),0) TTLINV FROM PMIS_ORD_HD WHERE TYPE IN ('P','A','M') AND TRANTYPE IN ('CHG') AND STATUS IN ('P') and custcode= '" & xCUSTNAME & "' " & _
                    "        Union " & _
                    "    SELECT ISNULL(SUM(CAST(TTLINVAMT AS DECIMAL(18,2))),0) TTLINV FROM PMIS_ORD_HIST WHERE TYPE IN ('P','A','M') AND TRANTYPE IN ('CHG') AND STATUS IN ('P') and custcode= '" & xCUSTNAME & "'" & _
                    "        Union " & _
                    "    SELECT ISNULL(SUM(CAST(AMOUNT AS DECIMAL(18,2))),0) TTLINV FROM CSMS_REPOR WHERE ACCT_NO = '" & xCUSTNAME & "' AND TRANSTYPE = 'R' AND DTE_COMP IS NOT NULL " & _
                    " ) T")
            If RsCreditLimit!BALANSE <> 0 Then
                'balance
                frmCMISCreditLimitInfo.txtCRAvailable.Text = ToDoubleNumber(RsCreditLimit!BALANSE)
                xpayments = ToDoubleNumber(RsCreditLimit!PAYMENT)
            Else
                'balance
                frmCMISCreditLimitInfo.txtCRAvailable.Text = 0  'ToDoubleNumber(frmCMISCreditLimitInfo.txtCrLimit.Text)
            End If
            'past due
            Set RsCreditLimit = gconDMIS.Execute("select ar.customercode,ar.invoicetype,ar.invoiceno,ar.amount_topay,dt.invoiceamount from amis_ar ar inner join amis_detail dt on ar.InvoiceType = dt.InvoiceType " & _
            " And ar.INVOICENO = dt.INVOICENO And ar.CustomerCode = dt.CustomerCode And ar.account_code = dt.acct_code " & _
            " inner join all_customer_terms tr on ar.customercode=tr.customercode " & _
            " Where DateDiff(Day, ar.invoicedate, getdate()) > tr.creditterm " & _
            " and ar.amount_topay>dt.invoiceamount and ar.invoicetype in ('PI','AI','MI','SI') " & _
            " and ar.customercode = '" & xCUSTNAME & "' " & _
            " and ar.status='P' and dt.status='P' " & _
            " Union " & _
            " select ar.customercode,ar.invoicetype,invoiceno,amount_topay,0 as invoiceamount from amis_ar ar inner join all_customer_terms tr on ar.customercode=tr.customercode where ar.invoicetype = 'SI' and status = 'P' " & _
            " and invoiceno not in (select invoiceno from amis_detail where amis_detail.invoicetype = 'SI' and status = 'P'and customercode = '" & xCUSTNAME & "' ) and ar.customercode = '" & xCUSTNAME & "' and DateDiff(Day, ar.invoicedate, getdate()) > tr.creditterm")
           
            If RsCreditLimit!amount_topay + RsCreditLimit!invoiceamount <> 0 Then
                frmCMISCreditLimitInfo.Check1.Value = 1
                frmCMISCreditLimitInfo.cmdViewP.Visible = True
            End If
            'expiry
            Set RsCreditLimit = gconDMIS.Execute("select DATEADD(month,(terms_expiry_year * 12)+terms_expiry_month, getdate()) as expiry,getdate() present from all_customer_terms where customercode = '" & xCustCode & "'")
            If Not (RsCreditLimit.EOF Or RsCreditLimit.BOF) Then
                If RsCreditLimit!expiry < RsCreditLimit!present Then
                    MessagePop InfoWarning, "Credit Limit Expired", "Credit Limit for this customer has expired! Please check your masterfile first.", 1000, 2
                End If
            End If
        
            'Limit
'            Dim xxlimit As Double
'            Dim xxbalance As Double
            xxlimit = frmCMISCreditLimitInfo.txtCrLimit.Text
            xxbalance = ToDoubleNumber(frmCMISCreditLimitInfo.txtCRAvailable.Text)
            txtCreditAvail.Text = xxlimit - ToDoubleNumber(xxbalance)
            txtCreditAvail.Text = ToDoubleNumber(NumericVal(txtCreditAvail.Text))
            txtInvAmt.Text = ToDoubleNumber(XBALANCE)
            
            frmCMISCreditLimitInfo.picCrDetails.Visible = False
            frmCMISCreditLimitInfo.PicCrInfo.Enabled = False
            If frmCMISCreditLimitInfo.Check1.Value = 1 Then
        '        frmCMISCreditLimitInfo.PicCrInfo.Enabled = True
            End If
            If N2Str2Zero(frmCMISCreditLimitInfo.txtCRAvailable.Text) + ToDoubleNumber(XBALANCE) > ToDoubleNumber(frmCMISCreditLimitInfo.txtCreditAvail.Text) Then
                frmCMISCreditLimitInfo.cmdView.Visible = True
        '        frmCMISCreditLimitInfo.PicCrInfo.Enabled = True
            End If
        '    Unload Me
            Set RsCreditLimit = Nothing
            '======================================================
            Else
                frmCMISCreditLimitInfo.PicCrInfo.Enabled = False
                MessagePop InfoWarning, "Credit Limit Required", "Credit Limit for Charge Transaction is required! Please check your masterfile first.", 1000, 2
                Me.Visible = False
                
            End If
        End If
    End If
End Sub
Sub LOADJOURNAL(xname As String, xbal As Double, xMODULENAME As String)
    xCUSTNAME = xname
    XBALANCE = xbal
    XMod_name = xMODULENAME
'    frmCMISCreditLimitInfo.Show vbModal
    Call Form_Load
End Sub
Sub FillLstDet()
    Dim RS As New ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSpaidna As New ADODB.Recordset
'    Set RSpaidna = New ADODB.Recordset
    If XMod_name = "SMIS" Then ' If vehicle invoice
        If xView = "I" Then
            Dim Item As ListItem 'variable for your subitems
            lstDetails.ListItems.Clear 'Clear your Listview before loading a new data
            RS.Open "select * from smis_salesorder where STATUS <> 'C' and vi_no is not null and vi_no not in (select invoiceno from cmis_off_dt where trantype = 'VI') union select * from smis_salesorder where vi_no in (select invoiceno from cmis_off_dt where trantype = 'VI' and or_num in (select or_num from cmis_off_hd where Cancel = 1))and  vi_no not in (select invoiceno from cmis_off_dt where trantype = 'VI' and or_num in (select or_num from cmis_off_hd where Cancel <> 1))", gconDMIS
            If Not RS.EOF And Not RS.BOF Then
                While Not RS.EOF
                    Set Item = lstDetails.ListItems.Add(, , RS!VI_NO)
                    Item.SubItems(1) = Null2String(RS!InvoicedDate)
                    Item.SubItems(2) = ToDoubleNumber(RS!BALTOFINANCED)
                    Item.SubItems(3) = ToDoubleNumber(RS!BALTOFINANCED)
                    RS.MoveNext
                Wend
            End If
            Set RS = Nothing
             Set RSpaidna = gconDMIS.Execute("select * from (select vi_no,isnull(sum(baltofinanced),0) as invoiceamount,(select isnull(sum(payment),0) from cmis_off_dt where trantype = 'VI'  and paidna = 1 and cuscde = (select cuscde from all_customer where acctname ='" & xCUSTNAME & "')and invoiceno=vi_no) available From SMIS_SalesOrder WHERE 'VI'+'-'+vi_no in (select trantype+'-'+invoiceno from cmis_off_dt where trantype = 'VI'  and paidna = 1 and cuscde = (select cuscde from all_customer where acctname ='" & xCUSTNAME & "')) group by vi_no) T where invoiceamount-available <> 0")
             If Not RSpaidna.EOF And Not RSpaidna.BOF Then
             While Not RSpaidna.EOF
                RS.Open "select so.VI_NO,so.invoiceddate,so.baltofinanced,dt.payment from smis_salesorder so inner join cmis_off_dt dt on so.vi_no=dt.invoiceno  where  dt.trantype = 'VI' and vi_no = '" & RSpaidna!VI_NO & "'", gconDMIS
                If Not RS.EOF And Not RS.BOF Then
    '                While Not RS.EOF
                          Set Item = lstDetails.ListItems.Add(, , RS!VI_NO)
                            Item.SubItems(1) = Null2String(RS!InvoicedDate)
                            Item.SubItems(2) = ToDoubleNumber(RS!BALTOFINANCED)
                            Item.SubItems(3) = ToDoubleNumber(ToDoubleNumber(RS!BALTOFINANCED) - ToDoubleNumber(RSpaidna!Available))
    '                        RS.MoveNext
    '                    Wend
                End If
                Set RS = Nothing
                RSpaidna.MoveNext
             Wend
            End If
        Else
            lstDetails.ListItems.Clear 'Clear your Listview before loading a new data
        Set RS = gconDMIS.Execute("select ar.customercode,ar.invoicetype,ar.invoiceno,ar.amount_topay,dt.invoiceamount,ar.invoicedate from amis_ar ar inner join amis_detail dt on ar.InvoiceType = dt.InvoiceType " & _
        " And ar.INVOICENO = dt.INVOICENO And ar.CustomerCode = dt.CustomerCode And ar.account_code = dt.acct_code " & _
        " inner join all_customer_terms tr on ar.customercode=tr.customercode " & _
        " Where DateDiff(Day, ar.invoicedate, getdate()) > tr.creditterm " & _
        " and ar.amount_topay>dt.invoiceamount and ar.invoicetype = 'VI' " & _
        " and ar.customercode = (select top 1 cuscde from all_customer where acctname='" & xCUSTNAME & "') " & _
        " and ar.status='P' and dt.status='P' " & _
        " Union " & _
        " select customercode,invoicetype,invoiceno,amount_topay,0 as invoiceamount,invoicedate from amis_ar where invoicetype = 'VI' and status = 'P' " & _
        " and invoiceno not in (select invoiceno from amis_detail where invoicetype = 'VI' and status = 'P' and customercode = (select top 1 cuscde from all_customer where acctname='" & xCUSTNAME & "') ) and customercode = (select top 1 cuscde from all_customer where acctname='" & xCUSTNAME & "')")
    
            While Not RS.EOF
                Set Item = lstDetails.ListItems.Add(, , RS!INVOICENO)
                Item.SubItems(1) = Null2String(RS!invoicedate)
                Item.SubItems(2) = ToDoubleNumber(RS!amount_topay)
                Item.SubItems(3) = ToDoubleNumber(ToDoubleNumber(RS!amount_topay) - ToDoubleNumber(RS!invoiceamount))
                RS.MoveNext
            Wend
        End If
'If parts or service'======================================================
    Else
            If xView = "I" Then
            lstDetails.ListItems.Clear 'Clear your Listview before loading a new data
            Set RS = gconDMIS.Execute("SELECT " & _
                    " INV_DATE,INV_AMT,INV_NO,INV_AMT-(SELECT ISNULL(SUM(PAYMENT),0) FROM CMIS_OFF_DT WHERE TRANTYPE IN ('PI','AI','MI','SI') AND PAIDNA = 1 AND CUSCDE = '" & xCUSTNAME & "'and TRANTYPE+'-'+invoiceno = INV_NO)  AS BALANSE " & _
                    " From " & _
                    " ( " & _
                    "    SELECT [TYPE]+''+'I'+'-'+TRANNO AS INV_NO,TRANDATE INV_DATE,ISNULL(TTLINVAMT ,0) INV_AMT FROM PMIS_ORD_HIST WHERE TYPE IN ('P','A','M') AND TRANTYPE IN ('CHG') AND STATUS IN ('P') AND CUSTCODE= '" & xCUSTNAME & "' " & _
                    "        Union " & _
                    "    SELECT [TYPE]+''+'I'+'-'+TRANNO AS INV_NO,TRANDATE INV_DATE,ISNULL(TTLINVAMT ,0) INV_AMT FROM PMIS_ORD_HD WHERE TYPE IN ('P','A','M') AND TRANTYPE IN ('CHG') AND STATUS IN ('P') AND CUSTCODE= '" & xCUSTNAME & "' " & _
                    "        Union " & _
                    "    SELECT 'SI'+'-'+INVOICE AS INV_NO,DTE_COMP INV_DATE,ISNULL(AMOUNT ,0) INV_AMT FROM CSMS_REPOR WHERE ACCT_NO = '" & xCUSTNAME & "' AND TRANSTYPE = 'R' AND DTE_COMP IS NOT NULL " & _
                    " ) " & _
                    " T WHERE ISNULL(INV_AMT-(SELECT ISNULL(SUM(PAYMENT),0) FROM CMIS_OFF_DT WHERE TRANTYPE IN ('PI','AI','MI','SI') AND PAIDNA = 1 AND CUSCDE = '" & xCUSTNAME & "'),0)<>0 ")

            If Not RS.EOF And Not RS.BOF Then
                While Not RS.EOF
                    Set Item = lstDetails.ListItems.Add(, , RS!INV_NO)
                    Item.SubItems(1) = Null2String(RS!INV_DATE)
                    Item.SubItems(2) = ToDoubleNumber(RS!INV_AMT)
                    Item.SubItems(3) = ToDoubleNumber(RS!BALANSE)
                    RS.MoveNext
                Wend
            End If
            Set RS = Nothing
        Else
            lstDetails.ListItems.Clear 'Clear your Listview before loading a new data
            Set RS = gconDMIS.Execute("select ar.customercode,ar.invoicetype,ar.invoiceno,ar.amount_topay,dt.invoiceamount,ar.invoicedate from amis_ar ar inner join amis_detail dt on ar.InvoiceType = dt.InvoiceType " & _
                " And ar.INVOICENO = dt.INVOICENO And ar.CustomerCode = dt.CustomerCode And ar.account_code = dt.acct_code " & _
                " inner join all_customer_terms tr on ar.customercode=tr.customercode " & _
                " Where DateDiff(Day, ar.invoicedate, getdate()) > tr.creditterm " & _
                " and ar.amount_topay>dt.invoiceamount and ar.invoicetype in ('PI','AI','MI','SI') " & _
                " and ar.customercode = '" & xCUSTNAME & "' " & _
                " and ar.status='P' and dt.status='P' " & _
                " Union " & _
                " select ar.customercode,ar.invoicetype,invoiceno,amount_topay,0 as invoiceamount,ar.invoicedate from amis_ar ar inner join all_customer_terms tr on ar.customercode=tr.customercode where ar.invoicetype = 'SI' and status = 'P' " & _
                " and invoiceno not in (select invoiceno from amis_detail where amis_detail.invoicetype = 'SI' and status = 'P'and customercode = '" & xCUSTNAME & "' ) and ar.customercode = '" & xCUSTNAME & "' and DateDiff(Day, ar.invoicedate, getdate()) > tr.creditterm")
    
            While Not RS.EOF
                Set Item = lstDetails.ListItems.Add(, , RS!INVOICENO)
                Item.SubItems(1) = Null2String(RS!invoicedate)
                Item.SubItems(2) = ToDoubleNumber(RS!amount_topay)
                Item.SubItems(3) = ToDoubleNumber(ToDoubleNumber(RS!amount_topay) - ToDoubleNumber(RS!invoiceamount))
                RS.MoveNext
            Wend
        End If
    End If
End Sub

Sub override_mode(Xo As String)
    override = Xo
End Sub
Sub fillcbo()
    Dim rsusers As New ADODB.Recordset
    cboUsers.Clear
    rsusers.Open "select * from all_rams_users", gconDMIS
        While Not rsusers.EOF
            cboUsers.AddItem rsusers!user_name
        rsusers.MoveNext
        Wend
    Set rsusers = Nothing
End Sub

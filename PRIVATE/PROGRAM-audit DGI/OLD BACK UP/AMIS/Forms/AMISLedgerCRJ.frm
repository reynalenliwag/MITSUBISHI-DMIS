VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISLedgerCRJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CASH RECEIPTS VOUCHERS"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   ForeColor       =   &H00E0E0E0&
   Icon            =   "AMISLedgerCRJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   11520
   Begin VB.Frame fraDetails 
      Height          =   5985
      Left            =   90
      TabIndex        =   0
      Top             =   -30
      Width           =   11355
      Begin MSComctlLib.ListView lvLedgerCRJ 
         Height          =   5175
         Left            =   60
         TabIndex        =   4
         Top             =   150
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   9128
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "VOUCHER NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DOC DATE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "OR NUMBER"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "INVOICE NO."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "INVOICE DATE"
            Object.Width           =   2909
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "INVOICE TYPE"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "INVOICE AMT"
            Object.Width           =   3245
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   90
         ScaleHeight     =   525
         ScaleWidth      =   11175
         TabIndex        =   1
         Top             =   5340
         Width           =   11175
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   8430
            MaxLength       =   20
            TabIndex        =   2
            Top             =   60
            Width           =   2625
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
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
            Height          =   255
            Left            =   6930
            TabIndex        =   3
            Top             =   90
            Width           =   1395
         End
      End
   End
End
Attribute VB_Name = "frmAMISLedgerCRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xLedger                                       As ListItem
Dim rsCRJ_Ledger                                  As ADODB.Recordset
Dim OR_NUM                                        As String

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    FillCRJ
End Sub

Sub FillCRJ()
    Dim tmpTotal                                  As Double
    tmpTotal = 0
    Set rsCRJ_Ledger = New ADODB.Recordset
    rsCRJ_Ledger.Open "SELECT HD.InvoiceType as HD_INVOICETYPE, HD.InvoiceNo AS HD_INVOICENO, CRJ.INVOICEDATE AS CRJ_INVOICEDATE, CRJ.INVOICEAMOUNT AS CRJ_INV_AMOUNT, HD.CustomerCode AS HD_CUST_CODE, CRJ.VoucherNo as CRJ_VOUCHERNO FROM AMIS_Journal_HD HD INNER JOIN AMIS_CRJ_Detail CRJ ON HD.InvoiceNo = CRJ.INVOICENO AND HD.InvoiceType = CRJ.INVOICETYPE WHERE " & _
                      "(HD.InvoiceNo = '" & INVOICENO & "') AND (HD.InvoiceType = '" & InvoiceType & "') AND (HD.CustomerCode = '" & CUSCODE & "')", gconDMIS
    If Not rsCRJ_Ledger.EOF And Not rsCRJ_Ledger.BOF Then
        Do While Not rsCRJ_Ledger.EOF
            Set xLedger = lvLedgerCRJ.ListItems.Add(, , Null2String(rsCRJ_Ledger!CRJ_VOUCHERNO))
            xLedger.SubItems(1) = Null2String(rsCRJ_Ledger!CRJ_INVOICEDATE)
            xLedger.SubItems(2) = GET_ORNUM(Null2String(rsCRJ_Ledger!HD_INVOICENO), Null2String(rsCRJ_Ledger!HD_INVOICETYPE))
            xLedger.SubItems(3) = Null2String(rsCRJ_Ledger!HD_INVOICENO)
            xLedger.SubItems(4) = Null2String(rsCRJ_Ledger!CRJ_INVOICEDATE)
            xLedger.SubItems(5) = Null2String(rsCRJ_Ledger!HD_INVOICETYPE)
            xLedger.SubItems(6) = ToDoubleNumber(N2Str2Zero(rsCRJ_Ledger!CRJ_INV_AMOUNT))
            tmpTotal = tmpTotal + ToDoubleNumber(N2Str2Zero(rsCRJ_Ledger!CRJ_INV_AMOUNT))
            txtTotal = Format(tmpTotal, "#,###,##0.00")
            rsCRJ_Ledger.MoveNext
        Loop
    Else
        MessagePop RecNotFound, "No Record", "No Cash Receipts Vouchers to view"
    End If
End Sub

Function GET_ORNUM(xINVOICENO As String, xINVOICETYPE As String) As String
    Dim rsGET_ORNUM                               As ADODB.Recordset
    Set rsGET_ORNUM = gconDMIS.Execute("SELECT OR_NUM FROM CMIS_OFF_DT WHERE INVOICENO = '" & xINVOICENO & "' AND TRANTYPE = '" & xINVOICETYPE & "'")
    If Not rsGET_ORNUM.EOF And Not rsGET_ORNUM.BOF Then
        GET_ORNUM = Null2String(rsGET_ORNUM!OR_NUM)
    End If
    Set rsGET_ORNUM = Nothing
End Function

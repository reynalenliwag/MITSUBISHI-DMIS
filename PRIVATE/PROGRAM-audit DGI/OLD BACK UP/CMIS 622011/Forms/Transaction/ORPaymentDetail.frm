VERSION 5.00
Object = "{F985F9B0-A252-46B5-A444-E023A386B6FE}#1.0#0"; "WIZBOX.OCX"
Begin VB.Form frmORPaymentDetail 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Official Receipt Payment Details"
   ClientHeight    =   6000
   ClientLeft      =   135
   ClientTop       =   885
   ClientWidth     =   8835
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "ORPaymentDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdChangeToCash 
      Caption         =   "Change To Cash"
      Height          =   255
      Left            =   4020
      TabIndex        =   57
      ToolTipText     =   "Change to Cash"
      Top             =   1830
      Width           =   1815
   End
   Begin VB.TextBox txtCARDBNKCODE 
      Height          =   315
      Left            =   2310
      TabIndex        =   48
      Top             =   3000
      Width           =   4485
   End
   Begin VB.TextBox txtCARDNUMBER 
      Height          =   315
      Left            =   2310
      TabIndex        =   47
      Top             =   3330
      Width           =   2085
   End
   Begin VB.TextBox txtCardAmount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2310
      TabIndex        =   46
      Top             =   3660
      Width           =   1635
   End
   Begin VB.TextBox txtCardDate 
      Height          =   315
      Left            =   7050
      TabIndex        =   45
      Top             =   3330
      Width           =   1635
   End
   Begin VB.TextBox txtCheckDate 
      Height          =   315
      Left            =   7050
      TabIndex        =   42
      Top             =   1140
      Width           =   1635
   End
   Begin VB.TextBox txtSukli 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7080
      TabIndex        =   39
      Top             =   4530
      Width           =   1605
   End
   Begin VB.TextBox txtBayadAmt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7080
      TabIndex        =   34
      Top             =   4200
      Width           =   1605
   End
   Begin VB.TextBox txtVat 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2310
      TabIndex        =   33
      Top             =   5520
      Width           =   1635
   End
   Begin VB.TextBox txtConsumed 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2310
      TabIndex        =   30
      Top             =   5190
      Width           =   1635
   End
   Begin VB.TextBox txtTax 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2310
      TabIndex        =   29
      Top             =   4860
      Width           =   1635
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2310
      TabIndex        =   28
      Top             =   4530
      Width           =   1635
   End
   Begin VB.TextBox txtOR_Amt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2310
      TabIndex        =   19
      Top             =   4200
      Width           =   1635
   End
   Begin VB.TextBox txtCashAmount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2310
      TabIndex        =   18
      Top             =   2400
      Width           =   1635
   End
   Begin VB.TextBox txtChkAmount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2310
      TabIndex        =   15
      Top             =   1800
      Width           =   1635
   End
   Begin VB.TextBox txtTseklase 
      Height          =   315
      Left            =   2310
      TabIndex        =   14
      Top             =   1470
      Width           =   4485
   End
   Begin VB.TextBox txtTseke 
      Height          =   315
      Left            =   2310
      TabIndex        =   13
      Top             =   1140
      Width           =   2085
   End
   Begin VB.TextBox txtBankCode 
      Height          =   315
      Left            =   2310
      TabIndex        =   4
      Top             =   810
      Width           =   4485
   End
   Begin VB.TextBox txtOR_NUM 
      Height          =   315
      Left            =   7050
      TabIndex        =   3
      Top             =   210
      Width           =   1635
   End
   Begin VB.TextBox txtModeOfPayment 
      Height          =   315
      Left            =   2310
      TabIndex        =   1
      Top             =   210
      Width           =   2295
   End
   Begin wizBox.Box Box2 
      Height          =   1485
      Left            =   60
      Top             =   720
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   2619
   End
   Begin wizBox.Box Box3 
      Height          =   615
      Left            =   60
      Top             =   2250
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   1085
   End
   Begin wizBox.Box Box4 
      Height          =   1845
      Left            =   60
      Top             =   4080
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   3254
   End
   Begin wizBox.Box Box1 
      Height          =   615
      Left            =   60
      Top             =   60
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   1085
   End
   Begin wizBox.Box Box5 
      Height          =   1125
      Left            =   60
      Top             =   2910
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   1984
   End
   Begin VB.Label Label40 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Bank"
      Height          =   315
      Left            =   210
      TabIndex        =   56
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label39 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Card Number"
      Height          =   315
      Left            =   210
      TabIndex        =   55
      Top             =   3330
      Width           =   1815
   End
   Begin VB.Label Label37 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Card Amount"
      Height          =   315
      Left            =   210
      TabIndex        =   54
      Top             =   3660
      Width           =   1815
   End
   Begin VB.Label Label36 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   53
      Top             =   3000
      Width           =   195
   End
   Begin VB.Label Label35 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   52
      Top             =   3330
      Width           =   195
   End
   Begin VB.Label Label33 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   51
      Top             =   3660
      Width           =   195
   End
   Begin VB.Label Label32 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Card Date"
      Height          =   315
      Left            =   4950
      TabIndex        =   50
      Top             =   3330
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   6810
      TabIndex        =   49
      Top             =   3330
      Width           =   195
   End
   Begin VB.Label Label31 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   285
      Left            =   6810
      TabIndex        =   44
      Top             =   240
      Width           =   195
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   43
      Top             =   210
      Width           =   195
   End
   Begin VB.Label Label29 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   6810
      TabIndex        =   41
      Top             =   1140
      Width           =   195
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Check"
      Height          =   315
      Left            =   4950
      TabIndex        =   40
      Top             =   1140
      Width           =   1815
   End
   Begin VB.Label Label27 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   6840
      TabIndex        =   38
      Top             =   4530
      Width           =   195
   End
   Begin VB.Label Label26 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   6840
      TabIndex        =   37
      Top             =   4200
      Width           =   195
   End
   Begin VB.Label Label25 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due"
      Height          =   315
      Left            =   4980
      TabIndex        =   36
      Top             =   4530
      Width           =   1815
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Tendered"
      Height          =   315
      Left            =   4980
      TabIndex        =   35
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label23 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   32
      Top             =   5520
      Width           =   195
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Vat"
      Height          =   315
      Left            =   210
      TabIndex        =   31
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   27
      Top             =   5190
      Width           =   195
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   26
      Top             =   4860
      Width           =   195
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   25
      Top             =   4530
      Width           =   195
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   24
      Top             =   4200
      Width           =   195
   End
   Begin VB.Label Label17 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Consumed"
      Height          =   315
      Left            =   210
      TabIndex        =   23
      Top             =   5190
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Tax"
      Height          =   315
      Left            =   210
      TabIndex        =   22
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Discount"
      Height          =   315
      Left            =   210
      TabIndex        =   21
      Top             =   4530
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payment"
      Height          =   315
      Left            =   210
      TabIndex        =   20
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   17
      Top             =   2400
      Width           =   195
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Amount"
      Height          =   315
      Left            =   210
      TabIndex        =   16
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   12
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   11
      Top             =   1470
      Width           =   195
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   10
      Top             =   1140
      Width           =   195
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   315
      Left            =   2070
      TabIndex        =   9
      Top             =   810
      Width           =   195
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Check Amount"
      Height          =   315
      Left            =   210
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Check Class"
      Height          =   315
      Left            =   210
      TabIndex        =   7
      Top             =   1470
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
      Height          =   315
      Left            =   210
      TabIndex        =   6
      Top             =   1140
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Bank"
      Height          =   315
      Left            =   210
      TabIndex        =   5
      Top             =   810
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Number"
      Height          =   285
      Left            =   4950
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mode of Payment"
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1815
   End
End
Attribute VB_Name = "frmORPaymentDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Vat_OR_DATE          As String

Private Sub cmdChangeToCash_Click()
If Function_Access(LOGID, "Acess_Edit") = False Then Exit Sub
    
    gconDMIS.Execute ("update CMIS_Off_Hd Set" & _
                    " Tseke = NULL," & _
                    " CheckDate = NULL," & _
                    " Tseklase = NULL," & _
                    " ChkAmount = 0," & _
                    " bankCode = NULL," & _
                    " bankBranch = NULL," & _
                    " TOF = 1," & _
                    " CashAmount = " & NumericVal(txtChkAmount.Text) & _
                    " where OR_NUM = '" & OR_NUMBER_GLOBAL & "'")
    Dim rsCashPos        As ADODB.Recordset
    Set rsCashPos = New ADODB.Recordset
    Set rsCashPos = gconDMIS.Execute("Select * from CMIS_Cash_Pos Where CUTDATE >='" & Vat_OR_DATE & "'")
    If Not rsCashPos.EOF And Not rsCashPos.BOF Then
        rsCashPos.MoveFirst
        Do While Not rsCashPos.EOF
            gconDMIS.Execute ("update CMIS_Cash_Pos set" & _
                            " CASH = CASH + " & NumericVal(txtChkAmount.Text) & "," & _
                            " [CHECK] = [CHECK] - " & NumericVal(txtChkAmount.Text) & _
                            " where ID = " & rsCashPos!Id)
            rsCashPos.MoveNext
        Loop
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then Unload Me
'    If KeyCode = vbKeyF1 Then cmdChangeToCash_Click
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    StoreMemVars
End Sub

Sub StoreMemVars()
    Dim rsOFF_HD         As ADODB.Recordset
    Set rsOFF_HD = New ADODB.Recordset
    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd Where OR_NUM = '" & OR_NUMBER_GLOBAL & "'")
    If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
        Vat_OR_DATE = Null2Date(rsOFF_HD!OR_DATE)
        If N2Str2Zero(rsOFF_HD!ChkAmount) > 0 Then
            txtModeOfPayment.Text = "Check Payment"
        ElseIf N2Str2Zero(rsOFF_HD!CashAmount) > 0 Then
            txtModeOfPayment.Text = "Cash Payment"
        ElseIf N2Str2Zero(rsOFF_HD!cardamount) > 0 Then
            txtModeOfPayment.Text = "Card Payment"
        Else
            txtModeOfPayment.Text = ""
        End If
        txtOR_NUM.Text = Null2String(rsOFF_HD!OR_NUM)
        txtTseke.Text = Null2String(rsOFF_HD!Tseke)
        txtCheckDate.Text = Null2String(rsOFF_HD!CheckDate)
        txtTseklase.Text = SetCheckClass(Null2String(rsOFF_HD!Tseklase))
        txtChkAmount.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!ChkAmount))
        txtCashAmount.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!CashAmount))
        If txtModeOfPayment.Text = "Card Payment" Then
            txtBankCode.Text = ""
            txtCARDBNKCODE.Text = SetBankName(Null2String(rsOFF_HD!cardbnkcde)) & " - " & Null2String(rsOFF_HD!Bankbranch)
        Else
            txtBankCode.Text = SetBankName(Null2String(rsOFF_HD!bankcode)) & " - " & Null2String(rsOFF_HD!Bankbranch)
            txtCARDBNKCODE.Text = ""
        End If
        txtCARDNUMBER.Text = Null2String(rsOFF_HD!cardnumber)
        txtCardDate.Text = Null2String(rsOFF_HD!carddate)
        txtCardAmount.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!cardamount))

        txtOR_Amt.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!OR_AMT))
        txtDiscount.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!discount))
        txtTax.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!tax))
        txtConsumed.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!Consumed))
        txtVat.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!Vat))
        txtBayadAmt.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!BayadAmt))
        txtSukli.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!Sukli))
    End If
End Sub

Function SetBankName(xxx As Variant)
    Dim rsSBOOK          As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select DESCNAME from CMIS_SBOOK Where Book = 'B' and CODE = '" & xxx & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetBankName = rsSBOOK!DESCNAME
    End If
End Function

'Function SetCheckClass(xxx As Variant)
'    Dim rsSBOOK          As ADODB.Recordset
'    Set rsSBOOK = New ADODB.Recordset
'    Set rsSBOOK = gconDMIS.Execute("Select DESCNAME from CMIS_SBOOK Where Book = 'F' and CODE = '" & xxx & "'")
'    Else
'        txtModeOfPayment.Text = "": txtOR_NUM.Text = "": txtBankCode.Text = ""
'        txtTseke.Text = "": txtCheckDate.Text = "": txtTseklase.Text = ""
'        txtChkAmount.Text = "0.00": txtCashAmount.Text = "0.00": txtOR_Amt.Text = "0.00"
'        txtDiscount.Text = "0.00": txtTax.Text = "0.00": txtConsumed.Text = "0.00"
'        txtVat.Text = "0.00": txtBayadAmt.Text = "0.00": txtSukli.Text = "0.00"
'        txtCARDBNKCODE.Text = "": txtCARDNUMBER.Text = "": txtCardDate.Text = "": txtCardAmount.Text = "0.00"
'    End If
'End Function

'Function SetBankName(xxx As Variant)
'    Dim rsSBOOK          As ADODB.Recordset
'    Set rsSBOOK = New ADODB.Recordset
'    Set rsSBOOK = gconDMIS.Execute("Select DESCNAME from CMIS_SBOOK Where Book = 'B' and CODE = '" & xxx & "'")
'    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
'        SetBankName = rsSBOOK!DESCNAME
'    End If
'End Function
'
Function SetCheckClass(xxx As Variant)
    Dim rsSBOOK          As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select DESCNAME from CMIS_SBOOK Where Book = 'F' and CODE = '" & xxx & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClass = rsSBOOK!DESCNAME
    End If
End Function

VERSION 5.00
Begin VB.Form frmCMISCARDPaymentEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Card Payment Entry Box"
   ClientHeight    =   2445
   ClientLeft      =   240
   ClientTop       =   750
   ClientWidth     =   4755
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CARDPaymentEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4755
   Begin VB.ComboBox txtBankBranch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   60
      Width           =   2415
   End
   Begin VB.ComboBox cboCRDBNKCDE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   2250
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1920
      Width           =   2415
   End
   Begin VB.PictureBox picCreditCard 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   60
      ScaleHeight     =   1305
      ScaleWidth      =   4695
      TabIndex        =   5
      Top             =   480
      Width           =   4695
      Begin VB.TextBox txtCARDDATE2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   16
         Top             =   420
         Width           =   810
      End
      Begin VB.TextBox txtCARDNUMBER 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   2160
         TabIndex        =   0
         Top             =   0
         Width           =   2385
      End
      Begin VB.TextBox txtCARDDATE 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   1
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtCARDAMOUNT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   2160
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   840
         Width           =   2400
      End
      Begin VB.Label Label 
         Caption         =   "mm/yyyy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3870
         TabIndex        =   17
         Top             =   570
         Width           =   675
      End
      Begin VB.Line Line 
         BorderWidth     =   2
         X1              =   2910
         X2              =   2820
         Y1              =   420
         Y2              =   750
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Card Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   12
         Top             =   60
         Width           =   1905
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1950
         TabIndex        =   11
         Top             =   30
         Width           =   165
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   10
         Top             =   450
         Width           =   1905
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1950
         TabIndex        =   9
         Top             =   450
         Width           =   165
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Card Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   8
         Top             =   870
         Width           =   1905
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1950
         TabIndex        =   7
         Top             =   870
         Width           =   165
      End
      Begin VB.Label labOR_NUM 
         Caption         =   "Label13"
         Height          =   255
         Left            =   2850
         TabIndex        =   6
         Top             =   870
         Width           =   945
      End
   End
   Begin VB.PictureBox picBankCode 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   30
      ScaleHeight     =   405
      ScaleWidth      =   4755
      TabIndex        =   13
      Top             =   60
      Width           =   4755
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1980
         TabIndex        =   15
         Top             =   0
         Width           =   165
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   14
         Top             =   90
         Width           =   1905
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2010
      TabIndex        =   4
      Top             =   1980
      Width           =   165
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Terminal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   1980
      Width           =   1905
   End
End
Attribute VB_Name = "frmCMISCARDPaymentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSBOOK                                                 As ADODB.Recordset

Function SetBankCode(XXX As Variant)
    Dim rsSBOOK                                             As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select Code from CMIS_SBOOK Where Book = 'B' and DescName = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetBankCode = rsSBOOK!Code
    End If
End Function

Function SetBankCode2(XXX As Variant)
    Dim rsSBOOK                                             As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select CusCde from CMIS_CardBank Where AcctName = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetBankCode2 = rsSBOOK!CUSCDE
    End If
End Function

Function SetBankCode3(XXX As Variant)
    Dim rsSBOOK                                             As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select CusCde from CMIS_CardCompany Where AcctName = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetBankCode3 = rsSBOOK!CUSCDE
    End If
End Function

Function SetCheckClassCode(XXX As Variant)
    Dim rsSBOOK                                             As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select Code from CMIS_SBOOK Where Book = 'F' and DescName = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClassCode = rsSBOOK!Code
    End If
End Function

Sub FillCbo()
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select ACCTNAME from CMIS_CardBank order by ACCTNAME ASC")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        Combo_Loadval cboCRDBNKCDE, rsSBOOK
    End If
    
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select ACCTNAME from CMIS_CardCOMPANY order by ACCTNAME ASC")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        Combo_Loadval txtBankBranch, rsSBOOK
    End If
    txtBankBranch.ListIndex = 0
    cboCRDBNKCDE.ListIndex = 0
End Sub

Sub Save_CARD_Payment()
    If COMPANY_CODE = M_COMPANY_CODE Then
        gconDMIS.Execute ("update CMIS_Off_Hd Set" & _
                          " CARDBNKCDE = " & N2Str2Null(SetBankCode2(cboCRDBNKCDE.Text)) & "," & _
                          " BANKCODE = " & N2Str2Null(SetBankCode2(txtBankBranch.Text)) & "," & _
                          " BANKBRANCH = " & N2Str2Null(txtBankBranch.Text) & "," & _
                          " CARDNUMBER = " & N2Str2Null(txtCARDNUMBER.Text) & "," & _
                          " CARDDATE = " & N2Str2Null(txtCardDate.Text) + "'1'" + N2Str2Null(txtCARDDATE2.Text) & _
                          " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL))
        AMOUNT_TENDERED = NumericVal(txtCardAmount.Text)
    Else
        gconDMIS.Execute ("update CMIS_Off_Hd Set" & _
                          " CARDBNKCDE = " & N2Str2Null(SetBankCode2(cboCRDBNKCDE.Text)) & "," & _
                          " BANKCODE = " & N2Str2Null(SetBankCode3(txtBankBranch.Text)) & "," & _
                          " BANKBRANCH = " & N2Str2Null(txtBankBranch.Text) & "," & _
                          " CARDNUMBER = " & N2Str2Null(txtCARDNUMBER.Text) & "," & _
                          " CARDDATE = " & N2Str2Null(txtCardDate.Text) + "'1'" + N2Str2Null(txtCARDDATE2.Text) & "," & _
                          " CARDAMOUNT = " & NumericVal(txtCardAmount.Text) & _
                          " where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL))
        RECEIPTS_AMOUNT = NumericVal(txtCardAmount.Text)
        Unload Me
        frmCMISCASHPaymentEntry.Show vbModal
    End If
    LogAudit "A", "CARD PAYMENT", txtCARDNUMBER
End Sub

Private Sub txtBankBranch_GotFocus()
    If VALID_COMPANY_CODE_FORHAI = True Or COMPANY_CODE = "HCI" Then
    Else
        VBComBoBoxDroppedDown txtBankBranch
    End If
End Sub

Private Sub txtBankBranch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCARDNUMBER.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
'    Case Else
'        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtCardDate.Text = Month(LOGDATE)
    txtCARDDATE2.Text = Year(LOGDATE)
    If COMPANY_CODE = M_COMPANY_CODE Then
        txtCardAmount.Text = NumericVal(RECEIPTS_AMOUNT) - CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR)
    Else
        txtCardAmount.Text = NumericVal(RECEIPTS_AMOUNT)
    End If
    
    
    'FillCbo
    
    Dim rsCreditCardBank As ADODB.Recordset
    Set rsCreditCardBank = New ADODB.Recordset
    rsCreditCardBank.Open "SELECT * FROM CMIS_CARDCOMPANY", gconDMIS, adOpenForwardOnly
    If Not rsCreditCardBank.EOF And Not rsCreditCardBank.BOF Then
        txtBankBranch.TabIndex = 1
        picCreditCard.Top = 480
        picBankCode.Visible = True
        FillCbo
    Else
        picBankCode.Visible = False
        picCreditCard.Top = 330
        txtCARDNUMBER.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    txtCardAmount.Text = ""
    txtCardAmount.Text = "0.00"
End Sub

Private Sub txtCARDAMOUNT_GotFocus()
    If NumericVal(txtCardAmount.Text) = 0 Then txtCardAmount.Text = "" Else txtCardAmount.Text = NumericVal(txtCardAmount.Text)
End Sub


Private Sub txtCARDAMOUNT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If VALID_COMPANY_CODE_FORHAI = True Then
            If txtCARDNUMBER = "" Then
                MsgBox "Please enter Credit Card No.", vbInformation, "Message"
                Exit Sub
            End If
        Else
            If cboCRDBNKCDE.Visible = True Then
                If cboCRDBNKCDE = "" Then
                    MsgBox "Please select Credit Card Company", vbInformation, "Card Company"
                    Exit Sub
                End If
            End If
            
            If txtCARDNUMBER = "" Then
                MsgBox "Please enter Credit Card No.", vbInformation, "Message"
                Exit Sub
            End If
        End If
        
        If COMPANY_CODE = M_COMPANY_CODE Then
            If Round(NumericVal(txtCardAmount.Text), 2) > (Round(NumericVal(RECEIPTS_AMOUNT), 2) - Round(NumericVal(CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR)), 2)) Then
                MsgBox "Payment Amount exceeds OR Amount...", vbInformation, "Message"
                Exit Sub
            End If
            If MsgBox("Card Description Correct?", vbQuestion + vbYesNo) = vbYes Then
                Save_CARD_Payment
                Unload Me
                frmCMISCASHPaymentEntry.Show vbModal
                If RECEIPTS_BALANCE > 0 Then
                    frmCMISOREntry.RefreshDisplay
                    frmCMISOREntry.picPayment.ZOrder 0
                    frmCMISOREntry.picPayment.Visible = True
                    frmCMISOREntry.optCASH.Enabled = True
                    frmCMISOREntry.optCASH.SetFocus
                    frmCMISOREntry.optCHECK.Enabled = True
                    frmCMISOREntry.optCARD.Enabled = False
                    frmCMISOREntry.optCARD.Value = False
                    frmCMISOREntry.optCANCEL.Value = False
                End If
            End If
        Else
            If MsgBox("Card Description Correct?", vbQuestion + vbYesNo) = vbYes Then
                If NumericVal(txtCardAmount.Text) <= 0 Then
                    MsgBox "Enter correct amount.", vbInformation, "Amount Received"
                ElseIf NumericVal(txtCardAmount.Text) < RECEIPTS_AMOUNT Then
                    If MsgBox("Amount entered is less than OR amount." & Chr(13) & "Accept payment?", vbQuestion + vbYesNo, "Card Amount") = vbYes Then
                        Save_CARD_Payment
                    End If
                ElseIf NumericVal(txtCardAmount.Text) > RECEIPTS_AMOUNT Then
                    If MsgBox("Amount entered is greater than OR amount." & Chr(13) & "Accept over payment?", vbQuestion + vbYesNo, "Card Amount") = vbYes Then
                        Save_CARD_Payment
                    End If
                Else
                    Save_CARD_Payment
                End If
            End If
        End If
    End If
End Sub

Private Sub txtCARDAMOUNT_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCARDAMOUNT_LostFocus()
    txtCardAmount.Text = ToDoubleNumber(txtCardAmount.Text)
End Sub

Private Sub txtCARDDATE_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCARDDATE_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
    If KeyAscii = 13 Then
        If txtCardDate.Text = "" Then
            MsgBox "Please enter month.", vbInformation, "Month"
            txtCardDate.SetFocus
        Else
            txtCARDDATE2.SetFocus
        End If
    End If
End Sub

Private Sub txtCARDDATE_LostFocus()
    If txtCardDate.Text = "" Then
        MsgBox "Please enter month.", vbInformation, "Month"
        txtCardDate.SetFocus
    ElseIf NumericVal(txtCardDate.Text) <= 0 Or NumericVal(txtCardDate) > 12 Then
        MsgBox "Invalid month range.", vbInformation, "Month"
        txtCardDate.SetFocus
    End If
End Sub

Private Sub txtCARDDATE2_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCARDDATE2_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
    If KeyAscii = 13 Then
        If txtCARDDATE2.Text = "" Then
            MsgBox "Please enter year.", vbInformation, "Month"
            txtCARDDATE2.SetFocus
        Else
            txtCardAmount.SetFocus
        End If
    End If
End Sub

Private Sub txtCARDDATE2_LostFocus()
    If txtCARDDATE2.Text = "" Then
        MsgBox "Please enter year.", vbInformation, "Year"
        txtCARDDATE2.SetFocus
    ElseIf NumericVal(txtCARDDATE2.Text) < Year(LOGDATE) Then
        MsgBox "Please check the expiration date.", vbInformation, "Year"
        txtCARDDATE2.SetFocus
    End If
End Sub

Private Sub txtCARDNUMBER_KeyPress(KeyAscii As Integer)
'KeyAscii = OnlyNumeric(KeyAscii)
    If KeyAscii = 13 Then
        If txtCARDNUMBER.Text = "" Then
            MsgBox "Please enter Credit Card No.", vbInformation, "Message"
            txtCARDNUMBER.SetFocus
            Exit Sub
        Else
            txtCardDate.SetFocus
        End If
    End If
End Sub


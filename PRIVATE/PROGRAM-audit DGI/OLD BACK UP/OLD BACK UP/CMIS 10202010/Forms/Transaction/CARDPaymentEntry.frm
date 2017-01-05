VERSION 5.00
Begin VB.Form frmCMISCARDPaymentEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Card Payment Entry Box"
   ClientHeight    =   1890
   ClientLeft      =   240
   ClientTop       =   750
   ClientWidth     =   4755
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CARDPaymentEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
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
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   2415
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
      Left            =   2220
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   1380
      Width           =   2370
   End
   Begin VB.TextBox txtCARDDATE 
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
      Left            =   2220
      TabIndex        =   3
      Top             =   960
      Width           =   2370
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
      Left            =   2220
      TabIndex        =   2
      Top             =   540
      Width           =   2385
   End
   Begin VB.TextBox txtBankBranch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2220
      TabIndex        =   0
      Top             =   2820
      Width           =   3945
   End
   Begin VB.Label labOR_NUM 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2910
      TabIndex        =   15
      Top             =   1410
      Width           =   945
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
      Left            =   2010
      TabIndex        =   14
      Top             =   1410
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
      Left            =   90
      TabIndex        =   13
      Top             =   1410
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
      Left            =   2010
      TabIndex        =   12
      Top             =   990
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
      Left            =   90
      TabIndex        =   11
      Top             =   990
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
      Left            =   2010
      TabIndex        =   10
      Top             =   570
      Width           =   165
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
      Left            =   90
      TabIndex        =   9
      Top             =   570
      Width           =   1905
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
      TabIndex        =   8
      Top             =   2850
      Width           =   165
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Branch"
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
      TabIndex        =   7
      Top             =   2850
      Width           =   1905
   End
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
      Left            =   2010
      TabIndex        =   6
      Top             =   150
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
      Left            =   90
      TabIndex        =   5
      Top             =   150
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
        SetBankCode = rsSBOOK!code
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

Function SetCheckClassCode(XXX As Variant)
    Dim rsSBOOK                                             As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select Code from CMIS_SBOOK Where Book = 'F' and DescName = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClassCode = rsSBOOK!code
    End If
End Function

Sub FillCbo()
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select ACCTNAME from CMIS_CardBank order by ACCTNAME ASC")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        Combo_Loadval cboCRDBNKCDE, rsSBOOK
    End If
End Sub

Sub Save_CARD_Payment()
    If COMPANY_CODE = M_COMPANY_CODE Then
        gconDMIS.Execute ("update CMIS_Off_Hd Set" & _
                          " CARDBNKCDE = " & N2Str2Null(SetBankCode2(cboCRDBNKCDE.Text)) & "," & _
                          " BANKBRANCH = " & N2Str2Null(txtBankBranch.Text) & "," & _
                          " CARDNUMBER = " & N2Str2Null(txtCARDNUMBER.Text) & "," & _
                          " CARDDATE = " & N2Str2Null(txtCardDate.Text) & _
                          " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL))
        AMOUNT_TENDERED = NumericVal(txtCardAmount.Text)
    Else
        gconDMIS.Execute ("update CMIS_Off_Hd Set" & _
                          " CARDBNKCDE = " & N2Str2Null(SetBankCode(cboCRDBNKCDE.Text)) & "," & _
                          " BANKBRANCH = " & N2Str2Null(txtBankBranch.Text) & "," & _
                          " CARDNUMBER = " & N2Str2Null(txtCARDNUMBER.Text) & "," & _
                          " CARDDATE = " & N2Str2Null(txtCardDate.Text) & "," & _
                          " CARDAMOUNT = " & NumericVal(txtCardAmount.Text) & _
                          " where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL))
        RECEIPTS_AMOUNT = NumericVal(txtCardAmount.Text)
        Unload Me
        frmCMISCASHPaymentEntry.Show vbModal
    End If
    LogAudit "A", "CARD PAYMENT", txtCARDNUMBER
End Sub

Private Sub cboCRDBNKCDE_GotFocus()
    VBComBoBoxDroppedDown cboCRDBNKCDE
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtCardDate.Text = LOGDATE
    If COMPANY_CODE = M_COMPANY_CODE Then
        txtCardAmount.Text = NumericVal(RECEIPTS_AMOUNT) - CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR)
    End If
    FillCbo
End Sub

Private Sub txtBankBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCARDAMOUNT_GotFocus()
    If NumericVal(txtCardAmount.Text) = 0 Then txtCardAmount.Text = "" Else txtCardAmount.Text = NumericVal(txtCardAmount.Text)
End Sub


Private Sub txtCARDAMOUNT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboCRDBNKCDE = "" Then
            MsgBox "Please select Credit Card Company", vbInformation, "Card Company"
            Exit Sub
        ElseIf txtCARDNUMBER = "" Then
            MsgBox "Please enter Credit Card No.", vbInformation, "Message"
            Exit Sub
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
                If NumericVal(txtCardAmount.Text) < RECEIPTS_AMOUNT Then
                    MsgBox "Amount entered is less than OR amount.", vbInformation, "Check Amount"
                    Exit Sub
                ElseIf NumericVal(txtCardAmount.Text) > RECEIPTS_AMOUNT Then
                    MsgBox "Amount entered is greater than OR amount.", vbInformation, "Check Amount"
                    Exit Sub
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

Private Sub txtCARDNUMBER_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub


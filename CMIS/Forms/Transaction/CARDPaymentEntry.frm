VERSION 5.00
Begin VB.Form frmCMISCARDPaymentEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Card Payment Entry Box"
   ClientHeight    =   2055
   ClientLeft      =   240
   ClientTop       =   750
   ClientWidth     =   6510
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CARDPaymentEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6510
   Begin VB.ComboBox txtBankBranch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   0
      Width           =   4305
   End
   Begin VB.ComboBox cboCRDBNKCDE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1680
      Width           =   4305
   End
   Begin VB.PictureBox picCreditCard 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   6375
      TabIndex        =   5
      Top             =   480
      Width           =   6375
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
         Height          =   375
         Left            =   2160
         MaxLength       =   16
         TabIndex        =   0
         Top             =   0
         Width           =   2865
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
         Width           =   2865
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
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   0
         Width           =   255
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
         Width           =   1785
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
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   255
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
      Left            =   30
      TabIndex        =   3
      Top             =   1740
      Width           =   1905
   End
End
Attribute VB_Name = "frmCMISCARDPaymentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSBOOK                                                         As ADODB.Recordset

Function SetBankCode(XXX As Variant)
    Dim rsSBOOK                                                     As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT Code FROM CMIS_SBOOK WHERE Book = 'B' AND DescName = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetBankCode = rsSBOOK!Code
    End If
End Function

Function SetBankCode2(XXX As Variant)
    Dim rsSBOOK                                                     As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT CusCde FROM CMIS_CardBank WHERE AcctName = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetBankCode2 = rsSBOOK!CUSCDE
    End If
End Function

Function SetBankCode3(XXX As Variant)
    Dim rsSBOOK                                                     As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT CusCde FROM CMIS_CardCompany WHERE AcctName = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetBankCode3 = rsSBOOK!CUSCDE
    End If
End Function

Function SetCheckClassCode(XXX As Variant)
    Dim rsSBOOK                                                     As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT Code FROM CMIS_SBOOK WHERE Book = 'F' AND DescName = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClassCode = rsSBOOK!Code
    End If
End Function

Sub FillCbo()
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT ACCTNAME FROM CMIS_CardBank ORDER BY ACCTNAME ASC")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        Combo_Loadval cboCRDBNKCDE, rsSBOOK
    End If
    
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT ACCTNAME FROM CMIS_CardCOMPANY ORDER BY ACCTNAME ASC")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        Combo_Loadval txtBankBranch, rsSBOOK
    End If
    txtBankBranch.ListIndex = 0
    cboCRDBNKCDE.ListIndex = 0
End Sub

Sub Save_CARD_Payment()

    Dim vTOTALDISC                                                   As String
    Dim vTOTALTAX                                                    As String
    
    Dim rsDiscEWT                                                   As New ADODB.Recordset
    Set rsDiscEWT = New ADODB.Recordset
    Set rsDiscEWT = gconDMIS.Execute("Select SUM(DISCOUNT) AS TOTALDISC,SUM(TAX) AS TOTALTAX from CMIS_Off_dt where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL) & "")
    If Not rsDiscEWT.EOF And Not rsDiscEWT.BOF Then
         vTOTALDISC = (rsDiscEWT!TOTALDISC)
         vTOTALTAX = (rsDiscEWT!TOTALTAX)
    End If
    
    If COMPANY_CODE = M_COMPANY_CODE Then
        gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET" & _
                          " CARDBNKCDE = " & N2Str2Null(SetBankCode2(cboCRDBNKCDE.Text)) & "," & _
                          " BANKBRANCH = " & N2Str2Null(txtBankBranch.Text) & "," & _
                          " CARDNUMBER = " & N2Str2Null(txtCARDNUMBER.Text) & "," & _
                          " CARDDATE = " & N2Str2Null(txtCARDDATE.Text) + "'1'" + N2Str2Null(txtCARDDATE2.Text) & _
                          " WHERE VAT = " & VAT_OR & " and OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL))
        AMOUNT_TENDERED = NumericVal(txtCARDAMOUNT.Text)
    Else
        gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET" & _
                          " CARDBNKCDE = " & N2Str2Null(SetBankCode2(cboCRDBNKCDE.Text)) & "," & _
                          " BANKBRANCH = " & N2Str2Null(txtBankBranch.Text) & "," & _
                          " CARDNUMBER = " & N2Str2Null(txtCARDNUMBER.Text) & "," & _
                          " CARDDATE = " & N2Str2Null(txtCARDDATE.Text) + "'1'" + N2Str2Null(txtCARDDATE2.Text) & "," & _
                          " CARDAMOUNT = " & NumericVal(txtCARDAMOUNT.Text) & "," & _
                          " DISCOUNT = " & NumericVal(vTOTALDISC) & "," & _
                          " TAX = " & NumericVal(vTOTALTAX) & "" & _
                          " WHERE OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL))
        RECEIPTS_AMOUNT = NumericVal(txtCARDAMOUNT.Text)
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
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "]" '"." & App.Revision & "]"
    txtCARDDATE.Text = Month(LOGDATE)
    txtCARDDATE2.Text = Year(LOGDATE)
    txtCARDNUMBER.MaxLength = 16
    
    If COMPANY_CODE = M_COMPANY_CODE Then
        txtCARDAMOUNT.Text = NumericVal(RECEIPTS_AMOUNT) - CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR)
    Else
        txtCARDAMOUNT.Text = NumericVal(RECEIPTS_AMOUNT)
    End If
    'FillCbo
    
'    If COMPANY_CODE = "DJM" Then
'        txtCARDAMOUNT.Locked = True
'    Else
        txtCARDAMOUNT.Locked = False
'    End If
    
    Dim rsCreditCardBank                                            As ADODB.Recordset
    Set rsCreditCardBank = New ADODB.Recordset
    rsCreditCardBank.Open "Select * from CMIS_CARDCOMPANY", gconDMIS, adOpenForwardOnly
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
    txtCARDAMOUNT.Text = ""
    txtCARDAMOUNT.Text = "0.00"
End Sub

Private Sub txtCARDAMOUNT_GotFocus()
    If NumericVal(txtCARDAMOUNT.Text) = 0 Then txtCARDAMOUNT.Text = "" Else txtCARDAMOUNT.Text = NumericVal(txtCARDAMOUNT.Text)
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
            If Round(NumericVal(txtCARDAMOUNT.Text), 2) > (Round(NumericVal(RECEIPTS_AMOUNT), 2) - Round(NumericVal(CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR)), 2)) Then
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
                If NumericVal(txtCARDAMOUNT.Text) <= 0 Then
                    MsgBox "Enter correct amount.", vbInformation, "Amount Received"
                ElseIf NumericVal(txtCARDAMOUNT.Text) < RECEIPTS_AMOUNT Then
                    If MsgBox("Amount entered is less than OR Amount." & Chr(13) & "Accept payment?", vbQuestion + vbYesNo, "Card Amount") = vbYes Then
                        Save_CARD_Payment
                    End If
                ElseIf NumericVal(txtCARDAMOUNT.Text) > RECEIPTS_AMOUNT Then
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
    txtCARDAMOUNT.Text = ToDoubleNumber(txtCARDAMOUNT.Text)
End Sub

Private Sub txtCARDDATE_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCARDDATE_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
    If KeyAscii = 13 Then
        If txtCARDDATE.Text = "" Then
            MsgBox "Please enter Month.", vbInformation, "Month"
            txtCARDDATE.SetFocus
        Else
            txtCARDDATE2.SetFocus
        End If
    End If
End Sub

Private Sub txtCARDDATE_LostFocus()
    If txtCARDDATE.Text = "" Then
        MsgBox "Please enter Month.", vbInformation, "Month"
        txtCARDDATE.SetFocus
    ElseIf NumericVal(txtCARDDATE.Text) <= 0 Or NumericVal(txtCARDDATE) > 12 Then
        MsgBox "Invalid Month range.", vbInformation, "Month"
        txtCARDDATE.SetFocus
    End If
End Sub

Private Sub txtCARDDATE2_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCARDDATE2_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
    If KeyAscii = 13 Then
        If txtCARDDATE2.Text = "" Then
            MsgBox "Please enter Year.", vbInformation, "Year"
            txtCARDDATE2.SetFocus
        Else
            txtCARDAMOUNT.SetFocus
        End If
    End If
End Sub

Private Sub txtCARDDATE2_LostFocus()
    If txtCARDDATE2.Text = "" Then
        MsgBox "Please enter Year.", vbInformation, "Year"
        txtCARDDATE2.SetFocus
    ElseIf NumericVal(txtCARDDATE2.Text) < Year(LOGDATE) Then
        MsgBox "Please check the expiration Date.", vbInformation, "Date"
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
            txtCARDDATE.SetFocus
        End If
        
        'JJE 3/4/2013
        If COMPANY_CODE = "DSSC" Then
            If Len(txtCARDNUMBER) < 16 Then
                MsgBox "Cardnumber must be 16 characters", vbOKOnly, "INVALID"
            End If
        End If
        'JJE
    End If
End Sub


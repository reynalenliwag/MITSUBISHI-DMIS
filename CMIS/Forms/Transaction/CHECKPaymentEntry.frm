VERSION 5.00
Begin VB.Form frmCMISCHECKPaymentEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Payment Entry Box"
   ClientHeight    =   3060
   ClientLeft      =   300
   ClientTop       =   960
   ClientWidth     =   6255
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CHECKPaymentEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6255
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3120
      Width           =   1275
   End
   Begin VB.ComboBox cboTseklase 
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
      ForeColor       =   &H00973640&
      Height          =   360
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1740
      Width           =   2715
   End
   Begin VB.ComboBox cboBankCode 
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
      ForeColor       =   &H00973640&
      Height          =   360
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   3915
   End
   Begin VB.TextBox txtChkAmount 
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
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   2160
      Width           =   1785
   End
   Begin VB.TextBox txtCheckDate 
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
      Top             =   1320
      Width           =   1785
   End
   Begin VB.TextBox txtTseke 
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
      MaxLength       =   10
      TabIndex        =   3
      Top             =   900
      Width           =   3945
   End
   Begin VB.TextBox txtBankBranch 
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
      MaxLength       =   200
      TabIndex        =   2
      Top             =   480
      Width           =   3945
   End
   Begin VB.Label Note 
      Caption         =   "NOTE: Multiple Checknumber in one OR is not allowed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   5295
   End
   Begin VB.Label labOR_NUM 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2910
      TabIndex        =   19
      Top             =   2190
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
      TabIndex        =   18
      Top             =   2190
      Width           =   165
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Amount"
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
      TabIndex        =   17
      Top             =   2190
      Width           =   1905
   End
   Begin VB.Label Label10 
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
      TabIndex        =   16
      Top             =   1770
      Width           =   165
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Class"
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
      TabIndex        =   15
      Top             =   1770
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
      TabIndex        =   14
      Top             =   1350
      Width           =   165
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Date"
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
      Top             =   1350
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
      TabIndex        =   12
      Top             =   930
      Width           =   165
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
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
      Top             =   930
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
      TabIndex        =   10
      Top             =   510
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
      TabIndex        =   9
      Top             =   510
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
      TabIndex        =   8
      Top             =   90
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
      TabIndex        =   7
      Top             =   90
      Width           =   1905
   End
End
Attribute VB_Name = "frmCMISCHECKPaymentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function SetBankCode(XXX As Variant)
    Dim rsSBOOK                                                     As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT Code FROM CMIS_SBOOK WHERE Book = 'B' AND DescName = " & N2Str2Null(XXX))
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetBankCode = rsSBOOK!Code
    End If
    Set rsSBOOK = Nothing
End Function

Function SetCheckClassCode(XXX As Variant)
    Dim rsSBOOK                                                     As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT Code FROM CMIS_SBOOK WHERE Book = 'F' AND DescName = " & N2Str2Null(XXX))
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClassCode = rsSBOOK!Code
    End If
    Set rsSBOOK = Nothing
End Function

Sub FillCbo()
    Dim rsSBOOK                                                     As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT DESCNAME FROM CMIS_SBOOK WHERE BOOK = 'F' ORDER BY DESCNAME ASC")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        Combo_Loadval cboTseklase, rsSBOOK
    End If
    
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT DESCNAME FROM CMIS_SBOOK WHERE BOOK = 'B' ORDER BY DESCNAME ASC")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        Combo_Loadval cboBankCode, rsSBOOK
    End If
    Set rsSBOOK = Nothing
End Sub

Sub Save_CHECK_Payment()
    
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
                          " BANKCODE = " & N2Str2Null(SetBankCode(cboBankCode.Text)) & "," & _
                          " BANKBRANCH = " & N2Str2Null(txtBankBranch.Text) & "," & _
                          " TSEKE = " & N2Str2Null(txtTseke.Text) & "," & _
                          " CHECKDATE = " & N2Str2Null(txtCheckDate.Text) & "," & _
                          " TSEKLASE = " & N2Str2Null(SetCheckClassCode(cboTseklase.Text)) & _
                          " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL))
        'RECEIPTS_AMOUNT
        AMOUNT_TENDERED = NumericVal(txtChkAmount.Text)
    Else
        gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET" & _
                          " BANKCODE = " & N2Str2Null(SetBankCode(cboBankCode.Text)) & "," & _
                          " BANKBRANCH = " & N2Str2Null(txtBankBranch.Text) & "," & _
                          " TSEKE = " & N2Str2Null(txtTseke.Text) & "," & _
                          " CHECKDATE = " & N2Str2Null(txtCheckDate.Text) & "," & _
                          " TSEKLASE = " & N2Str2Null(SetCheckClassCode(cboTseklase.Text)) & "," & _
                          " CHKAMOUNT = " & NumericVal(txtChkAmount.Text) & "," & _
                          " DISCOUNT = " & NumericVal(vTOTALDISC) & "," & _
                          " TAX = " & NumericVal(vTOTALTAX) & "" & _
                          " WHERE VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL))
        RECEIPTS_AMOUNT = NumericVal(txtChkAmount.Text)
        Unload Me
        frmCMISCASHPaymentEntry.Show vbModal
    End If
    LogAudit "A", "BANK PAYMENT", cboBankCode
End Sub

Private Sub cboBankCode_GotFocus()
    VBComBoBoxDroppedDown cboBankCode
End Sub

Private Sub cboTseklase_GotFocus()
    VBComBoBoxDroppedDown cboTseklase
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
    txtCheckDate.Text = LOGDATE
    If COMPANY_CODE = M_COMPANY_CODE Then
        txtChkAmount.Text = NumericVal(RECEIPTS_AMOUNT) - CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR)
    Else
        txtChkAmount.Text = NumericVal(RECEIPTS_AMOUNT)
    End If
    
    FillCbo
    txtTseke.MaxLength = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    txtChkAmount.Text = ""
End Sub



Private Sub txtBankBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCheckDate_GotFocus()
    If IsDate(txtCheckDate.Text) = True Then
        txtCheckDate.Text = Format(txtCheckDate.Text, "MM/DD/YYYY")
    Else
        txtCheckDate.Text = ""
    End If
End Sub

Private Sub txtCheckDate_LostFocus()
    If IsDate(txtCheckDate.Text) = True Then
        txtCheckDate.Text = Format(txtCheckDate.Text, "DD-MMM-YY")
    Else
        txtCheckDate.Text = ""
    End If
End Sub

Private Sub txtChkAmount_GotFocus()
    If NumericVal(txtChkAmount.Text) = 0 Then txtChkAmount.Text = "" Else txtChkAmount.Text = NumericVal(txtChkAmount.Text)
End Sub

Private Sub txtChkAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboBankCode.Text = "" Then
            MsgBox "Please select bank.", vbInformation, "Message"
            cboBankCode.SetFocus
            Exit Sub
        ElseIf txtTseke.Text = "" Then
            MsgBox "Please enter check number.", vbInformation, "Message"
            txtTseke.SetFocus
            Exit Sub
        ElseIf txtCheckDate.Text = "" Then
            MsgBox "Please enter check date.", vbInformation, "Message"
            txtCheckDate.SetFocus
            Exit Sub
        ElseIf cboTseklase.Text = "" Then
            MsgBox "Please select check class.", vbInformation, "Message"
            cboTseklase.SetFocus
            Exit Sub
        ElseIf txtChkAmount.Text = "" Then
            MsgBox "Please enter check amount.", vbInformation, "Message"
            txtChkAmount.SetFocus
            Exit Sub
            '        ElseIf Round(NumericVal(txtChkAmount.Text), 2) > (Round(NumericVal(RECEIPTS_AMOUNT), 2) - Round(NumericVal(CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR)), 2)) Then
            '            MsgBox "Payment Amount exceeds OR Amount...", vbInformation, "Message"
            '            Exit Sub
        End If

        If MsgBox("Check Description Correct?", vbQuestion + vbYesNo) = vbYes Then
            If COMPANY_CODE = M_COMPANY_CODE Then
                Save_CHECK_Payment
                Unload Me
                frmCMISCASHPaymentEntry.Show vbModal
                If RECEIPTS_BALANCE > 0 Then
                    frmCMISOREntry.RefreshDisplay
                    frmCMISOREntry.picPayment.ZOrder 0
                    frmCMISOREntry.picPayment.Visible = True
                    frmCMISOREntry.optCASH.Enabled = True
                    frmCMISOREntry.optCASH.SetFocus
                    frmCMISOREntry.optCHECK.Enabled = False
                    frmCMISOREntry.optCARD.Value = False
                    frmCMISOREntry.optCANCEL.Value = False
                End If
            Else
                If NumericVal(txtChkAmount.Text) < RECEIPTS_AMOUNT Then
                    If MsgBox("Amount entered is less than invoice amount." & Chr(13) & "Accept payment?", vbQuestion + vbYesNo, "Check Amount") = vbYes Then
                        Save_CHECK_Payment
                    End If
                ElseIf NumericVal(txtChkAmount.Text) > RECEIPTS_AMOUNT Then
                    If MsgBox("Amount entered is greater than invoice amount." & Chr(13) & "Accept over payment?", vbQuestion + vbYesNo, "Check Amount") = vbYes Then
                        Save_CHECK_Payment
                    End If
                Else
                    Save_CHECK_Payment
                End If
            End If
        End If
    End If
End Sub

Private Sub txtChkAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtChkAmount_LostFocus()
    txtChkAmount.Text = ToDoubleNumber(txtChkAmount.Text)
End Sub

Private Sub txtTseke_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub


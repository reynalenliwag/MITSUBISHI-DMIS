VERSION 5.00
Object = "{F985F9B0-A252-46B5-A444-E023A386B6FE}#1.0#0"; "WIZBOX.OCX"
Object = "{BBFAB1E6-C4DE-4F06-A9A7-1FBDDBDF668B}#1.0#0"; "SMALL3DDIGITS.OCX"
Begin VB.Form frmCMISCARDCHECKPaymentEntry 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check & Card Payment Entry Box"
   ClientHeight    =   5910
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6375
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "CARDCHECKPaymentEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture6 
      Height          =   735
      Left            =   1500
      ScaleHeight     =   675
      ScaleWidth      =   4665
      TabIndex        =   35
      Top             =   60
      Width           =   4725
      Begin wizSmall3DDigits.wizSmall3DDigit wizSmall3DDigit1 
         Height          =   735
         Left            =   -1320
         TabIndex        =   36
         Top             =   -30
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1296
      End
   End
   Begin wizBox.Box Box1 
      Height          =   2655
      Left            =   60
      Top             =   870
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4683
   End
   Begin VB.TextBox Text1 
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
      Left            =   2280
      TabIndex        =   23
      Top             =   4110
      Width           =   3945
   End
   Begin VB.TextBox txtCARDNUMBER 
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
      Left            =   2280
      TabIndex        =   22
      Top             =   4530
      Width           =   2385
   End
   Begin VB.TextBox txtCARDDATE 
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
      Left            =   2280
      TabIndex        =   21
      Top             =   4950
      Width           =   1785
   End
   Begin VB.TextBox txtCARDAMOUNT 
      Alignment       =   1  'Right Justify
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
      Left            =   2280
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   5370
      Width           =   1785
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3690
      Width           =   2715
   End
   Begin VB.ComboBox cboTseklase 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2670
      Width           =   2715
   End
   Begin VB.ComboBox cboBankCode 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   990
      Width           =   2715
   End
   Begin VB.TextBox txtChkAmount 
      Alignment       =   1  'Right Justify
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
      Left            =   2280
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   3090
      Width           =   1785
   End
   Begin VB.TextBox txtCheckDate 
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
      Left            =   2280
      TabIndex        =   3
      Top             =   2250
      Width           =   1785
   End
   Begin VB.TextBox txtTseke 
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
      Left            =   2280
      TabIndex        =   2
      Top             =   1830
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
      Left            =   2280
      TabIndex        =   1
      Top             =   1410
      Width           =   3945
   End
   Begin wizBox.Box Box2 
      Height          =   2235
      Left            =   60
      Top             =   3600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3942
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   37
      Top             =   90
      Width           =   1305
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Code"
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
      Left            =   150
      TabIndex        =   34
      Top             =   3720
      Width           =   1905
   End
   Begin VB.Label Label22 
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
      Left            =   2070
      TabIndex        =   33
      Top             =   3720
      Width           =   165
   End
   Begin VB.Label Label21 
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
      Left            =   150
      TabIndex        =   32
      Top             =   4140
      Width           =   1905
   End
   Begin VB.Label Label20 
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
      Left            =   2070
      TabIndex        =   31
      Top             =   4140
      Width           =   165
   End
   Begin VB.Label Label19 
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
      Left            =   150
      TabIndex        =   30
      Top             =   4560
      Width           =   1905
   End
   Begin VB.Label Label18 
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
      Left            =   2070
      TabIndex        =   29
      Top             =   4560
      Width           =   165
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Date"
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
      Left            =   150
      TabIndex        =   28
      Top             =   4980
      Width           =   1905
   End
   Begin VB.Label Label16 
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
      Left            =   2070
      TabIndex        =   27
      Top             =   4980
      Width           =   165
   End
   Begin VB.Label Label15 
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
      Left            =   150
      TabIndex        =   26
      Top             =   5400
      Width           =   1905
   End
   Begin VB.Label Label14 
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
      Left            =   2070
      TabIndex        =   25
      Top             =   5400
      Width           =   165
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2880
      TabIndex        =   24
      Top             =   2640
      Width           =   945
   End
   Begin VB.Label labOR_NUM 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2970
      TabIndex        =   18
      Top             =   3120
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
      Left            =   2070
      TabIndex        =   17
      Top             =   3120
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
      Left            =   150
      TabIndex        =   16
      Top             =   3120
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
      Left            =   2070
      TabIndex        =   15
      Top             =   2700
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
      Left            =   150
      TabIndex        =   14
      Top             =   2700
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
      Left            =   2070
      TabIndex        =   13
      Top             =   2280
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
      Left            =   150
      TabIndex        =   12
      Top             =   2280
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
      Left            =   2070
      TabIndex        =   11
      Top             =   1860
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
      Left            =   150
      TabIndex        =   10
      Top             =   1860
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
      Left            =   2070
      TabIndex        =   9
      Top             =   1440
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
      Left            =   150
      TabIndex        =   8
      Top             =   1440
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
      Left            =   2070
      TabIndex        =   7
      Top             =   1020
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
      Left            =   150
      TabIndex        =   6
      Top             =   1020
      Width           =   1905
   End
End
Attribute VB_Name = "frmCMISCARDCHECKPaymentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
txtCheckDate.Text = LOGDATE
FillCbo
End Sub

Sub FillCbo()
Dim rsSBOOK As ADODB.Recordset
Set rsSBOOK = New ADODB.Recordset
Set rsSBOOK = gconCMIS.Execute("Select DESCNAME from SBOOK Where BOOK = 'F' order by DESCNAME asc")
If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
   Combo_Loadval cboTseklase, rsSBOOK
End If
Set rsSBOOK = New ADODB.Recordset
Set rsSBOOK = gconCMIS.Execute("Select DESCNAME from SBOOK Where BOOK = 'B' order by DESCNAME asc")
If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
   Combo_Loadval cboBankCode, rsSBOOK
End If
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
If KeyCode = vbKeyReturn Then Save_CHECK_Payment
End Sub

Sub Save_CHECK_Payment()
gconCMIS.Execute ("Update OFF_HD Set" & _
                  " BANKCODE = " & N2Str2Null(SetBankCode(cboBankCode.Text)) & "," & _
                  " BANKBRANCH = " & N2Str2Null(txtBankBranch.Text) & "," & _
                  " TSEKE = " & N2Str2Null(txtTseke.Text) & "," & _
                  " CHECKDATE = " & N2Str2Null(txtCheckDate.Text) & "," & _
                  " TSEKLASE = " & N2Str2Null(SetCheckClassCode(cboTseklase.Text)) & "," & _
                  " CHKAMOUNT = " & NumericVal(txtChkAmount.Text) & "," & _
                  " PAIDNA = 1" & _
                  " where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL))
RECEIPTS_AMOUNT = NumericVal(txtChkAmount.Text)
Unload Me
frmCMISCASHPaymentEntry.Show vbModal
End Sub

Function SetBankCode(XXX As Variant)
Dim rsSBOOK As ADODB.Recordset
Set rsSBOOK = New ADODB.Recordset
Set rsSBOOK = gconCMIS.Execute("Select Code from SBOOK Where Book = 'B' and DescName = " & N2Str2Null(XXX))
If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
   SetBankCode = rsSBOOK!CODE
End If
End Function

Function SetCheckClassCode(XXX As Variant)
Dim rsSBOOK As ADODB.Recordset
Set rsSBOOK = New ADODB.Recordset
Set rsSBOOK = gconCMIS.Execute("Select Code from SBOOK Where Book = 'F' and DescName = " & N2Str2Null(XXX))
If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
   SetCheckClassCode = rsSBOOK!CODE
End If
End Function

Private Sub txtChkAmount_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtChkAmount_LostFocus()
txtChkAmount.Text = ToDoubleNumber(txtChkAmount.Text)
End Sub

Private Sub txtTseke_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumeric(KeyAscii)
End Sub

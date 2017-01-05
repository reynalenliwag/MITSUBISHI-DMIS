VERSION 5.00
Object = "{F985F9B0-A252-46B5-A444-E023A386B6FE}#1.0#0"; "WIZBOX.OCX"
Begin VB.Form frmCMISCASHCHECKPaymentEntry 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Payment Entry Box"
   ClientHeight    =   3435
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6720
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "CASHCHECKPaymentEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   3480
      Width           =   885
   End
   Begin VB.TextBox Text1 
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
      Top             =   2880
      Width           =   1785
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
      TabIndex        =   5
      Top             =   1830
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
      TabIndex        =   1
      Top             =   150
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
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   2250
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
      TabIndex        =   4
      Top             =   1410
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
      TabIndex        =   3
      Top             =   990
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
      TabIndex        =   2
      Top             =   570
      Width           =   3945
   End
   Begin wizBox.Box Box1 
      Height          =   615
      Left            =   60
      Top             =   2760
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   1085
   End
   Begin wizBox.Box Box2 
      Height          =   2625
      Left            =   60
      Top             =   60
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   4630
   End
   Begin VB.Label Label15 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2970
      TabIndex        =   23
      Top             =   2910
      Width           =   945
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
      TabIndex        =   22
      Top             =   2910
      Width           =   165
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Amount"
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
      TabIndex        =   21
      Top             =   2910
      Width           =   1905
   End
   Begin VB.Label labOR_NUM 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2970
      TabIndex        =   19
      Top             =   2280
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
      TabIndex        =   18
      Top             =   2280
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
      TabIndex        =   17
      Top             =   2280
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
      TabIndex        =   16
      Top             =   1860
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
      TabIndex        =   15
      Top             =   1860
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
      TabIndex        =   14
      Top             =   1440
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
      TabIndex        =   13
      Top             =   1440
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
      TabIndex        =   12
      Top             =   1020
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
      TabIndex        =   11
      Top             =   1020
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
      TabIndex        =   10
      Top             =   600
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
      TabIndex        =   9
      Top             =   600
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
      TabIndex        =   8
      Top             =   180
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
      TabIndex        =   7
      Top             =   180
      Width           =   1905
   End
End
Attribute VB_Name = "frmCMISCASHCHECKPaymentEntry"
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

VERSION 5.00
Object = "{F985F9B0-A252-46B5-A444-E023A386B6FE}#1.0#0"; "WIZBOX.OCX"
Begin VB.Form frmCMISCASHCARDPaymentEntry 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Card Payment Entry Box"
   ClientHeight    =   3000
   ClientLeft      =   105
   ClientTop       =   645
   ClientWidth     =   6705
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "CASHCARDPaymentEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   30
      TabIndex        =   0
      Top             =   3060
      Width           =   495
   End
   Begin wizBox.Box Box2 
      Height          =   2265
      Left            =   60
      Top             =   30
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   3995
   End
   Begin wizBox.Box Box1 
      Height          =   615
      Left            =   60
      Top             =   2340
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   1085
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
      TabIndex        =   17
      Text            =   "0.00"
      Top             =   2460
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
      TabIndex        =   1
      Top             =   150
      Width           =   2715
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
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   1830
      Width           =   1785
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
      TabIndex        =   4
      Top             =   1410
      Width           =   1785
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
      TabIndex        =   20
      Top             =   2490
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
      TabIndex        =   19
      Top             =   2490
      Width           =   165
   End
   Begin VB.Label Label9 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2970
      TabIndex        =   18
      Top             =   2490
      Width           =   945
   End
   Begin VB.Label labOR_NUM 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2970
      TabIndex        =   16
      Top             =   1860
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
      TabIndex        =   15
      Top             =   1860
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
      Left            =   150
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   1440
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
      Left            =   150
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   1020
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
      Left            =   150
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   180
      Width           =   165
   End
   Begin VB.Label Label1 
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
      TabIndex        =   6
      Top             =   180
      Width           =   1905
   End
End
Attribute VB_Name = "frmCMISCASHCARDPaymentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSBOOK As ADODB.Recordset

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
FillCbo
End Sub

Sub FillCbo()
Set rsSBOOK = New ADODB.Recordset
Set rsSBOOK = gconCMIS.Execute("Select DESCNAME from SBOOK Where BOOK = 'B' order by DESCNAME asc")
If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
   Combo_Loadval cboCRDBNKCDE, rsSBOOK
End If
End Sub

Sub Save_CARD_Payment()
gconCMIS.Execute ("Update OFF_HD Set" & _
                  " CARDBNKCDE = " & N2Str2Null(SetBankCode(cboCRDBNKCDE.Text)) & "," & _
                  " BANKBRANCH = " & N2Str2Null(txtBankBranch.Text) & "," & _
                  " CARDNUMBER = " & N2Str2Null(txtCARDNUMBER.Text) & "," & _
                  " CARDDATE = " & N2Str2Null(txtCARDDATE.Text) & "," & _
                  " CARDAMOUNT = " & NumericVal(txtCARDAMOUNT.Text) & "," & _
                  " PAIDNA = 1" & _
                  " where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL))
RECEIPTS_AMOUNT = NumericVal(txtCARDAMOUNT.Text)
Unload Me
frmCMISCASHPaymentEntry.Show vbModal
End Sub

Function SetBankCode(XXX As Variant)
Dim rsSBOOK As ADODB.Recordset
Set rsSBOOK = New ADODB.Recordset
Set rsSBOOK = gconCMIS.Execute("Select Code from SBOOK Where Book = 'B' and DescName = '" & XXX & "'")
If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
   SetBankCode = rsSBOOK!CODE
End If
End Function

Function SetCheckClassCode(XXX As Variant)
Dim rsSBOOK As ADODB.Recordset
Set rsSBOOK = New ADODB.Recordset
Set rsSBOOK = gconCMIS.Execute("Select Code from SBOOK Where Book = 'F' and DescName = '" & XXX & "'")
If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
   SetCheckClassCode = rsSBOOK!CODE
End If
End Function

Private Sub txtBankBranch_KeyPress(KeyAscii As Integer)
KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCARDAMOUNT_GotFocus()
If NumericVal(txtCARDAMOUNT.Text) = 0 Then txtCARDAMOUNT.Text = "" Else txtCARDAMOUNT.Text = NumericVal(txtCARDAMOUNT.Text)
End Sub

Private Sub txtCARDAMOUNT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   If MsgBox("Card Description Correct?", vbQuestion + vbYesNo) = vbYes Then
      Save_CARD_Payment
   End If
End If
End Sub

Private Sub txtCARDAMOUNT_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCARDAMOUNT_LostFocus()
txtCARDAMOUNT.Text = ToDoubleNumber(txtCARDAMOUNT.Text)
End Sub

'Private Sub txtCARDDATE_GotFocus()
'If IsDate(txtCARDDATE.Text) = True Then txtCARDDATE.Text = Format(txtCARDDATE.Text, "MM/DD/YYYY") Else txtCARDDATE.Text = ""
'End Sub

'Private Sub txtCARDDATE_LostFocus()
'If IsDate(txtCARDDATE.Text) = True Then txtCARDDATE.Text = Format(txtCARDDATE.Text, "DD-MMM-YY") Else txtCARDDATE.Text = ""
'End Sub

Private Sub txtCARDNUMBER_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumeric(KeyAscii)
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCASHPOSITIONCardInBank 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Deposited Card"
   ClientHeight    =   5130
   ClientLeft      =   180
   ClientTop       =   540
   ClientWidth     =   7740
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CardInBank.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   5025
      Left            =   60
      ScaleHeight     =   5025
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin MSFlexGridLib.MSFlexGrid grdBANKDEPO 
         Height          =   3345
         Left            =   60
         TabIndex        =   5
         Top             =   15
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5900
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         MousePointer    =   99
         FormatString    =   "  Date              |              Bank Name             |   Time     |  Amount Deposit  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CardInBank.frx":030A
      End
      Begin VB.TextBox txtTotalCardDepo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   14
         Top             =   3510
         Width           =   1635
      End
      Begin VB.TextBox txtDeposit 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   3510
         Width           =   1635
      End
      Begin VB.TextBox txtCARDNum 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   3870
         Width           =   1635
      End
      Begin VB.TextBox txtCARDDATE 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   4230
         Width           =   1635
      End
      Begin VB.TextBox txtDeposit_To 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   4590
         Width           =   5355
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Card Deposit :"
         Height          =   315
         Left            =   3870
         TabIndex        =   15
         Top             =   3540
         Width           =   1815
      End
      Begin VB.Label Label61 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   3540
         Width           =   195
      End
      Begin VB.Label Label62 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Card Deposit"
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   3540
         Width           =   1815
      End
      Begin VB.Label Label65 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   3900
         Width           =   195
      End
      Begin VB.Label Label66 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Card Number"
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   3900
         Width           =   1815
      End
      Begin VB.Label Label67 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   4260
         Width           =   195
      End
      Begin VB.Label Label68 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date"
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   4260
         Width           =   1815
      End
      Begin VB.Label Label69 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   4620
         Width           =   195
      End
      Begin VB.Label Label70 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit to Bank"
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   4620
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmCASHPOSITIONCardInBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function SetBankName(XXX As Variant)
    Dim rsBankName                                                    As ADODB.Recordset
    Set rsBankName = New ADODB.Recordset
    Set rsBankName = gconDMIS.Execute("Select DescName from CMIS_SBOOK Where Book = 'B' and Code = '" & XXX & "'")
    If Not rsBankName.EOF And Not rsBankName.BOF Then
        SetBankName = rsBankName!DESCNAME
    End If
    Set rsBankName = Nothing
End Function

Sub InitGrid()
    cleargrid grdBankDepo
    grdBankDepo.FormatString = "Date              |               Bank Name             |   Time     |  Amount Deposit  "
    grdBankDepo.ColWidth(4) = 1
End Sub

Sub StoreMemvars()
    Dim TOTAL_CARD_IN_BANK                                            As Double
    Dim PlsLoveMe                                                     As Integer
    Dim rsBANKDEPO                                                    As ADODB.Recordset
    Set rsBANKDEPO = New ADODB.Recordset
    Set rsBANKDEPO = gconDMIS.Execute("Select * from CMIS_BankDepo WHERE [TYPE] = '3' AND CUTDATE = '" & CASHPOSITION_CUTOFF_DATE & "' ORDER BY ID ASC")
    'Set rsBANKDEPO = gconDMIS.Execute("Select * from CMIS_BankDepo WHERE [TYPE] = '3' AND DATDEPOSIT = '" & CASHPOSITION_CUTOFF_DATE & "' ORDER BY ID ASC")
    If Not rsBANKDEPO.EOF And Not rsBANKDEPO.BOF Then
        rsBANKDEPO.MoveFirst: PlsLoveMe = 0: TOTAL_CARD_IN_BANK = 0
        Do While Not rsBANKDEPO.EOF
            PlsLoveMe = PlsLoveMe + 1
            grdBankDepo.AddItem Null2String(rsBANKDEPO!bankcode) & Chr(9) & SetBankName(Null2String(rsBANKDEPO!bankcode)) & Chr(9) & Null2String(rsBANKDEPO!timdeposit) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPO!Deposit)) & Chr(9) & rsBANKDEPO!Id
            TOTAL_CARD_IN_BANK = TOTAL_CARD_IN_BANK + N2Str2Zero(rsBANKDEPO!Deposit)
            If PlsLoveMe = 1 Then grdBankDepo.RemoveItem 1
            rsBANKDEPO.MoveNext
        Loop
    End If
    txtTotalCardDepo.Text = ToDoubleNumber(TOTAL_CARD_IN_BANK)
    Set rsBANKDEPO = Nothing
End Sub

Sub StoreBANKDEPODetails(XXX As Variant)
    Dim rsBANKDEPO2                                                   As ADODB.Recordset
    Set rsBANKDEPO2 = New ADODB.Recordset
    Set rsBANKDEPO2 = gconDMIS.Execute("Select * from CMIS_BankDepo Where ID = " & XXX)
    If Not rsBANKDEPO2.EOF And Not rsBANKDEPO2.BOF Then
        txtDeposit.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO2!Deposit))
        txtCARDNum.Text = Null2String(rsBANKDEPO2!cardnumber)
        txtCardDate.Text = Null2String(rsBANKDEPO2!carddate)
        txtDeposit_To.Text = Null2String(rsBANKDEPO2!Deposit_To)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitGrid
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub grdBANKDEPO_Click()
    grdBankDepo.Col = 4
    If grdBankDepo.Text <> "" Then
        StoreBANKDEPODetails grdBankDepo.Text
    End If
End Sub

Private Sub grdBANKDEPO_GotFocus()
    grdBankDepo.Col = 4
    If grdBankDepo.Text <> "" Then
        StoreBANKDEPODetails grdBankDepo.Text
    End If
End Sub


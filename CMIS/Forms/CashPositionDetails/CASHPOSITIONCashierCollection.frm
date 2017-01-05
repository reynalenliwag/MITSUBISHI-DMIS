VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCASHPOSITIONCashierCollection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cashier Collection"
   ClientHeight    =   6045
   ClientLeft      =   180
   ClientTop       =   540
   ClientWidth     =   8505
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CASHPOSITIONCashierCollection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1305
      Left            =   60
      ScaleHeight     =   1305
      ScaleWidth      =   8385
      TabIndex        =   5
      Top             =   4680
      Width           =   8385
      Begin VB.TextBox txtTotalCollection 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   6690
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   60
         Width           =   1605
      End
      Begin VB.TextBox txtOR_NUM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   6690
         TabIndex        =   4
         Top             =   870
         Width           =   1605
      End
      Begin VB.TextBox txtChkNumber 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1530
         TabIndex        =   2
         Top             =   870
         Width           =   1815
      End
      Begin VB.TextBox txtChkDate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1530
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtChkAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   6690
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6570
         TabIndex        =   16
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Collection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5070
         TabIndex        =   15
         Top             =   90
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6570
         TabIndex        =   13
         Top             =   900
         Width           =   195
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "O.R. No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5070
         TabIndex        =   12
         Top             =   900
         Width           =   1485
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6570
         TabIndex        =   11
         Top             =   510
         Width           =   195
      End
      Begin VB.Label Label50 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         TabIndex        =   10
         Top             =   900
         Width           =   195
      End
      Begin VB.Label Label49 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         TabIndex        =   9
         Top             =   495
         Width           =   195
      End
      Begin VB.Label Label47 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   900
         Width           =   1365
      End
      Begin VB.Label Label46 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label45 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5070
         TabIndex        =   6
         Top             =   510
         Width           =   1485
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdPettyPay 
      Height          =   4575
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorSel    =   -2147483633
      BackColorBkg    =   -2147483633
      Appearance      =   0
      MousePointer    =   99
      FormatString    =   " Code           |   Bank Name                                      |    Time            | Check Amount   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "CASHPOSITIONCashierCollection.frx":030A
   End
End
Attribute VB_Name = "frmCASHPOSITIONCashierCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOFF_HD                                                        As ADODB.Recordset
Dim rsINCASH                                                        As ADODB.Recordset

Function SetBankName(XXX As Variant)
    Dim rsBankName                                                  As ADODB.Recordset
    Set rsBankName = New ADODB.Recordset
    Set rsBankName = gconDMIS.Execute("SELECT DescName FROM CMIS_SBOOK WHERE Book = 'B' AND Code = '" & XXX & "'")
    If Not rsBankName.EOF And Not rsBankName.BOF Then
        SetBankName = rsBankName!DESCNAME
    End If
    Set rsBankName = Nothing
End Function

Sub InitGrid()
    cleargrid grdPettyPay
    If TYPE_ON_HAND = "CARD" Then
        grdPettyPay.FormatString = " Customer Code | Customer Name                               |    Time            | Card Amount    "
        Label45.Caption = "Card Amount"
    ElseIf TYPE_ON_HAND = "CHECK" Then
        grdPettyPay.FormatString = " Code           |   Bank Name                                      |    Time            | Check Amount   "
        Label45.Caption = "Check Amount"
    End If
    grdPettyPay.ColWidth(4) = 1
End Sub

Sub StoreMemVars()
    'LAST UPDATE: 01/04/2006
    Dim TaoLang                                                     As Integer
    Dim TALA                                                        As Double
    
    Set rsOFF_HD = New ADODB.Recordset
    Set rsINCASH = New ADODB.Recordset
    If TYPE_ON_HAND = "CARD" Then
        If CASH_OPTIONS = "CASH_COL" Then
            If CASHPOSITION_CUTOFF_DATE <> LOGDATE Then
                Set rsOFF_HD = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE TOF = '3' AND OR_DATE = '" & CASHPOSITION_CUTOFF_DATE & "' AND cardamount > 0 ORDER BY ID ASC")
            Else
                Set rsOFF_HD = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE TOF = '3' AND Deposit = 0 AND cardamount > 0 ORDER BY ID ASC")
            End If
            InitGrid
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                rsOFF_HD.MoveFirst: TaoLang = 0: TALA = 0
                Do While Not rsOFF_HD.EOF
                    TaoLang = TaoLang + 1
                    'grdPettyPay.AddItem Null2String(rsOFF_HD!Cuscde) & Chr(9) & Null2String(rsOFF_HD!cusname) & Chr(9) & Null2String(rsOFF_HD!TimeCreate) & Chr(9) & ToDoubleNumber((N2Str2Zero(rsOFF_HD!cardamount) - (N2Str2Zero(rsOFF_HD!discount) + N2Str2Zero(rsOFF_HD!tax)))) & Chr(9) & rsOFF_HD!Id
                    'TALA = TALA + (N2Str2Zero(rsOFF_HD!cardamount) - (N2Str2Zero(rsOFF_HD!discount) + N2Str2Zero(rsOFF_HD!tax)))
                    grdPettyPay.AddItem Null2String(rsOFF_HD!CUSCDE) & Chr(9) & Null2String(rsOFF_HD!cusname) & Chr(9) & Null2String(rsOFF_HD!TimeCreate) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOFF_HD!CardAmount)) & Chr(9) & rsOFF_HD!Id
                    TALA = TALA + N2Str2Zero(rsOFF_HD!CardAmount)
                    If TaoLang = 1 Then grdPettyPay.RemoveItem 1
                    rsOFF_HD.MoveNext
                Loop
            End If
            txtTotalCollection.Text = ToDoubleNumber(TALA)
            Set rsOFF_HD = Nothing
        End If
    ElseIf TYPE_ON_HAND = "CHECK" Then
        If CASH_OPTIONS = "CASH_COL" Then
            If CASHPOSITION_CUTOFF_DATE <> LOGDATE Then
                Set rsOFF_HD = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE TOF = '2' AND OR_DATE = '" & CASHPOSITION_CUTOFF_DATE & "' AND chkamount > 0 ORDER BY ID ASC")
            Else
                Set rsOFF_HD = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE TOF = '2' AND Deposit = 0 AND chkamount > 0 ORDER BY ID ASC")
            End If
            InitGrid
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                rsOFF_HD.MoveFirst: TaoLang = 0: TALA = 0
                Do While Not rsOFF_HD.EOF
                    TaoLang = TaoLang + 1
                    grdPettyPay.AddItem Null2String(rsOFF_HD!bankcode) & Chr(9) & SetBankName(Null2String(rsOFF_HD!bankcode)) & Chr(9) & Null2String(rsOFF_HD!TimeCreate) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOFF_HD!CHKAMOUNT)) & Chr(9) & rsOFF_HD!Id
                    TALA = TALA + N2Str2Zero(rsOFF_HD!CHKAMOUNT)
                    If TaoLang = 1 Then grdPettyPay.RemoveItem 1
                    rsOFF_HD.MoveNext
                Loop
            End If
            txtTotalCollection.Text = ToDoubleNumber(TALA)
            Set rsOFF_HD = Nothing
        ElseIf CASH_OPTIONS = "CHECK_EN" Then
            Set rsINCASH = gconDMIS.Execute("SELECT * FROM CMIS_InCash WHERE Deposit = 0 AND CHKAMOUNT > 0 ORDER BY ID ASC")
            InitGrid
            If Not rsINCASH.EOF And Not rsINCASH.BOF Then
                rsINCASH.MoveFirst
                TaoLang = 0
                TALA = 0
                Do While Not rsINCASH.EOF
                    TaoLang = TaoLang + 1
                    grdPettyPay.AddItem Null2String(rsINCASH!bankcode) & Chr(9) & SetBankName(Null2String(rsINCASH!bankcode)) & Chr(9) & Null2String(rsINCASH!timeincash) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsINCASH!CHKAMOUNT)) & Chr(9) & rsINCASH!Id
                    TALA = TALA + N2Str2Zero(rsINCASH!CHKAMOUNT)
                    If TaoLang = 1 Then grdPettyPay.RemoveItem 1
                    rsINCASH.MoveNext
                Loop
            End If
            txtTotalCollection.Text = ToDoubleNumber(TALA)
            Set rsINCASH = Nothing
        End If
        'If CASH_OPTIONS = "PET_REPL" Then Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd Where Deposit = 0 and chkamount > 0 order by ID asc")
        'If CASH_OPTIONS = "LTO_REPL" Then Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd Where Deposit = 0 and chkamount > 0 order by ID asc")
        'If CASH_OPTIONS = "PET_ADV" Then Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd Where Deposit = 0 and chkamount > 0 order by ID asc")
    End If
End Sub

Sub StoreOFF_HDDetails(XXX As Variant)
    If CASH_OPTIONS = "CASH_COL" Then
        Set rsOFF_HD = New ADODB.Recordset
        Set rsOFF_HD = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE ID = " & XXX)
        If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
            If TYPE_ON_HAND = "CARD" Then
                txtChkDate.Text = Null2String(rsOFF_HD!carddate)
                txtChkNumber.Text = Null2String(rsOFF_HD!cardnumber)
                txtChkAmount.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!CardAmount))
            ElseIf TYPE_ON_HAND = "CHECK" Then
                txtChkDate.Text = Null2String(rsOFF_HD!CheckDate)
                txtChkNumber.Text = Null2String(rsOFF_HD!Tseke)
                txtChkAmount.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!CHKAMOUNT))
            End If
            txtOR_NUM.Text = Null2String(rsOFF_HD!OR_NUM)
        End If
    Else
        Set rsOFF_HD = New ADODB.Recordset
        Set rsOFF_HD = gconDMIS.Execute("SELECT * FROM CMIS_InCash WHERE ID = " & XXX)
        If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
            txtChkDate.Text = Null2String(rsOFF_HD!CHKDATE)
            txtChkNumber.Text = Null2String(rsOFF_HD!CHKNUMBER)
            txtChkAmount.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!CHKAMOUNT))
            txtOR_NUM.Text = Null2String(rsOFF_HD!incashdate)
        End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "]" '"." & App.Revision & "]"
    If CASH_OPTIONS = "CASH_COL" Then
        If TYPE_ON_HAND = "CARD" Then
            Label45.Caption = "Card Amount"
            Label46.Caption = "Card Date"
            Label47.Caption = "Card Number"
            Me.Caption = "CARD ON HAND - Cashier Collection"
        Else
            Me.Caption = "CHECK ON HAND - Cashier Collection"
        End If
    Else
        Label1.Caption = "Date of Incash"
        Me.Caption = "CHECK ON HAND - Check Encashment"
    End If
    InitGrid
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub grdPettyPay_Click()
    grdPettyPay.Col = 4
    If grdPettyPay.Text <> "" Then
        StoreOFF_HDDetails grdPettyPay.Text
        grdPettyPay.SetFocus
    End If
End Sub

Private Sub grdPettyPay_GotFocus()
    grdPettyPay.Col = 4
    If grdPettyPay.Text <> "" Then
        StoreOFF_HDDetails grdPettyPay.Text
        grdPettyPay.SetFocus
    End If
End Sub


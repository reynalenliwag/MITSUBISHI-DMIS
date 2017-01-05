VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.Ocx"
Begin VB.Form frmCASHPOSITIONTallySheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tally Sheet"
   ClientHeight    =   6195
   ClientLeft      =   870
   ClientTop       =   900
   ClientWidth     =   9000
   ForeColor       =   &H00F5F5F5&
   Icon            =   "TallySheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3105
      Left            =   60
      ScaleHeight     =   3105
      ScaleWidth      =   8895
      TabIndex        =   15
      Top             =   60
      Width           =   8895
      Begin VB.TextBox txtAvailableLTOFund 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   7140
         TabIndex        =   47
         Top             =   2700
         Width           =   1635
      End
      Begin VB.TextBox txtEnd 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   7140
         TabIndex        =   8
         Top             =   60
         Width           =   1635
      End
      Begin VB.TextBox txtAvailablePettyCashFund 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   7140
         TabIndex        =   14
         Top             =   2340
         Width           =   1635
      End
      Begin VB.TextBox txtCardCollection 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   7140
         TabIndex        =   13
         Top             =   1980
         Width           =   1635
      End
      Begin VB.TextBox txtCheckCollection 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   7140
         TabIndex        =   12
         Top             =   1620
         Width           =   1635
      End
      Begin VB.TextBox txtCardDepo 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   7140
         TabIndex        =   11
         Top             =   1260
         Width           =   1635
      End
      Begin VB.TextBox txtCheckDepo 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   7140
         TabIndex        =   10
         Top             =   900
         Width           =   1635
      End
      Begin VB.TextBox txtCashDepo 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   7140
         TabIndex        =   9
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox txtBegin 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   60
         Width           =   1635
      End
      Begin VB.TextBox txtAdvancesFromCollection 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Top             =   2700
         Width           =   1635
      End
      Begin VB.TextBox txtLTO 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   2340
         Width           =   1635
      End
      Begin VB.TextBox txtPettyCash 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   1980
         Width           =   1635
      End
      Begin VB.TextBox txtCash 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox txtCheck 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   900
         Width           =   1635
      End
      Begin VB.TextBox txtCard 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1260
         Width           =   1635
      End
      Begin VB.TextBox txtCashCollection 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   1620
         Width           =   1635
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6930
         TabIndex        =   49
         Top             =   2730
         Width           =   195
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Available L.T.O. Fund"
         Height          =   315
         Left            =   5010
         TabIndex        =   48
         Top             =   2730
         Width           =   1815
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Balance"
         Height          =   315
         Left            =   5010
         TabIndex        =   45
         Top             =   90
         Width           =   1815
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6930
         TabIndex        =   44
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Available Petty Cash Fund"
         Height          =   315
         Left            =   5010
         TabIndex        =   43
         Top             =   2370
         Width           =   1875
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6930
         TabIndex        =   42
         Top             =   2370
         Width           =   195
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Card Collection"
         Height          =   315
         Left            =   5010
         TabIndex        =   41
         Top             =   2010
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6930
         TabIndex        =   40
         Top             =   2010
         Width           =   195
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Collection"
         Height          =   315
         Left            =   5010
         TabIndex        =   39
         Top             =   1650
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6930
         TabIndex        =   38
         Top             =   1650
         Width           =   195
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Card Deposit"
         Height          =   315
         Left            =   5010
         TabIndex        =   37
         Top             =   1290
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6930
         TabIndex        =   36
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Deposit"
         Height          =   315
         Left            =   5010
         TabIndex        =   35
         Top             =   930
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6930
         TabIndex        =   34
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Deposit"
         Height          =   315
         Left            =   5010
         TabIndex        =   33
         Top             =   570
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6930
         TabIndex        =   32
         Top             =   570
         Width           =   195
      End
      Begin VB.Line Line1 
         X1              =   -60
         X2              =   8850
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Balance"
         Height          =   315
         Left            =   60
         TabIndex        =   31
         Top             =   90
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1950
         TabIndex        =   30
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Advances from Collection"
         Height          =   315
         Left            =   60
         TabIndex        =   29
         Top             =   2730
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1950
         TabIndex        =   28
         Top             =   2730
         Width           =   195
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total LTO"
         Height          =   315
         Left            =   60
         TabIndex        =   27
         Top             =   2370
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1950
         TabIndex        =   26
         Top             =   2370
         Width           =   195
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Petty Cash"
         Height          =   315
         Left            =   60
         TabIndex        =   25
         Top             =   2010
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1950
         TabIndex        =   24
         Top             =   2010
         Width           =   195
      End
      Begin VB.Label Label61 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1950
         TabIndex        =   23
         Top             =   570
         Width           =   195
      End
      Begin VB.Label Label62 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash on Hand"
         Height          =   315
         Left            =   60
         TabIndex        =   22
         Top             =   570
         Width           =   1815
      End
      Begin VB.Label Label65 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1950
         TabIndex        =   21
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Label66 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Check on Hand"
         Height          =   315
         Left            =   60
         TabIndex        =   20
         Top             =   930
         Width           =   1815
      End
      Begin VB.Label Label67 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1950
         TabIndex        =   19
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label Label68 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Card on Hand"
         Height          =   315
         Left            =   60
         TabIndex        =   18
         Top             =   1290
         Width           =   1815
      End
      Begin VB.Label Label69 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1950
         TabIndex        =   17
         Top             =   1650
         Width           =   195
      End
      Begin VB.Label Label70 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Collection"
         Height          =   315
         Left            =   60
         TabIndex        =   16
         Top             =   1650
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdOFF_HD 
      Height          =   2895
      Left            =   60
      TabIndex        =   46
      Top             =   3240
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorSel    =   -2147483633
      BackColorBkg    =   -2147483633
      Appearance      =   0
      MousePointer    =   99
      FormatString    =   "  OR No.          |    OR Date       |         Payee                        |   Type     |      OR Amount       "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "TallySheet.frx":030A
   End
End
Attribute VB_Name = "frmCASHPOSITIONTallySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCash_Pos, rsOFF_HD                                              As ADODB.Recordset
Attribute rsOFF_HD.VB_VarUserMemId = 1073938432
Dim TOTAL_CASH_COLL, TOTAL_CHECK_COLL, TOTAL_CARD_COLL                As Double
Attribute TOTAL_CASH_COLL.VB_VarUserMemId = 1073938434
Attribute TOTAL_CHECK_COLL.VB_VarUserMemId = 1073938434
Attribute TOTAL_CARD_COLL.VB_VarUserMemId = 1073938434

Sub StoreMemvars()
    Dim STUPID_ME                                                     As Integer
    Dim TOF_VAR                                                       As String
    Set rsCash_Pos = New ADODB.Recordset
    Set rsCash_Pos = gconDMIS.Execute("Select * from CMIS_Cash_Pos Where CUTDATE = '" & CASHPOSITION_CUTOFF_DATE & "'")
    If Not rsCash_Pos.EOF And Not rsCash_Pos.BOF Then
        txtBEGIN.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!Begin))
        txtCASH.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CASH))
        txtCHECK.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CHECK))
        txtCARD.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CARD))
        txtPETTYCASH.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!EXPENSE) + N2Str2Zero(rsCash_Pos!ADVANCES) + N2Str2Zero(rsCash_Pos!REPLENISH))
        txtLTO.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!LTO_EXP) + N2Str2Zero(rsCash_Pos!LTO_ADV) + N2Str2Zero(rsCash_Pos!LTO_REPL))
        txtCASHDEPO.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CashDepo))
        txtCHECKDEPO.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CheckDepo))
        txtCARDDEPO.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CardDepo))

        If N2Str2Zero(rsCash_Pos!FUND) < NumericVal(txtPETTYCASH.Text) Then
            txtAvailablePettyCashFund.Text = "0.00"
        Else
            txtAvailablePettyCashFund.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!FUND) - NumericVal(txtPETTYCASH.Text))
        End If
        If N2Str2Zero(rsCash_Pos!LTO) < NumericVal(txtLTO.Text) Then
            txtAvailableLTOFund.Text = "0.00"
        Else
            txtAvailableLTOFund.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!LTO) - NumericVal(txtLTO.Text))
        End If
        Set rsOFF_HD = New ADODB.Recordset
        'Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd WHERE CUTDATE IS NULL AND OR_AMT > 0 order by OR_NUM asc")
        'Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd WHERE CANCEL = 0 AND CUTDATE = '" & CASHPOSITION_CUTOFF_DATE & "' AND OR_AMT > 0 order by OR_NUM asc")
        'If rsOFF_HD.EOF And rsOFF_HD.BOF Then
        Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd WHERE CANCEL = 0 AND (CUTDATE = '" & CASHPOSITION_CUTOFF_DATE & "' OR OR_DATE = '" & CASHPOSITION_CUTOFF_DATE & "') AND OR_AMT > 0 order by OR_NUM asc")
        'End If
        TOTAL_CASH_COLL = 0: TOTAL_CHECK_COLL = 0: TOTAL_CARD_COLL = 0: STUPID_ME = 0
        If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
            rsOFF_HD.MoveFirst:
            Do While Not rsOFF_HD.EOF
                STUPID_ME = STUPID_ME + 1
                'Set rsOFF_DT = New ADODB.Recordset
                'Set rsOFF_DT = gconDMIS.Execute("Select SUM(PAYMENT) as TotalPayment from CMIS_Off_Dt Where OR_NUM = " & N2Str2Null(rsOFF_HD!OR_NUM))
                'If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                '   If Null2String(rsOFF_HD!TOF) = "1" Then
                '      TOTAL_CASH_COLL = TOTAL_CASH_COLL + N2Str2Zero(rsOFF_DT!TotalPayment)
                '   End If
                '   If Null2String(rsOFF_HD!TOF) = "2" Then
                '      TOTAL_CHECK_COLL = TOTAL_CHECK_COLL + N2Str2Zero(rsOFF_DT!TotalPayment)
                '   End If
                '   If Null2String(rsOFF_HD!TOF) = "3" Then
                '      TOTAL_CARD_COLL = TOTAL_CARD_COLL + N2Str2Zero(rsOFF_DT!TotalPayment)
                '   End If
                'End If
                If Null2String(rsOFF_HD!TOF) = "1" Then
                    TOF_VAR = "CASH"
                    TOTAL_CASH_COLL = TOTAL_CASH_COLL + N2Str2Zero(rsOFF_HD!OR_AMT)
                End If
                If Null2String(rsOFF_HD!TOF) = "2" Then
                    TOF_VAR = "CHECK"
                    TOTAL_CHECK_COLL = TOTAL_CHECK_COLL + N2Str2Zero(rsOFF_HD!OR_AMT)
                End If
                If Null2String(rsOFF_HD!TOF) = "3" Then
                    TOF_VAR = "CARD"
                    TOTAL_CARD_COLL = TOTAL_CARD_COLL + N2Str2Zero(rsOFF_HD!OR_AMT)
                End If
                grdOFF_HD.AddItem Null2String(rsOFF_HD!OR_NUM) & Chr(9) & Null2String(rsOFF_HD!OR_DATE) & Chr(9) & Null2String(rsOFF_HD!cusname) & Chr(9) & TOF_VAR & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOFF_HD!OR_AMT))
                If STUPID_ME = 1 Then grdOFF_HD.RemoveItem 1
                rsOFF_HD.MoveNext
            Loop
        End If
        txtCashCollection.Text = ToDoubleNumber(TOTAL_CASH_COLL)
        txtCheckCollection.Text = ToDoubleNumber(TOTAL_CHECK_COLL)
        txtCardCollection.Text = ToDoubleNumber(TOTAL_CARD_COLL)

        If N2Str2Zero(rsCash_Pos!LTO) < NumericVal(txtLTO.Text) Then
            txtAdvancesFromCollection.Text = ToDoubleNumber(Abs(N2Str2Zero(rsCash_Pos!LTO) - NumericVal(txtLTO.Text)))
        Else
            txtAdvancesFromCollection.Text = "0.00"
        End If
        If N2Str2Zero(rsCash_Pos!FUND) < NumericVal(txtPETTYCASH.Text) Then
            txtAdvancesFromCollection.Text = ToDoubleNumber(NumericVal(txtAdvancesFromCollection.Text) + (Abs(N2Str2Zero(rsCash_Pos!FUND) - NumericVal(txtPETTYCASH.Text))))
        End If
        txtCASH.Text = ToDoubleNumber(NumericVal(txtCASH.Text) - NumericVal(txtAdvancesFromCollection.Text))
        txtEND.Text = ToDoubleNumber((NumericVal(txtBEGIN.Text) + TOTAL_CASH_COLL + TOTAL_CHECK_COLL + TOTAL_CARD_COLL) - (NumericVal(txtCASHDEPO.Text) + NumericVal(txtCHECKDEPO.Text) + NumericVal(txtCARDDEPO.Text)))
    End If
    Set rsCash_Pos = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    StoreMemvars
    Screen.MousePointer = 0
End Sub


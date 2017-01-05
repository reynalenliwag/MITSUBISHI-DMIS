VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReconcileAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reconcile Account"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   ForeColor       =   &H00E0E0E0&
   Icon            =   "ReconcileAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   8685
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6840
      TabIndex        =   16
      Top             =   2310
      Width           =   1665
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5190
      TabIndex        =   15
      Top             =   2310
      Width           =   1665
   End
   Begin VB.CommandButton cmdOption 
      Caption         =   "&Beginning Bank Recon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2700
      TabIndex        =   19
      Top             =   2310
      Width           =   2505
   End
   Begin VB.TextBox txtEndingBal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6810
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   1710
      Width           =   1695
   End
   Begin VB.TextBox txtOpeningBal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6810
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   1260
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtLast 
      Height          =   315
      Left            =   2640
      TabIndex        =   11
      Top             =   1350
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   48758785
      CurrentDate     =   40002
   End
   Begin VB.ComboBox cboBank 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   900
      Width           =   3315
   End
   Begin MSComCtl2.DTPicker dtCurrent 
      Height          =   315
      Left            =   2640
      TabIndex        =   12
      Top             =   1770
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   48758785
      CurrentDate     =   40002
   End
   Begin VB.Label lblBank 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6090
      TabIndex        =   18
      Top             =   900
      Width           =   2415
   End
   Begin VB.Label lblBankName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   180
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   9
      Left            =   5010
      TabIndex        =   9
      Top             =   1830
      Width           =   135
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   8
      Left            =   5010
      TabIndex        =   8
      Top             =   1380
      Width           =   135
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   7
      Left            =   90
      TabIndex        =   7
      Top             =   930
      Width           =   135
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Balance:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   5160
      TabIndex        =   6
      Top             =   1830
      Width           =   1380
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   5160
      TabIndex        =   5
      Top             =   1380
      Width           =   1515
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account to Reconcile:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   930
      Width           =   2340
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Statement Date:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2145
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Statement Date:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1410
      Width           =   1830
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the details from your paper statement in appropriate fields, and then click Next."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   420
      Width           =   7365
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reconcile Account"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   150
      Width           =   1785
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   825
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   8715
   End
End
Attribute VB_Name = "frmReconcileAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xEnding                                            As Double

Private Sub cbobank_Click()
    Dim rsAllBank                                      As ADODB.Recordset
    Set rsAllBank = gconDMIS.Execute("select * from All_Banks where BANKNAME = '" & cboBank.Text & "'")
    If Not rsAllBank.EOF And Not rsAllBank.BOF Then
        lblBankName.Caption = Null2String(rsAllBank!BankName)
        lblBank.Caption = Null2String(rsAllBank!BankAcctNo)
        txtOpeningBal.Text = ToDoubleNumber(N2Str2Zero(rsAllBank!STARTING_BALANCE))
        xEnding = ToDoubleNumber(N2Str2Zero(rsAllBank!ENDING_BALANCE))
        txtEndingBal.Text = "0.00"

        If Null2Date(rsAllBank!LASTDATE_RECON) = "" Then
            dtCurrent = LOGDATE
            dtLast.Visible = False
            Label(2).Visible = False
            Dim rsFirstMonth                           As ADODB.Recordset
            Set rsFirstMonth = gconDMIS.Execute("select MIN(JDATE) as FirstDate from AMIS_vw_RECONDATA where BankAcctNo = '" & cboBank.Text & "'")
            If Not rsFirstMonth.EOF And Not rsFirstMonth.BOF Then
                dtCurrent = lastDay(Null2Date(rsFirstMonth!FirstDate))
                'dtCurrent = lastDay(dtCurrent)
            End If
            Set rsFirstMonth = Nothing
        Else
            dtLast = Null2Date(rsAllBank!LASTDATE_RECON)
        End If

        If dtLast = lastDay(Null2Date(rsAllBank!LASTDATE_RECON)) Then
            If Null2Date(rsAllBank!LASTDATE_RECON) <> "" Then
                dtLast.Visible = True
                Label(2).Visible = True
            End If
            dtCurrent = lastDay(Null2Date(DateAdd("m", 1, rsAllBank!LASTDATE_RECON)))
        ElseIf Null2Date(rsAllBank!LASTDATE_RECON) <> "" Then
            dtLast.Visible = True
            Label(2).Visible = True
            '            dtLast = Null2Date(rsAllBank!LASTDATE_RECON)
            dtCurrent = lastDay(Null2Date(rsAllBank!LASTDATE_RECON))
        End If
        '        dtCurrent.SetFocus
    End If
    Set rsAllBank = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If cboBank.Text = "" Then
        MessagePop InfoWarning, "Entry required...", "Please select Bank Account"
    ElseIf txtEndingBal.Text <= 0 Then
        MessagePop InfoWarning, "Entry required...", "Please enter Statement Ending Balance"
    Else
        '    Dim rsStarting As ADODB.Recordset
        '    Set rsStarting = gconDMIS.Execute("select * from All_Banks where BankAcctNo = '" & cboBank.Text & "'")
        '    If Not rsStarting.EOF And Not rsStarting.BOF Then
        '
        '    End If
        '    ReconAccount = Me.cboBank.Text
        '    ReconDate = Me.dtCurrent.Value
        '    ReconBankName = "Reconcile Account - " + lblBankName.Caption
        rEndingBalance = ToDoubleNumber(Me.txtEndingBal.Text)
        FrmBankReconNew.lblAccount.Caption = Me.lblBank.Caption
        FrmBankReconNew.lblDateAsOf.Caption = Me.dtCurrent.Value
        FrmBankReconNew.lblBank.Caption = "Reconcile Account - " + lblBankName.Caption
        frmReconcileAccount.ZOrder 1
    End If
End Sub

Private Sub cmdOption_Click()
    If cboBank.Text = "" Then
        MsgBox "Kindly select bank.", vbExclamation, "Select Bank"
        Exit Sub
    Else
        frmBankReconBeginning.Show
    End If
End Sub

Private Sub dtCurrent_LostFocus()
    Dim rsAllBank                                      As ADODB.Recordset
    Set rsAllBank = gconDMIS.Execute("select * from All_Banks where BANKACCTNO = '" & cboBank.Text & "'")
    If Not rsAllBank.EOF And Not rsAllBank.BOF Then
        If Null2Date(rsAllBank!LASTDATE_RECON) <> "" Then
            If dtCurrent.Value <= dtLast.Value Then
                MessagePop InfoWarning, "Bank Reconcillation", "Date selected has been reconciled"
                dtCurrent.SetFocus
            End If
            '            Dim rsFirstMonth As ADODB.Recordset
            '            Set rsFirstMonth = gconDMIS.Execute("select * from AMIS_vw_RECONDATA where BankAcctNo = '" & cboBank.Text & "'")
            '            If Not rsFirstMonth.EOF And Not rsFirstMonth.BOF Then
            '                dtCurrent = firstDay(Null2Date(rsFirstMonth!JDate))
            '                dtCurrent = lastDay(dtCurrent)
            '            End If
            '            Set rsFirstMonth = Nothing
            '        Else
            '            If dtCurrent = lastDay(Null2Date(DateAdd("m", 1, rsAllBank!LASTDATE_RECON))) Then
            '                Exit Sub
            '            ElseIf dtCurrent > lastDay(Null2Date(rsAllBank!LASTDATE_RECON)) Then
            '                MessagePop InfoWarning, "Bank Reconcillation", "Selected month must be reconciled first"
            '                dtCurrent.SetFocus
            '            End If
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        'Unload Me
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim rsLoadBank                                     As ADODB.Recordset
    Set rsLoadBank = New ADODB.Recordset
    rsLoadBank.Open "select BankName from ALL_BANKS", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsLoadBank.EOF And Not rsLoadBank.EOF Then
        cboBank.Clear
        Do While Not rsLoadBank.EOF
            cboBank.AddItem Null2String(rsLoadBank!BankName)
            rsLoadBank.MoveNext
        Loop
    End If
    dtLast.Value = firstDay(LOGDATE)
    dtCurrent.Value = lastDay(LOGDATE)
End Sub

Private Sub txtEndingBal_GotFocus()
    txtEndingBal.Text = NumericVal(txtEndingBal.Text)
End Sub

Private Sub txtEndingBal_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtEndingBal_LostFocus()
    txtEndingBal.Text = ToDoubleNumber(txtEndingBal.Text)
End Sub

Private Sub txtOpeningBal_GotFocus()
    txtOpeningBal.Text = NumericVal(txtOpeningBal.Text)
End Sub

Private Sub txtOpeningBal_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtOpeningBal_LostFocus()
    txtOpeningBal.Text = ToDoubleNumber(txtOpeningBal.Text)
End Sub

Function GetBankCode(xBankCode As String) As String
    Dim rsBankCode                                     As ADODB.Recordset
    Set rsBankCode = New ADODB.Recordset
    rsBankCode.Open "Select * from ALL_BANKS where BankAcctNo = '" & xBankCode & "'", gconDMIS, adOpenKeyset
    If Not rsBankCode.EOF And Not rsBankCode.BOF Then
        GetBankCode = rsBankCode!BankName
    End If
End Function

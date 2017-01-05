VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
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
      Left            =   6720
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
      Left            =   4950
      TabIndex        =   15
      Top             =   2310
      Width           =   1665
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
      Left            =   6720
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
      Left            =   6720
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
      Format          =   20709377
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
      Width           =   2205
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
      Format          =   20709377
      CurrentDate     =   40002
   End
   Begin VB.Label lblBank 
      Height          =   375
      Left            =   2460
      TabIndex        =   18
      Top             =   2250
      Visible         =   0   'False
      Width           =   1545
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
      Left            =   4920
      TabIndex        =   17
      Top             =   900
      Width           =   3495
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
      Left            =   4920
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
      Left            =   4920
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
      Left            =   330
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
      Left            =   5070
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
      Left            =   5070
      TabIndex        =   5
      Top             =   1380
      Width           =   1515
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account:"
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
      Left            =   480
      TabIndex        =   4
      Top             =   930
      Width           =   750
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
      Left            =   480
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
      Left            =   480
      TabIndex        =   2
      Top             =   1350
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
Dim xEnding As Double

Private Sub cbobank_Click()
    Dim rsAllBank As ADODB.Recordset
    Set rsAllBank = gconDMIS.Execute("select * from All_Banks where BANKACCTNO = '" & cboBank.Text & "'")
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
            Dim rsFirstMonth As ADODB.Recordset
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
            dtLast.Visible = True
            Label(2).Visible = True
            dtCurrent = lastDay(Null2Date(DateAdd("m", 1, rsAllBank!LASTDATE_RECON)))
        ElseIf Null2Date(rsAllBank!LASTDATE_RECON) <> "" Then
            dtLast.Visible = True
            Label(2).Visible = True
'            dtLast = Null2Date(rsAllBank!LASTDATE_RECON)
            dtCurrent = lastDay(Null2Date(rsAllBank!LASTDATE_RECON))
        End If
'        dtCurrent.SetFocus
    End If
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
    FrmBankReconNew.lblAccount.Caption = Me.cboBank.Text
    FrmBankReconNew.lblDateAsOf.Caption = Me.dtCurrent.Value
    FrmBankReconNew.lblBank.Caption = "Reconcile Account - " + lblBankName.Caption
    frmReconcileAccount.ZOrder 1
End If
End Sub

Private Sub dtCurrent_LostFocus()
    Dim rsAllBank As ADODB.Recordset
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
    rsLoadBank.Open "select BankAcctNo from ALL_BANKS", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsLoadBank.EOF And Not rsLoadBank.EOF Then
        cboBank.Clear
        Do Until rsLoadBank.EOF
            cboBank.AddItem Null2String(rsLoadBank!BankAcctNo)
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
'    KeyAscii = OnlyNumeric(KeyAscii)
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

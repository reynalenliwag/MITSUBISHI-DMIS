VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmViewCashPosition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Cash Position Receipt"
   ClientHeight    =   7200
   ClientLeft      =   405
   ClientTop       =   1710
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "ViewCashPosition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8925
   Begin VB.PictureBox picDetails 
      Height          =   3195
      Left            =   6030
      ScaleHeight     =   3135
      ScaleWidth      =   2685
      TabIndex        =   70
      Top             =   2610
      Width           =   2745
      Begin VB.CommandButton Command1 
         Caption         =   "Petty Cash Expenses..."
         Height          =   345
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   2625
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Petty Cash Advances..."
         Height          =   345
         Left            =   30
         TabIndex        =   5
         Top             =   420
         Width           =   2625
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Check in Bank..."
         Height          =   345
         Left            =   30
         TabIndex        =   6
         Top             =   810
         Width           =   2625
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Check on Hand..."
         Height          =   345
         Left            =   30
         TabIndex        =   7
         Top             =   1200
         Width           =   2625
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Card on Bank..."
         Height          =   345
         Left            =   30
         TabIndex        =   8
         Top             =   1590
         Width           =   2625
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Card on Hand..."
         Height          =   345
         Left            =   30
         TabIndex        =   9
         Top             =   1980
         Width           =   2625
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Tally Sheet..."
         Height          =   345
         Left            =   30
         TabIndex        =   10
         Top             =   2370
         Width           =   2625
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Cashier Cash Count..."
         Height          =   345
         Left            =   30
         TabIndex        =   11
         Top             =   2760
         Width           =   2625
      End
   End
   Begin wizButton.cmd cmdDetails 
      Height          =   3315
      Left            =   5970
      TabIndex        =   71
      Top             =   2550
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   5847
      TX              =   "cmd1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "ViewCashPosition.frx":030A
   End
   Begin VB.CommandButton cmdF11 
      Caption         =   "F11 - Calculator"
      Height          =   345
      Left            =   6960
      TabIndex        =   3
      ToolTipText     =   "View Calculator"
      Top             =   5910
      Width           =   1905
   End
   Begin VB.CommandButton cmdF9 
      Caption         =   "F9 - Details"
      Height          =   345
      Left            =   5490
      TabIndex        =   2
      ToolTipText     =   "View Details"
      Top             =   5910
      Width           =   1425
   End
   Begin VB.CommandButton cmdF6 
      Caption         =   "F6 - Check Payment for Petty Cash"
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Check Payment for Petty Cash"
      Top             =   5910
      Width           =   3765
   End
   Begin VB.CommandButton cmdF4 
      Caption         =   "F4 - Switch"
      Height          =   345
      Left            =   90
      TabIndex        =   0
      ToolTipText     =   "Switch"
      Top             =   5910
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5805
      Left            =   0
      ScaleHeight     =   5805
      ScaleWidth      =   8925
      TabIndex        =   12
      Top             =   0
      Width           =   8925
      Begin VB.TextBox txtCutDate 
         Height          =   315
         Left            =   2220
         TabIndex        =   30
         Top             =   90
         Width           =   1635
      End
      Begin VB.TextBox txtCASH 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   29
         Top             =   840
         Width           =   1635
      End
      Begin VB.TextBox txtCHECK 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   28
         Top             =   1170
         Width           =   1635
      End
      Begin VB.TextBox txtCARD 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   27
         Top             =   1500
         Width           =   1635
      End
      Begin VB.TextBox txtPETTYFUND 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4650
         TabIndex        =   26
         Top             =   3870
         Width           =   1635
      End
      Begin VB.TextBox txtPETTYCASH 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4650
         TabIndex        =   25
         Top             =   4200
         Width           =   1635
      End
      Begin VB.TextBox txtCARDDEPO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   24
         Top             =   1500
         Width           =   1635
      End
      Begin VB.TextBox txtCHECKDEPO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   23
         Top             =   1170
         Width           =   1635
      End
      Begin VB.TextBox txtCASHDEPO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   22
         Top             =   840
         Width           =   1635
      End
      Begin VB.TextBox txtADVANCES 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   21
         Top             =   3030
         Width           =   1635
      End
      Begin VB.TextBox txtEXPENSE 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   20
         Top             =   2700
         Width           =   1635
      End
      Begin VB.TextBox txtREPLENISH 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   19
         Top             =   2370
         Width           =   1635
      End
      Begin VB.TextBox txtEND 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   18
         Top             =   3030
         Width           =   1635
      End
      Begin VB.TextBox txtBEGIN 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   17
         Top             =   2700
         Width           =   1635
      End
      Begin VB.TextBox txtAR 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   16
         Top             =   2370
         Width           =   1635
      End
      Begin VB.TextBox txtPettyCAFromCollection 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4650
         TabIndex        =   15
         Top             =   5040
         Width           =   1635
      End
      Begin VB.TextBox txtRemainingPettyFund 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4650
         TabIndex        =   14
         Top             =   4710
         Width           =   1635
      End
      Begin VB.TextBox txtTotalAdvances 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4650
         TabIndex        =   13
         Top             =   5370
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cut-Off Date"
         Height          =   315
         Left            =   120
         TabIndex        =   69
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash on Hand"
         Height          =   315
         Left            =   120
         TabIndex        =   68
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Check on Hand"
         Height          =   315
         Left            =   120
         TabIndex        =   67
         Top             =   1170
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Card on Hand"
         Height          =   315
         Left            =   120
         TabIndex        =   66
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1980
         TabIndex        =   65
         Top             =   840
         Width           =   195
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1980
         TabIndex        =   64
         Top             =   1170
         Width           =   195
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1980
         TabIndex        =   63
         Top             =   1500
         Width           =   195
      End
      Begin VB.Label labMaximum 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Petty Cash Fund"
         Height          =   315
         Left            =   120
         TabIndex        =   62
         Top             =   3870
         Width           =   3375
      End
      Begin VB.Label labTotalExpenses 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Petty Cash Expenses"
         Height          =   315
         Left            =   120
         TabIndex        =   61
         Top             =   4200
         Width           =   3375
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   4410
         TabIndex        =   60
         Top             =   3870
         Width           =   195
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   4410
         TabIndex        =   59
         Top             =   4200
         Width           =   195
      End
      Begin VB.Label Label30 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1980
         TabIndex        =   58
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6360
         TabIndex        =   57
         Top             =   1500
         Width           =   195
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6360
         TabIndex        =   56
         Top             =   1170
         Width           =   195
      End
      Begin VB.Label Label29 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6360
         TabIndex        =   55
         Top             =   840
         Width           =   195
      End
      Begin VB.Label Label31 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Card Deposit"
         Height          =   315
         Left            =   4500
         TabIndex        =   54
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label34 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Deposit"
         Height          =   315
         Left            =   4500
         TabIndex        =   53
         Top             =   1170
         Width           =   1815
      End
      Begin VB.Label Label38 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Deposit"
         Height          =   315
         Left            =   4500
         TabIndex        =   52
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1980
         TabIndex        =   51
         Top             =   3030
         Width           =   195
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1980
         TabIndex        =   50
         Top             =   2700
         Width           =   195
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1980
         TabIndex        =   49
         Top             =   2370
         Width           =   195
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Advances"
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   3030
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Expense"
         Height          =   315
         Left            =   120
         TabIndex        =   47
         Top             =   2700
         Width           =   1815
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Replenish"
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   2370
         Width           =   1815
      End
      Begin VB.Label Label33 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6360
         TabIndex        =   45
         Top             =   3030
         Width           =   195
      End
      Begin VB.Label Label35 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6360
         TabIndex        =   44
         Top             =   2700
         Width           =   195
      End
      Begin VB.Label Label36 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   6360
         TabIndex        =   43
         Top             =   2370
         Width           =   195
      End
      Begin VB.Label Label37 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Balance"
         Height          =   315
         Left            =   4500
         TabIndex        =   42
         Top             =   3030
         Width           =   1815
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Balance"
         Height          =   315
         Left            =   4500
         TabIndex        =   41
         Top             =   2700
         Width           =   1815
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Receivable"
         Height          =   315
         Left            =   4500
         TabIndex        =   40
         Top             =   2370
         Width           =   1815
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   4410
         TabIndex        =   39
         Top             =   5040
         Width           =   195
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   4410
         TabIndex        =   38
         Top             =   4710
         Width           =   195
      End
      Begin VB.Label labCashAdvances 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Cash Advances from Collection"
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   5040
         Width           =   4065
      End
      Begin VB.Label labRemaining 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining Petty Cash Fund"
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   4710
         Width           =   4065
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   4410
         TabIndex        =   35
         Top             =   5370
         Width           =   195
      End
      Begin VB.Label labTotalAdvances 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Advances from Collection"
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   5370
         Width           =   4065
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   60
         X2              =   8715
         Y1              =   1890
         Y2              =   1905
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   60
         X2              =   8715
         Y1              =   3420
         Y2              =   3435
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Summary of Collection"
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   540
         Width           =   4095
      End
      Begin VB.Label labBreakDown 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Breakdown of Petty Cash"
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label labFundStatusMonitoring 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Cash Fund Status Monitoring"
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   3540
         Width           =   4095
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   4620
         X2              =   6300
         Y1              =   4590
         Y2              =   4590
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   780
      MouseIcon       =   "ViewCashPosition.frx":0326
      MousePointer    =   99  'Custom
      Picture         =   "ViewCashPosition.frx":0478
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Move to Next Record"
      Top             =   6330
      Width           =   705
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Prev"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   90
      MouseIcon       =   "ViewCashPosition.frx":07D0
      MousePointer    =   99  'Custom
      Picture         =   "ViewCashPosition.frx":0922
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "Move to Previous Record"
      Top             =   6330
      Width           =   705
   End
End
Attribute VB_Name = "frmViewCashPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCash_Pos                                                        As ADODB.Recordset

Sub rsRefresh()
    Set rsCash_Pos = New ADODB.Recordset
    Set rsCash_Pos = gconDMIS.Execute("Select * from CMIS_Cash_Pos order by CUTDATE asC")
End Sub

Sub StoreMemVars()
    If Not rsCash_Pos.EOF And Not rsCash_Pos.BOF Then
        txtCutDate.Text = Null2String(rsCash_Pos!CUTDATE)
        'If txtCutDate.Text <> CURRENT_CUTOFF_DATE Then
        '    cmdF6.Enabled = False: cmdF9.Enabled = False
        'Else
        cmdF6.Enabled = True: cmdF9.Enabled = True
        'End If
        CASHPOSITION_CUTOFF_DATE = Null2Date(rsCash_Pos!CUTDATE)
        txtCASH.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CASH))
        txtCHECK.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CHECK))
        txtCARD.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CARD))

        txtCASHDEPO.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CashDepo))
        txtCHECKDEPO.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CheckDepo))
        txtCARDDEPO.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CardDepo))

        txtAR.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!ar))
        txtBEGIN.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!Begin))
        txtEND.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!End))

        If IsLTOIsPettyCash = "LTO" Then
            txtREPLENISH.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!LTO_REPL))
            txtEXPENSE.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!LTO_EXP))
            txtADVANCES.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!LTO_ADV))

            txtPETTYFUND.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!LTO))
            txtPETTYCASH.Text = ToDoubleNumber(NumericVal(txtREPLENISH.Text) + NumericVal(txtEXPENSE.Text) + NumericVal(txtADVANCES.Text))

            If N2Str2Zero(rsCash_Pos!LTO) < NumericVal(txtPETTYCASH.Text) Then
                txtRemainingPettyFund.Text = "0.00"
                txtPettyCAFromCollection.Text = ToDoubleNumber(NumericVal(txtPETTYCASH.Text) - N2Str2Zero(rsCash_Pos!LTO))
            Else
                txtRemainingPettyFund.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!LTO) - NumericVal(txtPETTYCASH.Text))
                txtPettyCAFromCollection.Text = "0.00"
            End If
        Else
            txtREPLENISH.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!REPLENISH))
            txtEXPENSE.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!EXPENSE))
            txtADVANCES.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!ADVANCES))

            txtPETTYFUND.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!FUND))
            txtPETTYCASH.Text = ToDoubleNumber(NumericVal(txtREPLENISH.Text) + NumericVal(txtEXPENSE.Text) + NumericVal(txtADVANCES.Text))
            If N2Str2Zero(rsCash_Pos!FUND) < NumericVal(txtPETTYCASH.Text) Then
                txtRemainingPettyFund.Text = "0.00"
                txtPettyCAFromCollection.Text = ToDoubleNumber(NumericVal(txtPETTYCASH.Text) - N2Str2Zero(rsCash_Pos!FUND))
            Else
                txtRemainingPettyFund.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!FUND) - NumericVal(txtPETTYCASH.Text))
                txtPettyCAFromCollection.Text = "0.00"
            End If
        End If
        If N2Str2Zero(rsCash_Pos!LTO) < (N2Str2Zero(rsCash_Pos!LTO_EXP) + N2Str2Zero(rsCash_Pos!LTO_ADV) + N2Str2Zero(rsCash_Pos!LTO_REPL)) Then
            txtTotalAdvances.Text = ToDoubleNumber((N2Str2Zero(rsCash_Pos!LTO_EXP) + N2Str2Zero(rsCash_Pos!LTO_ADV) + N2Str2Zero(rsCash_Pos!LTO_REPL)) - N2Str2Zero(rsCash_Pos!LTO))
        Else
            txtTotalAdvances.Text = "0.00"
        End If
        If N2Str2Zero(rsCash_Pos!FUND) < N2Str2Zero(rsCash_Pos!REPLENISH) + N2Str2Zero(rsCash_Pos!EXPENSE) + N2Str2Zero(rsCash_Pos!ADVANCES) Then
            txtTotalAdvances.Text = ToDoubleNumber(((N2Str2Zero(rsCash_Pos!REPLENISH) + N2Str2Zero(rsCash_Pos!EXPENSE) + N2Str2Zero(rsCash_Pos!ADVANCES)) - N2Str2Zero(rsCash_Pos!FUND)) + NumericVal(txtTotalAdvances.Text))
        End If
        If VALID_COMPANY_CODE_FORHAI = True Then
        Else
            If N2Str2Zero(rsCash_Pos!LTO) = 0 Then cmdF4.Enabled = False
        End If
        txtCASH.Text = ToDoubleNumber(NumericVal(txtCASH.Text) - NumericVal(txtTotalAdvances.Text))
    End If
End Sub

Private Sub cmdF11_Click()
    Shell "calc.exe"
End Sub

Private Sub cmdF4_Click()

    If IsLTOIsPettyCash = "PETTY" Then
        IsLTOIsPettyCash = "LTO"
        cmdF6.Caption = "F6 - Check Payment for L.T.O."
        Command1.Caption = "LTO Expenses..."
        Command2.Caption = "LTO Advances..."
        labBreakDown.Caption = "Breakdown of L.T.O."
        labFundStatusMonitoring.Caption = "L.T.O. Fund Status Monitoring"
        labMaximum.Caption = "Maximum L.T.O. Registration Fund"
        labTotalExpenses.Caption = "Total L.T.O. Expenses"
        labRemaining.Caption = "Remaining L.T.O. Fund"
        labCashAdvances.Caption = "L.T.O. Advances from Collection"
        labTotalAdvances.Caption = "Total Advances from Collection"
    Else
        IsLTOIsPettyCash = "PETTY"
        cmdF6.Caption = "F6 - Check Payment for Petty Cash"
        Command1.Caption = "Petty Cash Expenses"
        Command2.Caption = "Petty Cash Advances"
        labBreakDown.Caption = "Breakdown of Petty Cash"
        labFundStatusMonitoring.Caption = "Petty Cash Fund Status Monitoring"
        labMaximum.Caption = "Maximum Petty Cash Fund"
        labTotalExpenses.Caption = "Total Petty Cash Expenses"
        labRemaining.Caption = "Remaining Petty Cash Fund"
        labCashAdvances.Caption = "Petty Cash Advances from Collection"
        labTotalAdvances.Caption = "Total Advances from Collection"
    End If
    StoreMemVars
End Sub

Private Sub cmdF6_Click()
    If IsLTOIsPettyCash = "LTO" Then
        frmCASHPOSITIONCheckPaymentForLTO.Show vbModal
    Else
        frmCASHPOSITIONCheckPaymentForPettyCash.Show vbModal
    End If
    StoreMemVars
End Sub

Private Sub cmdF9_Click()

    cmdDetails.ZOrder 0: cmdDetails.Visible = True
    picDetails.ZOrder 0: picDetails.Visible = True
    On Error Resume Next
    Command1.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsCash_Pos.MoveNext
    If rsCash_Pos.EOF Then
        rsCash_Pos.MoveLast
        ShowLastRecordMsg
        'MsgBox "Last Record!"
    End If
    StoreMemVars
End Sub

Private Sub cmdPrev_Click()
    rsCash_Pos.MovePrevious
    If rsCash_Pos.BOF Then
        rsCash_Pos.MoveFirst
        MsgBox "First Record!"
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsCash_Pos.MovePrevious
    If rsCash_Pos.BOF Then
        rsCash_Pos.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub Command1_Click()
    If IsLTOIsPettyCash = "LTO" Then
        frmCASHPOSITIONLTOExpenses.Show vbModal
        rsRefresh
        rsCash_Pos.Find "CUTDATE = '" & txtCutDate.Text & "'"
        StoreMemVars
    Else
        frmCASHPOSITIONPettyCashExpenses.Show vbModal
        rsRefresh
        rsCash_Pos.Find "CUTDATE = '" & txtCutDate.Text & "'"
        StoreMemVars
    End If
End Sub

Private Sub Command2_Click()
    If IsLTOIsPettyCash = "LTO" Then
        frmCASHPOSITIONLTOAdvances.Show vbModal
    Else
        frmCASHPOSITIONPettyCashAdvances.Show vbModal
    End If
End Sub

Private Sub Command4_Click()
    frmCASHPOSITIONCheckInBank.Show vbModal
End Sub

Private Sub Command5_Click()
    TYPE_ON_HAND = "CHECK"
    frmCASHPOSITIONSelectCashOptions.Show vbModal
End Sub

Private Sub Command6_Click()
    frmCASHPOSITIONCardInBank.Show vbModal
End Sub

Private Sub Command7_Click()
    TYPE_ON_HAND = "CARD"
    frmCASHPOSITIONSelectCashOptions.Show vbModal
End Sub

Private Sub Command8_Click()
    frmCASHPOSITIONTallySheet.Show vbModal
End Sub

Private Sub Command9_Click()
    frmCASHPOSITIONCashCount.Show vbModal
    StoreMemVars
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If picDetails.Visible = True Then
                cmdDetails.ZOrder 1: cmdDetails.Visible = False
                picDetails.ZOrder 1: picDetails.Visible = False
            Else
                Unload Me
            End If
        Case vbKeyF4
            cmdF4.Value = True
        Case vbKeyF5
            frmCASHPOSITIONCheckPaymentForPettyCash.Show vbModal
        Case vbKeyF6
            If cmdF6.Enabled = True Then cmdF6.Value = True
        Case vbKeyF9
            If cmdF9.Enabled = True Then cmdF9.Value = True
        Case vbKeyF11
            cmdF11.Value = True
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Dim rsProfile                                                     As ADODB.Recordset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_Profile WHERE MODULENAME = 'CMIS'")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        PERIODMONTH = N2Str2Zero(rsProfile!PERIODMONTH)
        PERIODYEAR = N2Str2Zero(rsProfile!PERIODYEAR)
    Else
        PERIODMONTH = Month(Now)
        PERIODYEAR = Year(Now)
    End If
    Set rsProfile = Nothing
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    cmdDetails.ZOrder 1: cmdDetails.Visible = False
    picDetails.ZOrder 1: picDetails.Visible = False
    rsRefresh
    If Not rsCash_Pos.EOF And Not rsCash_Pos.BOF Then rsCash_Pos.MoveLast
    StoreMemVars
    IsLTOIsPettyCash = "PETTY"
End Sub


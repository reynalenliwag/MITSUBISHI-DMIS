VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMISAORCalculator 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6510
   ClientLeft      =   2145
   ClientTop       =   480
   ClientWidth     =   8820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AORCalculator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboTerms 
      Height          =   345
      Left            =   1875
      TabIndex        =   16
      Top             =   2355
      Width           =   1665
   End
   Begin VB.TextBox txtTotalInterest 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6030
      Width           =   2115
   End
   Begin VB.TextBox txtTotalPrincipal 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   6165
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6030
      Width           =   2265
   End
   Begin VB.ComboBox cboVDetails 
      Height          =   345
      Left            =   1860
      TabIndex        =   2
      Top             =   90
      Width           =   3765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7875
      TabIndex        =   19
      Top             =   2250
      Width           =   840
   End
   Begin VB.TextBox txtBaltofin 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   1875
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1740
      Width           =   3690
   End
   Begin VB.TextBox txtDownPayment 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   1875
      TabIndex        =   14
      Top             =   1335
      Width           =   3690
   End
   Begin VB.TextBox txtAOR 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   4080
      TabIndex        =   17
      Top             =   2355
      Width           =   1440
   End
   Begin VB.TextBox txtNetSalesAmount 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   1875
      TabIndex        =   13
      Top             =   930
      Width           =   3690
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "OK"
      Height          =   375
      Left            =   6975
      TabIndex        =   18
      Top             =   2250
      Width           =   765
   End
   Begin VB.TextBox txtNetMonthlyAmorization 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1260
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   300
      Width           =   3090
   End
   Begin MSFlexGridLib.MSFlexGrid gridOne 
      Height          =   3105
      Left            =   75
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2775
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   5477
      _Version        =   393216
      Cols            =   5
      BackColorBkg    =   -2147483633
      FocusRect       =   2
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   $"AORCalculator.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport rptAOR 
      Left            =   4200
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtFinComp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   1860
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   495
      Width           =   2790
   End
   Begin VB.CommandButton cmdViewSelect 
      Caption         =   "SELECT"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4650
      TabIndex        =   4
      Top             =   495
      Width           =   930
   End
   Begin VB.PictureBox picViewVehicles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   1875
      ScaleHeight     =   3765
      ScaleWidth      =   3675
      TabIndex        =   23
      Top             =   870
      Visible         =   0   'False
      Width           =   3705
      Begin XtremeReportControl.ReportControl lvViewVehicles 
         Height          =   2370
         Left            =   60
         TabIndex        =   8
         Top             =   900
         Width           =   3555
         _Version        =   655364
         _ExtentX        =   6271
         _ExtentY        =   4180
         _StockProps     =   64
         BorderStyle     =   4
         MultipleSelection=   0   'False
      End
      Begin VB.OptionButton optRular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Rural Rate"
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   630
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton optUrban 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Urban Rate"
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2130
         TabIndex        =   7
         Top             =   630
         Width           =   1590
      End
      Begin VB.ComboBox cboFinCom 
         Height          =   345
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   285
         Width           =   3465
      End
      Begin VB.CommandButton cmdCancelView2 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2775
         TabIndex        =   10
         Top             =   3330
         Width           =   825
      End
      Begin VB.CommandButton cmdSelectViewVehicles 
         Caption         =   "Select "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1890
         TabIndex        =   9
         Top             =   3330
         Width           =   825
      End
      Begin VB.Label lblOne 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Financing Co "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   150
         TabIndex        =   24
         Top             =   0
         Width           =   2505
      End
   End
   Begin VB.Label labPayments 
      Caption         =   "Total Interest"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   75
      TabIndex        =   28
      Top             =   6075
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Total Principal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   4260
      TabIndex        =   30
      Top             =   6075
      Width           =   1815
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Vehicle Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   60
      TabIndex        =   11
      Top             =   150
      Width           =   1260
   End
   Begin VB.Label Label3 
      Caption         =   "Balance To be Financed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   525
      Left            =   60
      TabIndex        =   21
      Top             =   1740
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Down Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   60
      TabIndex        =   20
      Top             =   1410
      Width           =   1275
   End
   Begin VB.Label lblRate 
      AutoSize        =   -1  'True
      Caption         =   "AOR::"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   4095
      TabIndex        =   25
      Top             =   2145
      Width           =   465
   End
   Begin VB.Label lblMonths 
      AutoSize        =   -1  'True
      Caption         =   "Terms:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   1860
      TabIndex        =   22
      Top             =   2145
      Width           =   600
   End
   Begin VB.Label lblPrincipal 
      AutoSize        =   -1  'True
      Caption         =   "Net Sales Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   60
      TabIndex        =   12
      Top             =   1005
      Width           =   1305
   End
   Begin VB.Label lblPayment 
      AutoSize        =   -1  'True
      Caption         =   "Net Monthly Mortization"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   5760
      TabIndex        =   0
      Top             =   60
      Width           =   1980
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Financing Company"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   585
      Width           =   1650
   End
End
Attribute VB_Name = "frmSMISAORCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFinCom_Click()
    Dim xPerct                                    As String
    If cboFinCom.ListIndex = -1 Then: Exit Sub
    xPerct = IIf(optRular.Value = True, "ISNULL(Rperct, 0)", " ISNULL(Uperct , 0)")
    flex_FillReportView gconDMIS.Execute("SELECT cast(TERM as varchar) + ' Months' , " & xPerct & " from SMIS_FINCOM_RATE where  FINCOMID= " & cboFinCom.ItemData(cboFinCom.ListIndex)), lvViewVehicles

End Sub

Private Sub cboTerms_Change()
    Command1.Enabled = False
    txtAOR_Change
End Sub

Private Sub cboTerms_Validate(Cancel As Boolean)
    If IsNumeric(cboTerms) = False Then
        Cancel = True
    End If
End Sub

Private Sub cmdCalculate_Click()

    Dim termno                                    As Integer
    Dim payment                                   As Double
    Dim interest_Total                            As Double
    Dim payment_Total                             As Double
    Dim Principal                                 As Currency
    Dim Term                                      As Integer
    Dim Interest                                  As Double
    Dim intMonth                                  As Integer
    Dim ytdInterest                               As Double
    Dim ytdApplied                                As Double
    Dim MonthlyInterest                           As Double



    Principal = CCur(txtBaltofin.Text)

    Term = Int(cboTerms.Text)
    Interest = Val(txtAOR.Text)

    If Principal = 0 Then
        txtNetMonthlyAmorization.Text = 0
    Else
        If Interest = 0 Then
            Interest = 0
            payment = (Principal / Term)


        Else
            Interest = Interest / 1200
            payment = (Principal * Interest / (1 - ((1 / (1 + Interest) ^ Term))))
        End If
        txtNetMonthlyAmorization.Text = FormatCurrency(payment)
    End If



    'rowcount = 1
    gridOne.Rows = 1

    gconDMIS.Execute ("delete from ALL_AOR WHERE LogName=" & N2Str2Null(LOGNAME))

    For termno = 1 To Term
        MonthlyInterest = MyRound(Interest * Principal)
        Principal = MyRound(Principal - (payment - MonthlyInterest))

        If termno = Term And Principal <> 0 Then
            payment = payment + Principal
            Principal = 0
        End If

        'interest_Total = MyRound(interest_Total + MonthlyInterest)
        'payment_Total = payment_Total + payment

        'ytdInterest = ytdInterest + MonthlyInterest
        'ytdApplied = ytdApplied + (payment - MonthlyInterest)



        gridOne.AddItem termno & Chr(9) & FormatCurrency(payment) & Chr(9) & FormatCurrency(MonthlyInterest) & Chr(9) & FormatCurrency((payment - MonthlyInterest)) & Chr(9) & FormatCurrency(Principal)

        gconDMIS.Execute ("Insert Into ALL_AOR (LogName,termno, payment,principal, monthlyinterest ,balance) " _
                        & " values(" & N2Str2Null(LOGNAME) & "," & termno & "," & payment & "," & Principal & "," & MonthlyInterest & " ," & payment - MonthlyInterest & " )")




    Next
    txtTotalInterest = FormatCurrency(payment_Total)
    txtTotalPrincipal = FormatCurrency(interest_Total)
    Command1.Enabled = True

End Sub







Public Function MyRound(payment As Currency)

    payment = payment + 0.005
    payment = payment * 100
    payment = Int(payment)
    payment = payment / 100
    MyRound = payment
End Function


Private Sub cmdCancelView2_Click()
    ShowHidePictureBox picViewVehicles.hwnd, False, Me
End Sub

Private Sub cmdSelectViewVehicles_Click()
    If cboFinCom.ListIndex = -1 Then Exit Sub

    If lvViewVehicles.SelectedRows.Count = 0 Then
        txtFinComp.Text = cboFinCom.Text
        ShowHidePictureBox picViewVehicles.hwnd, False, Me
        Exit Sub
    End If


    With lvViewVehicles.SelectedRows.Row(0)
        txtFinComp.Text = cboFinCom.Text
        cboTerms.Text = Replace(.Record(0).Value, "Months", "")
        txtAOR.Text = .Record(1).Value
    End With
    ShowHidePictureBox picViewVehicles.hwnd, False, Me
End Sub

Private Sub cmdViewSelect_Click()
    ShowHidePictureBox picViewVehicles.hwnd, Not (picViewVehicles.Visible), Me
    If picViewVehicles.Visible Then
        cboFinCom.SetFocus
    Else

    End If
End Sub

Private Sub Command1_Click()
    With rptAOR
        .Formulas(0) = "VehicleName='" & cboVDetails & "'"
        .Formulas(1) = "finco='" & txtFinComp & "'"
        .Formulas(2) = "netsaleprice='" & txtNetSalesAmount & "'"
        .Formulas(3) = "downpayment ='" & txtDownPayment & "'"
        .Formulas(4) = "downpaymentpercent='" & (NumericVal(txtDownPayment) / NumericVal(txtNetSalesAmount) * 100) & "%'"
        .Formulas(5) = "baltofin='" & txtBaltofin & "'"
        .Formulas(6) = "AOR ='" & txtAOR.Text & "%'"
        .Formulas(7) = "NoTerms = '" & cboTerms.Text & " Months '"
    End With
    PrintSQLReport rptAOR, SMIS_REPORT_PATH & "AOR.rpt", "{ALL_AOR.LogName}='" & LOGNAME & "'", DMIS_REPORT_Connection, 1


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    FillCombo "select ID, COMPANY FROM SMIS_FINCOM", 0, 1, cboFinCom
    FillCombo "select descript from ALL_MODEL", -1, 0, cboVDetails
    Call ReportControlAddColumnHeader(lvViewVehicles, "TERM, PERCENTAGE")
    ReportControlPaintManager lvViewVehicles
    SetComboMaxLength cboTerms, 4
End Sub

Private Sub Form_Unload(Cancel As Integer)

    gconDMIS.Execute ("delete from ALL_AOR WHERE LogName=" & N2Str2Null(LOGNAME))

End Sub

Private Sub lvViewVehicles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdSelectViewVehicles_Click
    End If
End Sub

Private Sub lvViewVehicles_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdSelectViewVehicles_Click
End Sub

Private Sub optRular_Click()
    cboFinCom_Click
End Sub

Private Sub optUrban_Click()
    cboFinCom_Click
End Sub


Private Sub txtAOR_Change()
    If NumericVal(cboTerms) = 0 Then: Exit Sub
    Command1.Enabled = False
    '   On Error GoTo adder:
    Dim Principal, Term, Interest, payment

    Principal = NumericVal(txtDownPayment.Text)

    Term = Int(NumericVal(cboTerms.Text))
    Interest = Val(txtAOR.Text)

    If Principal = 0 Then
        txtNetMonthlyAmorization.Text = 0
    Else
        Interest = Interest / 1200
        payment = (Principal * Interest / (1 - ((1 / (1 + Interest) ^ Term))))
        txtNetMonthlyAmorization.Text = FormatCurrency(payment)
    End If
    '   Exit Sub
    'adder:
    '   Err.Clear


End Sub

Private Sub txtAOR_Validate(Cancel As Boolean)
    If IsNumeric(txtAOR) = False Then: Cancel = True
End Sub

Private Sub txtDownPayment_change()
    Command1.Enabled = False
    txtBaltofin = FormatCurrency(NumericVal(txtNetSalesAmount) - NumericVal(txtDownPayment), 2, vbTrue, vbTrue, vbTrue)
    txtAOR_Change
End Sub

Private Sub txtDownPayment_Validate(Cancel As Boolean)
    If IsNumeric(NumericVal(txtDownPayment)) = False Then
        Cancel = True
    End If
    txtDownPayment = FormatCurrency(NumericVal(txtDownPayment))
End Sub

Private Sub txtNetSalesAmount_Change()
    Command1.Enabled = False

    txtBaltofin = FormatCurrency(NumericVal(txtNetSalesAmount) - NumericVal(txtDownPayment), 2, vbTrue, vbTrue, vbTrue)
    txtAOR_Change
End Sub

Private Sub txtNetSalesAmount_Validate(Cancel As Boolean)
    If IsNumeric(NumericVal(txtNetSalesAmount)) = False Then
        Cancel = True
    End If
    txtNetSalesAmount = FormatCurrency(NumericVal(txtNetSalesAmount))
End Sub

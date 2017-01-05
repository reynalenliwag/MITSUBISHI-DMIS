VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPMISAC_ARIFormation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accessories Requistion Number Formation"
   ClientHeight    =   5400
   ClientLeft      =   1620
   ClientTop       =   5880
   ClientWidth     =   4410
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lbl9 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   3300
      MaxLength       =   3
      TabIndex        =   19
      Text            =   "123"
      Top             =   660
      Width           =   525
   End
   Begin MSComCtl2.DTPicker dtTranDate 
      Height          =   375
      Left            =   1890
      TabIndex        =   0
      Top             =   90
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   58261507
      CurrentDate     =   38957
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tracking code of parts && Acc."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   60
      TabIndex        =   14
      Top             =   6900
      Width           =   4155
      Begin VB.CommandButton Command1 
         Caption         =   "Clear"
         Height          =   300
         Left            =   2850
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   630
         Width           =   1215
      End
      Begin VB.OptionButton opt0 
         Caption         =   "0 = last issuance"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   690
         Width           =   3255
      End
      Begin VB.OptionButton opt2 
         Caption         =   "2 =second issuance"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   3285
      End
      Begin VB.OptionButton opt1 
         Caption         =   "1 = first issuance"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Width           =   3885
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Payment Classification"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1005
      Left            =   150
      TabIndex        =   10
      Top             =   3540
      Width           =   4155
      Begin VB.OptionButton optC 
         Caption         =   "C = Customer paid"
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   3285
      End
      Begin VB.OptionButton optI 
         Caption         =   "I = Internal paid"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   3285
      End
      Begin VB.OptionButton optW2 
         Caption         =   "W = Warranty"
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   690
         Width           =   3405
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Service type/application"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   150
      TabIndex        =   7
      Top             =   2700
      Width           =   4155
      Begin VB.OptionButton optG 
         Caption         =   "G = General job"
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   3735
      End
      Begin VB.OptionButton optB 
         Caption         =   "B =Body and paint"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Origin Request"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   135
      TabIndex        =   1
      Top             =   1215
      Width           =   4155
      Begin VB.OptionButton optO 
         Caption         =   "O =Others"
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   1170
         Width           =   3885
      End
      Begin VB.OptionButton optJ 
         Caption         =   "J = Jobber"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   930
         Width           =   1845
      End
      Begin VB.OptionButton optM 
         Caption         =   "M = Sales and marketing department"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   690
         Width           =   4005
      End
      Begin VB.OptionButton optW 
         Caption         =   "W = Walk-in/over the counter"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3885
      End
      Begin VB.OptionButton optS 
         Caption         =   "S = Service department"
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   3885
      End
   End
   Begin VB.TextBox txtedit 
      Height          =   360
      Left            =   330
      TabIndex        =   37
      Top             =   6810
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3660
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Cancel"
      Top             =   4680
      Width           =   675
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Save Selected Options"
      Top             =   4680
      Width           =   675
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   525
      Left            =   450
      Top             =   600
      Width           =   3645
   End
   Begin VB.Shape Shape1 
      Height          =   585
      Left            =   420
      Top             =   570
      Width           =   3705
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   5040
      X2              =   5310
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   3330
      X2              =   3810
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   2940
      X2              =   3210
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   2520
      X2              =   2790
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2100
      X2              =   2370
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1650
      X2              =   1920
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1200
      X2              =   1470
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   750
      X2              =   1020
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date"
      Height          =   315
      Left            =   120
      TabIndex        =   36
      Top             =   150
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      Height          =   285
      Index           =   10
      Left            =   3480
      TabIndex        =   35
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label lbl11 
      Alignment       =   2  'Center
      Caption         =   "PI "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   34
      Top             =   630
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   285
      Index           =   9
      Left            =   3120
      TabIndex        =   33
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   285
      Index           =   8
      Left            =   2970
      TabIndex        =   32
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   285
      Index           =   7
      Left            =   2640
      TabIndex        =   31
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label lbl8 
      Alignment       =   2  'Center
      Caption         =   "PI "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2850
      TabIndex        =   30
      Top             =   660
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " 67"
      Height          =   285
      Index           =   5
      Left            =   2190
      TabIndex        =   29
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label lbl6_7 
      Alignment       =   2  'Center
      Caption         =   "PI "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   28
      Top             =   660
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   27
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   285
      Index           =   3
      Left            =   1350
      TabIndex        =   26
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      Caption         =   "PI "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2010
      TabIndex        =   25
      Top             =   660
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   24
      Top             =   1230
      Width           =   465
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      Caption         =   "PI "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   23
      Top             =   660
      Width           =   495
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      Caption         =   "PI "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   22
      Top             =   660
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      Height          =   285
      Index           =   0
      Left            =   390
      TabIndex        =   21
      Top             =   1230
      Width           =   465
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "AR "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   660
      TabIndex        =   20
      Top             =   660
      Width           =   495
   End
End
Attribute VB_Name = "frmPMISAC_ARIFormation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PISNUMBER                                          As String

Sub GetSeries()
    With frmPMISAC_ARISForms

        If txtedit = "EDIT" Then
            If lbl2.Caption = Mid(.txtReferencePIS, 3, 1) And lbl3.Caption = Mid(.txtReferencePIS, 4, 1) And lbl4.Caption = Mid(.txtReferencePIS, 5, 1) Then
                lbl9.Text = Mid(.txtReferencePIS, 9, 3)
                Exit Sub
            End If
        End If
    End With

    Dim RSORD_HD                                       As ADODB.Recordset
    Set RSORD_HD = New ADODB.Recordset
    RSORD_HD.Open "select is_series from PMIS_vw_PRS where TYPE = 'A' AND sales_origin = '" & lbl2.Caption & "' and SI_TYPE ='" & lbl3.Caption & "' and  PAY_CLASS = '" & lbl4.Caption & "'  order by is_series desc ", gconDMIS
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        lbl9.Text = Format(NumericVal(RSORD_HD![IS_SERIES] + 1), "000")
    Else
        lbl9.Text = "001"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If lbl2.Caption = "" Or lbl3.Caption = "" Or lbl4.Caption = "" Then
        MsgBox "ARS number not completed!"
        Exit Sub
    End If
    PISNUMBER = LTrim(lbl1.Caption) & LTrim(lbl2.Caption) & LTrim(lbl3.Caption) & LTrim(lbl4.Caption) & LTrim(lbl6_7.Caption) & LTrim(lbl8.Caption) & Format(NumericVal(LTrim(lbl9)), "000") & LTrim(lbl11.Caption)

    Dim RSORD_HD                                       As ADODB.Recordset
    Set RSORD_HD = New ADODB.Recordset
    RSORD_HD.Open "select refpisno,TRANNO from PMIS_vw_PRS where TYPE = 'A' AND refpisno = '" & PISNUMBER & "' and status='P'", gconDMIS
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        MsgBox "ARS number already exist in Transaction number : " & Null2String(RSORD_HD!TRANNO)
        Exit Sub
    End If

    With frmPMISAC_ARISForms
        .txtReferencePIS = PISNUMBER
        .txtTranDate = Format(dtTranDate, "MM/dd/yyyy")
        If Mid(PISNUMBER, 5, 1) = "I" Then
            .txtCustCode.Text = "D00005"
            .SetCustomer
        End If
    End With
    cmdCancel.Value = True
End Sub

Private Sub Command1_Click()
    opt0.Value = False
    opt1.Value = False
    opt2.Value = False
    lbl11.Caption = ""
End Sub

Private Sub dtTranDate_Change()
    lbl1.Caption = "AR"
    lbl6_7.Caption = Format(dtTranDate, "yy")
    If Format(dtTranDate, "MM") = "01" Then
        lbl8.Caption = "A"
    ElseIf Format(dtTranDate, "MM") = "02" Then
        lbl8.Caption = "B"
    ElseIf Format(dtTranDate, "MM") = "03" Then
        lbl8.Caption = "C"
    ElseIf Format(dtTranDate, "MM") = "04" Then
        lbl8.Caption = "D"
    ElseIf Format(dtTranDate, "MM") = "05" Then
        lbl8.Caption = "E"
    ElseIf Format(dtTranDate, "MM") = "06" Then
        lbl8.Caption = "F"
    ElseIf Format(dtTranDate, "MM") = "07" Then
        lbl8.Caption = "G"
    ElseIf Format(dtTranDate, "MM") = "08" Then
        lbl8.Caption = "H"
    ElseIf Format(dtTranDate, "MM") = "09" Then
        lbl8.Caption = "I"
    ElseIf Format(dtTranDate, "MM") = "10" Then
        lbl8.Caption = "J"
    ElseIf Format(dtTranDate, "MM") = "11" Then
        lbl8.Caption = "K"
    ElseIf Format(dtTranDate, "MM") = "12" Then
        lbl8.Caption = "L"
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    optS.Value = False
    optW.Value = False
    optM.Value = False
    optJ.Value = False
    optO.Value = False

    optG.Value = False
    optB.Value = False

    optC.Value = False
    optI.Value = False
    optW2.Value = False

    opt1.Value = False
    opt2.Value = False
    opt0.Value = False

    dtTranDate.Value = Format(Now, "MM/dd/yyyy")
    lbl1.Caption = "AR"
    lbl2.Caption = ""
    lbl3.Caption = ""
    lbl4.Caption = ""
    lbl6_7.Caption = Format(dtTranDate, "yy")
    If Format(dtTranDate, "MM") = "01" Then
        lbl8.Caption = "A"
    ElseIf Format(dtTranDate, "MM") = "02" Then
        lbl8.Caption = "B"
    ElseIf Format(dtTranDate, "MM") = "03" Then
        lbl8.Caption = "C"
    ElseIf Format(dtTranDate, "MM") = "04" Then
        lbl8.Caption = "D"
    ElseIf Format(dtTranDate, "MM") = "05" Then
        lbl8.Caption = "E"
    ElseIf Format(dtTranDate, "MM") = "06" Then
        lbl8.Caption = "F"
    ElseIf Format(dtTranDate, "MM") = "07" Then
        lbl8.Caption = "G"
    ElseIf Format(dtTranDate, "MM") = "08" Then
        lbl8.Caption = "H"
    ElseIf Format(dtTranDate, "MM") = "09" Then
        lbl8.Caption = "I"
    ElseIf Format(dtTranDate, "MM") = "10" Then
        lbl8.Caption = "J"
    ElseIf Format(dtTranDate, "MM") = "11" Then
        lbl8.Caption = "K"
    ElseIf Format(dtTranDate, "MM") = "12" Then
        lbl8.Caption = "L"
    End If
    lbl9 = ""
    lbl11.Caption = ""
End Sub

Private Sub optS_Click()
    If optS.Value = True Then
        lbl2.Caption = "S"
        If lbl2.Caption <> "" And lbl3.Caption <> "" And lbl4.Caption <> "" Then
            GetSeries
        End If
    End If
End Sub

Private Sub optW_Click()
    If optW.Value = True Then
        lbl2.Caption = "W"
        If lbl2.Caption <> "" And lbl3.Caption <> "" And lbl4.Caption <> "" Then
            GetSeries
        End If
    End If
End Sub

Private Sub optM_Click()
    If optM.Value = True Then
        lbl2.Caption = "M"
        If lbl2.Caption <> "" And lbl3.Caption <> "" And lbl4.Caption <> "" Then
            GetSeries
        End If
    End If
End Sub

Private Sub optJ_Click()
    If optJ.Value = True Then
        lbl2.Caption = "J"
        If lbl2.Caption <> "" And lbl3.Caption <> "" And lbl4.Caption <> "" Then
            GetSeries
        End If
    End If
End Sub

Private Sub optO_Click()
    If optO.Value = True Then
        lbl2.Caption = "O"
        If lbl2.Caption <> "" And lbl3.Caption <> "" And lbl4.Caption <> "" Then
            GetSeries
        End If
    End If
End Sub

Private Sub optG_Click()
    If optG.Value = True Then
        lbl3.Caption = "G"
        If lbl2.Caption <> "" And lbl3.Caption <> "" And lbl4.Caption <> "" Then
            GetSeries
        End If
    End If
End Sub

Private Sub optB_Click()
    If optB.Value = True Then
        lbl3.Caption = "B"
        If lbl2.Caption <> "" And lbl3.Caption <> "" And lbl4.Caption <> "" Then
            GetSeries
        End If
    End If
End Sub

Private Sub optC_Click()
    If optC.Value = True Then
        lbl4.Caption = "C"
        If lbl2.Caption <> "" And lbl3.Caption <> "" And lbl4.Caption <> "" Then
            GetSeries
        End If
    End If
End Sub

Private Sub optI_Click()
    If optI.Value = True Then
        lbl4.Caption = "I"
        If lbl2.Caption <> "" And lbl3.Caption <> "" And lbl4.Caption <> "" Then
            GetSeries
        End If
    End If
End Sub

Private Sub optW2_Click()
    If optW2.Value = True Then
        lbl4.Caption = "W"
        If lbl2.Caption <> "" And lbl3.Caption <> "" And lbl4.Caption <> "" Then
            GetSeries
        End If
    End If
End Sub

Private Sub opt1_Click()
    If opt1.Value = True Then lbl11.Caption = "1"
End Sub

Private Sub opt2_Click()
    If opt2.Value = True Then lbl11.Caption = "2"
End Sub

Private Sub opt0_Click()
    If opt0.Value = True Then lbl11.Caption = "0"
End Sub


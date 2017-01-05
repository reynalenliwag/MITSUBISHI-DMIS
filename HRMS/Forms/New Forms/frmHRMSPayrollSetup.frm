VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmHRMSPayrollSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll date Setup"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "frmHRMSPayrollSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   5895
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   2190
      ScaleHeight     =   945
      ScaleWidth      =   1740
      TabIndex        =   0
      Top             =   3180
      Width           =   1740
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
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
         Left            =   930
         MouseIcon       =   "frmHRMSPayrollSetup.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPayrollSetup.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Left            =   240
         MouseIcon       =   "frmHRMSPayrollSetup.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPayrollSetup.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   2310
      ScaleHeight     =   945
      ScaleWidth      =   1590
      TabIndex        =   25
      Top             =   3180
      Width           =   1590
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
         Height          =   795
         Left            =   810
         MouseIcon       =   "frmHRMSPayrollSetup.frx":0EF0
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPayrollSetup.frx":1042
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   120
         MouseIcon       =   "frmHRMSPayrollSetup.frx":1380
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPayrollSetup.frx":14D2
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox fmepay 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   30
      ScaleHeight     =   3105
      ScaleWidth      =   5805
      TabIndex        =   3
      Top             =   30
      Width           =   5835
      Begin VB.TextBox txtDesc 
         Height          =   360
         Left            =   1470
         TabIndex        =   11
         Top             =   810
         Width           =   4215
      End
      Begin VB.TextBox txtFrom1 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1470
         Width           =   1245
      End
      Begin VB.TextBox txtTo1 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1920
         Width           =   1245
      End
      Begin VB.TextBox txtTo2 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   4410
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1920
         Width           =   1245
      End
      Begin VB.TextBox txtFrom2 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   4410
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1470
         Width           =   1245
      End
      Begin VB.ComboBox cboCutOff 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2670
         Width           =   1725
      End
      Begin VB.ComboBox cboYear 
         Height          =   360
         Left            =   4050
         TabIndex        =   5
         Text            =   "cboYear"
         Top             =   2670
         Width           =   1605
      End
      Begin VB.ComboBox cboMonth 
         Height          =   360
         Left            =   2010
         TabIndex        =   4
         Text            =   "cboMonth"
         Top             =   2670
         Width           =   1995
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   -30
         TabIndex        =   28
         Top             =   0
         Width           =   5865
         _Version        =   655364
         _ExtentX        =   10345
         _ExtentY        =   609
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   900
         TabIndex        =   24
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   330
         TabIndex        =   23
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   900
         TabIndex        =   22
         Top             =   1590
         Width           =   495
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1155
         TabIndex        =   21
         Top             =   2010
         Width           =   240
      End
      Begin VB.Label lblCode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1470
         TabIndex        =   20
         Top             =   450
         Width           =   1245
      End
      Begin VB.Label lblUSE 
         BackColor       =   &H000000FF&
         Caption         =   "Label1"
         Height          =   315
         Left            =   4050
         TabIndex        =   19
         Top             =   420
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   4125
         TabIndex        =   18
         Top             =   2010
         Width           =   240
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   3870
         TabIndex        =   17
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "2nd Cut-Off"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   7
         Left            =   4410
         TabIndex        =   16
         Top             =   1230
         Width           =   960
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "1st Cut-Off"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   6
         Left            =   1470
         TabIndex        =   15
         Top             =   1230
         Width           =   915
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Period Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   4230
         TabIndex        =   14
         Top             =   2400
         Width           =   990
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Period Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   2055
         TabIndex        =   13
         Top             =   2400
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cutt off Period"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   330
         TabIndex        =   12
         Top             =   2400
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmHRMSPayrollSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADD_EDIT                                                          As String

Public Function LimitChar(ByVal alpha As String, ByVal k As Integer)
    If InStr(alpha, Chr(k)) > 0 Or k = 8 Then
        LimitChar = k
    Else
        LimitChar = 0
    End If
End Function

Sub FillDetails()
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_PAYROLLSETUP WHERE CODE = 1")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        lblCode.Caption = Null2String(RSTMP!CODE)
        txtDesc.Text = Null2String(RSTMP!Description)
        txtFrom1.Text = Null2String(RSTMP!FROMDATE1)
        txtTo1.Text = Null2String(RSTMP!TODATE1)
        txtFrom2.Text = Null2String(RSTMP!FROMDATE2)
        txtTo2.Text = Null2String(RSTMP!TODATE2)

        If Null2String(RSTMP!NOTEDBY2) = "1" Then
            cboCutOff.Text = "1st Cut-Off"
        ElseIf Null2String(RSTMP!NOTEDBY2) = "2" Then
            cboCutOff.Text = "2nd Cut-Off"
        Else
            MsgBox "Cut-Off not set!"
        End If

        If NumericVal(RSTMP!PERIODMONTH) > 0 And NumericVal(RSTMP!PERIODMONTH) <= 12 Then
            cboMOnth = MonthName(NumericVal(RSTMP!PERIODMONTH))
        Else
            MsgBox "Month not set!"
            cboMOnth = ""
        End If
    End If
    Set RSTMP = Nothing
End Sub

Private Sub cmdCancel_Click()
    Picture2.Visible = False
    Picture3.Visible = True
    fmepay.Enabled = False
End Sub

Private Sub cmdEdit_Click()
    'If Function_Access(LOGID, "ACESS_EDIT", "WORKING DAY SETUP") = False Then Exit Sub
    ADD_EDIT = "EDIT"
    Picture2.Visible = True
    Picture3.Visible = False
    fmepay.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    'If Function_Access(LOGID, "ACESS_PRINT", "WORKING DAY SETUP") = False Then Exit Sub
End Sub

Private Sub cmdSave_Click()

    Dim RSTMP                                                         As New ADODB.Recordset
    Dim vDESC                                                         As String
    Dim vFROM1                                                        As String
    Dim vFROM2                                                        As String
    Dim vTO1                                                          As String
    Dim vTO2                                                          As String
    Dim VCUT_OFF                                                      As String
    Dim vMONTH                                                        As String
    Dim vYEAR                                                         As String

    If txtDesc.Text = "" Then
        MessagePop InfoOk, "Required Field", "Payroll Description Missing"
        txtDesc.SetFocus
        Exit Sub
    End If

    If txtFrom1.Text = "" Or txtTo1.Text = "" Then
        If txtFrom1.Text = "" Then
            MessagePop InfoOk, "Required Field", "From Day Cannot be Empty"
            txtFrom1.SetFocus
        End If
        If txtTo1.Text = "" Then
            MessagePop InfoOk, "Required Field", "To Day Cannot be Empty"
            txtTo1.SetFocus
        End If
        Exit Sub
    End If

    If txtFrom2.Text = "" Or txtTo2.Text = "" Then
        If txtFrom2.Text = "" Then
            MessagePop InfoOk, "Required Field", "From Day Cannot be Empty"
            txtFrom2.SetFocus
        End If
        If txtTo2.Text = "" Then
            MessagePop InfoOk, "Required Field", "To Day Cannot be Empty"
            txtTo2.SetFocus
        End If
        Exit Sub
    End If

    If cboCutOff.Text = "" Then
        MsgBox "Cutoff Period is Required Field. Please Select Cutoff Period From List", vbInformation
        cboCutOff.SetFocus
        Exit Sub
    End If

    If cboMOnth.Text = "" Then
        MsgBox "Payroll Period (Month) is Required Field. Please Select Period Month List", vbInformation
        cboMOnth.SetFocus
        Exit Sub
    End If

    If cboyear.Text = "" Then
        MsgBox "Payroll Period (Year) is Required Field. Please Select Period Year List", vbInformation
        cboyear.SetFocus
        Exit Sub
    End If

    If cboCutOff.Text = "1st Cut-Off" Then
        CUTTOFF_CODE = 1
        frmMain.StatusBar1.Panels(10).Text = "1st Cut-off " & "-" & PAY_MONTH & " " & PAY_YEAR
    Else
        CUTTOFF_CODE = 2
        frmMain.StatusBar1.Panels(10).Text = "2nd Cut-off " & "-" & PAY_MONTH & " " & PAY_YEAR
    End If

    PAY_MONTH = What_month(cboMOnth)
    PAY_YEAR = cboyear

    vDESC = N2Str2Null(txtDesc.Text)
    vFROM1 = N2Str2Null(txtFrom1.Text)
    vFROM2 = N2Str2Null(txtFrom2.Text)
    vTO1 = N2Str2Null(txtTo1.Text)
    vTO2 = N2Str2Null(txtTo2.Text)
    VCUT_OFF = N2Str2Null(CUTTOFF_CODE)
    vMONTH = N2Str2Null(What_month(cboMOnth))
    vYEAR = N2Str2Null(cboyear)

    PAYROLLCODE_FROM1 = txtFrom1.Text
    PAYROLLCODE_FROM2 = txtFrom2.Text
    PAYROLLCODE_TO1 = txtTo1.Text
    PAYROLLCODE_TO2 = txtTo2.Text

    SQL_STATEMENT = "UPDATE HRMS_PAYROLLSETUP SET DESCRIPTION = " & vDESC & _
                    ", FromDate1 = " & vFROM1 & _
                    ", ToDate1 = " & vTO1 & _
                    ", FromDate2 = " & vFROM2 & _
                    ", ToDate2 = " & vTO2 & _
                    ", NOTEDBY2 = " & VCUT_OFF & _
                    ", PERIODMONTH = " & vMONTH & _
                    ", PERIODYEAR= " & vYEAR & _
                  " WHERE CODE = '" & lblCode.Caption & "'"

    gconDMIS.Execute SQL_STATEMENT
    NEW_LogAudit "E", "WORKING DAY SETUP", SQL_STATEMENT, lblCode.Caption, "", "", "", ""
    SQL_STATEMENT = ""
    GetThePayrollCode

    MessagePop InfoOk, "Record Saved", "HRMS Payroll Setting Sucessfully Updated", 1000, 0
    Dim FRM                                                           As Form
    For Each FRM In Forms
        If Not (UCase(FRM.NAME) = UCase("frmMain") Or UCase(FRM.NAME) = UCase("frmMainMenu") Or UCase(FRM.NAME) = UCase(Me.NAME)) Then
            Unload FRM
        End If
    Next
    Unload Me
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (WORKING DAY SETUP)"
            Call frmALL_AuditInquiry.DisplayHistory(lblCode.Caption, "WORKING DAY SETUP")
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    With cboCutOff
        .AddItem "1st Cut-Off"
        .AddItem "2nd Cut-Off"
    End With
    fillcbomonth cboMOnth
    'FillcboYear cboyear
    fillcombo_up cboyear
    FillDetails
    DrawXPCtl Me
End Sub

Private Sub txtFrom1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789", KeyAscii)
    End If
End Sub

Private Sub txtFrom2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789", KeyAscii)
    End If
End Sub

Private Sub txtTo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789", KeyAscii)
    End If
End Sub

Private Sub txtTo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("0123456789", KeyAscii)
    End If
End Sub


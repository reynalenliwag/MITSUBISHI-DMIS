VERSION 5.00
Begin VB.Form frmHRMSGenerate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Payroll"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3870
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Generate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   3870
   Begin VB.CheckBox chkConfidential 
      Caption         =   "Process Confidential Employees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   1260
      Width           =   3675
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      MouseIcon       =   "Generate.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "Generate.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancel"
      Top             =   4230
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      MouseIcon       =   "Generate.frx":079A
      MousePointer    =   99  'Custom
      Picture         =   "Generate.frx":08EC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancel"
      Top             =   4260
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CheckBox chkContractual 
      Caption         =   "Process for Contractual Employees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   1740
      Width           =   3675
   End
   Begin VB.CheckBox chkAllowanceBase 
      Caption         =   "Process for Allowance Base Employees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   2010
      Width           =   3675
   End
   Begin VB.CheckBox chkProbReg 
      Caption         =   "Process for Probationary/Regular Employees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   1500
      Width           =   3675
   End
   Begin VB.ComboBox cboYear 
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
      Left            =   2460
      TabIndex        =   2
      Text            =   "cboYear"
      Top             =   570
      Width           =   1035
   End
   Begin VB.ComboBox cboMonth 
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
      Left            =   570
      TabIndex        =   1
      Text            =   "cboMonth"
      Top             =   570
      Width           =   1845
   End
   Begin VB.ComboBox cboQuensina 
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
      Left            =   570
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   2925
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
      Height          =   825
      Left            =   1890
      MouseIcon       =   "Generate.frx":0C2A
      MousePointer    =   99  'Custom
      Picture         =   "Generate.frx":0D7C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancel"
      Top             =   2400
      Width           =   885
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1020
      MouseIcon       =   "Generate.frx":10BA
      MousePointer    =   99  'Custom
      Picture         =   "Generate.frx":120C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Process Generation of Payroll"
      Top             =   2400
      Width           =   885
   End
End
Attribute VB_Name = "frmHRMSGenerate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
: Option Explicit
Dim rsEmpInfo, rsDeductions, rsPAYROLL, rsLoanmasDet                  As ADODB.Recordset
Attribute rsDeductions.VB_VarUserMemId = 1073938432
Attribute rsPAYROLL.VB_VarUserMemId = 1073938432
Attribute rsLoanmasDet.VB_VarUserMemId = 1073938432
Dim FromDate, ToDate                                                  As String
Attribute FromDate.VB_VarUserMemId = 1073938436
Attribute ToDate.VB_VarUserMemId = 1073938436

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    gconDMIS.Execute ("Delete From HRMS_Payroll")

    gconDMIS.Execute ("Delete From HRMS_SSS")
    gconDMIS.Execute ("Delete From HRMS_sssdet")

    gconDMIS.Execute ("Delete From HRMS_ATM")
    gconDMIS.Execute ("Delete From HRMS_ATMDet")

    gconDMIS.Execute ("Delete From HRMS_TinDet")
    gconDMIS.Execute ("Delete From HRMS_Tin")

    gconDMIS.Execute ("Delete From HRMS_PhilHealth")
    gconDMIS.Execute ("Delete From HRMS_PhilHealthdet")

    gconDMIS.Execute ("Delete From HRMS_PagIbig")
    gconDMIS.Execute ("Delete From HRMS_PagIbigdet")

    gconDMIS.Execute ("Delete From HRMS_LoanMas")
    gconDMIS.Execute ("Delete From HRMS_LoanMasDet")
End Sub

Private Sub cmdGenerate_Click()
     
    Dim nix As String
    Dim YEAR As String
     
    
    
    If Function_Access(LOGID, "Acess_Process", "PROCESS GENERATE PAYROLL") = False Then Exit Sub

    PROCESS_OPTION = ""
    Dim matt                                                          As Integer
    If cboQuensina.Text = "1st Cut-Off" Then
        matt = 1
    Else
        matt = 2
    End If

    If (cboQuensina.Text) = "1st Cut-Off" Then
        GENFROM = DateSerial(cboYear.Text, What_month(cboMonth.Text), PAYROLLCODE_FROM1)
        GENTO = DateSerial(cboYear.Text, What_month(cboMonth.Text), PAYROLLCODE_TO1)
        If PAYROLLCODE_FROM1 > PAYROLLCODE_TO1 Then
            GENFROM = DateSerial(NumericVal(cboYear.Text), What_month(cboMonth.Text) - 1, PAYROLLCODE_FROM1)
            If What_month(cboMonth.Text) = 1 Then
                GENFROM = DateSerial(NumericVal(cboYear.Text) - 1, 12, PAYROLLCODE_FROM1)
            End If
        End If
    Else
        GENFROM = DateSerial(cboYear.Text, What_month(cboMonth.Text), PAYROLLCODE_FROM2)
        GENTO = DateSerial(cboYear.Text, What_month(cboMonth.Text), PAYROLLCODE_TO2)
    End If

    If chkConfidential.Value = 1 Then
        PROCESS_OPTION = " EMPLEVEL = 'M'"
    End If
    If PROCESS_OPTION <> "" Then
        If chkProbReg.Value = 1 Then
            PROCESS_OPTION = PROCESS_OPTION & " OR EMPLEVEL = 'E'"
        End If
    Else
        PROCESS_OPTION = " EMPLEVEL = 'E'"
    End If
    If PROCESS_OPTION <> "" Then
        If chkContractual.Value = 1 Then
            PROCESS_OPTION = PROCESS_OPTION & " OR EMPLEVEL = 'C'"
        End If
    Else
        PROCESS_OPTION = " EMPLEVEL = 'C'"
    End If
    If PROCESS_OPTION <> "" Then
        If chkAllowanceBase.Value = 1 Then
            PROCESS_OPTION = PROCESS_OPTION & " OR EMPLEVEL = 'A'"
        End If
    Else
        PROCESS_OPTION = " EMPLEVEL = 'A'"
    End If
    If chkAllowanceBase.Value = 0 And chkConfidential.Value = 0 And chkContractual.Value = 0 And chkProbReg.Value = 0 Then
        MsgBox "Please select an option to process...", vbInformation, "No Option to Process"
        Exit Sub
    End If
    PROCESS_OPTION = "(" & PROCESS_OPTION & ")"
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT * FROM HRMS_PAYROLL WHERE " & PROCESS_OPTION & " AND CUT_OFF = '" & matt & "' AND PAY_MONTH = " & What_month(cboMonth.Text) & " AND PAY_YEAR = " & cboYear.Text & "")
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        If MsgQuestionBox("Payroll for " & cboQuensina.Text & " " & cboMonth.Text & " " & cboYear.Text & " has already been generated and posted." & vbCrLf & _
                          "Do You want to Clear Generated Payroll? ", "Generate Payroll") = True Then
            If UCase(LOGLEVEL) <> "ADM" And Null2String(rsPAYROLL!payrollstatus) = "P" Then
                MsgBox "You do not have sufficient level to clear the posted payroll!"
                Exit Sub
            End If
            gconDMIS.Execute "DELETE FROM HRMS_SSSDET WHERE " & PROCESS_OPTION & " AND CUT_OFF = '" & matt & "' AND PAY_MONTH = " & What_month(cboMonth.Text) & " AND PAY_YEAR = " & cboYear.Text & ""
            gconDMIS.Execute "DELETE FROM HRMS_PHILHEALTHDET WHERE " & PROCESS_OPTION & " AND CUT_OFF = '" & matt & "' AND PAY_MONTH = " & What_month(cboMonth.Text) & " AND PAY_YEAR = " & cboYear.Text & ""
            gconDMIS.Execute "DELETE FROM HRMS_PAGIBIGDET WHERE " & PROCESS_OPTION & " AND CUT_OFF = '" & matt & "' AND PAY_MONTH = " & What_month(cboMonth.Text) & " AND PAY_YEAR = " & cboYear.Text & ""
            gconDMIS.Execute "DELETE FROM HRMS_TINDET WHERE " & PROCESS_OPTION & " AND CUT_OFF = '" & matt & "' AND PAY_MONTH = " & What_month(cboMonth.Text) & " AND PAY_YEAR = " & cboYear.Text & ""
            gconDMIS.Execute "DELETE FROM HRMS_ATMDET WHERE " & PROCESS_OPTION & " AND CUT_OFF = '" & matt & "' AND PAY_MONTH = " & What_month(cboMonth.Text) & " AND PAY_YEAR = " & cboYear.Text & ""
            gconDMIS.Execute "DELETE FROM HRMS_PAYROLL WHERE " & PROCESS_OPTION & " AND CUT_OFF = '" & matt & "' AND PAY_MONTH = " & What_month(cboMonth.Text) & " AND PAY_YEAR = " & cboYear.Text & ""
            
            
'            If matt = 1 Then
'                    matt = 2
'            ElseIf matt = 2 Then
'                    matt = 1
'            End If
            
            nix = What_month(cboMonth.Text)
            YEAR = cboYear.Text

            If matt = 1 Then
                matt = 2
                If nix = 1 Then
                   nix = 12
                   YEAR = YEAR - 1
                Else
                   nix = nix - 1
                End If
            
            ElseIf matt = 2 Then
                matt = 1
            End If
            
            'orig
            'gconDMIS.Execute "DELETE FROM HRMS_PAYROLL_DET WHERE " & PROCESS_OPTION & " AND CUT_OFF = '" & matt & "' AND PAY_MONTH = " & What_month(cboMonth.Text) & " AND PAY_YEAR = " & cboYear.Text & ""
            
            gconDMIS.Execute "DELETE FROM HRMS_PAYROLL_DET WHERE " & PROCESS_OPTION & " AND CUT_OFF = '" & matt & "' AND PAY_MONTH = " & nix & " AND PAY_YEAR = " & YEAR & ""
            
            
            OVERWRAYT = False
            MsgBox "Payroll for " & cboQuensina.Text & " " & cboMonth.Text & " " & cboYear.Text & " has been cleared.", vbInformation, "HRMS Payroll Generation"
        End If
    Else
        OVERWRAYT = True
        frmHRMSProgress.labCutOff.Caption = cboQuensina.Text
        frmHRMSProgress.Show vbModal
    End If
    Exit Sub
Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'DrawXPCtl Me
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim rsCutoff                                                      As ADODB.Recordset
    Set rsCutoff = New ADODB.Recordset
    Set rsCutoff = gconDMIS.Execute("SELECT PERIODMONTH,PERIODYEAR,NOTEDBY2 FROM HRMS_PAYROLLSETUP")
    If Not (rsCutoff.EOF And rsCutoff.BOF) Then
        If NumericVal(rsCutoff!NOTEDBY2) = 1 Then
            cboQuensina.Clear
            cboQuensina.AddItem "1st Cut-Off"
            cboQuensina.Text = "1st Cut-Off"
        ElseIf NumericVal(rsCutoff!NOTEDBY2) = 2 Then
            cboQuensina.Clear
            cboQuensina.AddItem "2nd Cut-Off"
            cboQuensina.Text = "2nd Cut-Off"
        Else
            MsgBox "Cut-off not set"
        End If
        cboMonth.Clear
        cboMonth.AddItem MonthName(Null2String(rsCutoff!PERIODMONTH))
        cboMonth.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
        cboYear.Clear
        cboYear.AddItem Null2String(rsCutoff!PERIODYEAR)
        cboYear.Text = Null2String(rsCutoff!PERIODYEAR)
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub


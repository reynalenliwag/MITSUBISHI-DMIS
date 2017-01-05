VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSVehicleByModel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle By Model"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "VehicleByModel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   5445
   Begin VB.ComboBox cboModel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   2010
      TabIndex        =   2
      Text            =   "cboModel"
      Top             =   90
      Width           =   3345
   End
   Begin Crystal.CrystalReport rptByModel 
      Left            =   120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Vehicle By Model"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   765
      Left            =   4590
      MouseIcon       =   "VehicleByModel.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "VehicleByModel.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   540
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   765
      Left            =   3870
      MouseIcon       =   "VehicleByModel.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "VehicleByModel.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Vehicle Model"
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
      Height          =   225
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   1875
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2100
      TabIndex        =   0
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "frmCSMSVehicleByModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPOR                                            As ADODB.Recordset

Sub FillCbo()
    Dim rsS_Model                                      As ADODB.Recordset
    Set rsS_Model = New ADODB.Recordset
    rsS_Model.Open "Select DISTINCT upper(model) Model from CSMS_CUSVEH where isnull(model,'')<>'' ", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboModel.Clear
        Do While Not rsS_Model.EOF
            cboModel.AddItem Null2String(rsS_Model!MODEL)
            rsS_Model.MoveNext
        Loop
    End If
    Set rsS_Model = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "VEHICLE BY MODEL") = False Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo Errorcode
    Dim CTR                                            As Integer

    Set rsREPOR = New ADODB.Recordset
    Set rsREPOR = gconDMIS.Execute("Select COUNT(*)AS CTR from CSMS_CusVeh Where MODEL = '" & cboModel.Text & "'")
    If Not rsREPOR.EOF And Not rsREPOR.BOF Then
        CTR = rsREPOR!CTR

        rptByModel.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
        rptByModel.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
        rptByModel.Formulas(2) = "Printedby ='" & LOGNAME & "'"
        rptByModel.Formulas(3) = "ModelTotal ='" & CTR & "'"
        rptByModel.Formulas(4) = "FilterModel ='" & cboModel.Text & "'"

        rptByModel.ReportTitle = "VEHICLE BY MODEL (" & UCase(cboModel.Text) & ")"
        PrintSQLReport rptByModel, CSMS_REPORT_PATH & "VehicleByModel.rpt", "UCASE({cusveh.model}) = '" & UCase(cboModel.Text) & "'", CSMS_REPORT_CONNECTION, 1    'JUN 02/12/2008

        'LogAudit "V", "VEHICLE BY MODEL - REPORTS ", cboModel
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "VEHICLE BY MODEL", "", "", "", "MODEL : " & cboModel, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    Else
        Screen.MousePointer = 0
        ShowNoRecord
        Exit Sub
    End If
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE BY MODEL)"
            Call frmALL_AuditInquiry.DisplayHistory("", "VEHICLE BY MODEL", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    FillCbo
    cboModel.ListIndex = 0
    Screen.MousePointer = 0
End Sub

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSPartsPickList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parts Pick-List"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3765
   Icon            =   "frmCSMSPartsPickList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   3765
   Begin VB.ComboBox cboValue 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   1290
      TabIndex        =   2
      Top             =   600
      Width           =   2310
   End
   Begin VB.OptionButton optFromEstimate 
      Caption         =   "From Estimate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1860
      TabIndex        =   1
      Top             =   120
      Width           =   1755
   End
   Begin VB.OptionButton optFromAppointment 
      Caption         =   "Job Order"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1395
   End
   Begin Crystal.CrystalReport rptParts_Pick_List 
      Left            =   690
      Top             =   1410
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Parts Pick-List Report"
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      MouseIcon       =   "frmCSMSPartsPickList.frx":27A2
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSPartsPickList.frx":28F4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   1110
      Width           =   705
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2190
      MouseIcon       =   "frmCSMSPartsPickList.frx":2D3F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSPartsPickList.frx":2E91
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   1110
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Job Order"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   5
      Top             =   690
      Width           =   915
   End
End
Attribute VB_Name = "frmCSMSPartsPickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "PARTS PICK LIST") = False Then Exit Sub

    On Error GoTo ErrorCode


    If optFromEstimate.Value = True Then

        Dim rsParts_Pick_ListE                         As ADODB.Recordset
        Set rsParts_Pick_ListE = New ADODB.Recordset


        Set rsParts_Pick_ListE = gconDMIS.Execute("Select * from CSMS_EstDetails Where ESTIMATENO = '" & cboValue.Text & "'")


        If Not rsParts_Pick_ListE.EOF And Not rsParts_Pick_ListE.BOF Then
            Screen.MousePointer = 11

            'JUN 01/05/2008
            rptParts_Pick_List.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptParts_Pick_List.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptParts_Pick_List.Formulas(2) = "Printedby = '" & LOGNAME & "'"

            PrintSQLReport rptParts_Pick_List, CSMS_REPORT_PATH & "Parts_Pick_List_FromEstimate.rpt", "{CSMS_EstDetails.ESTIMATENO} = '" & cboValue.Text & "'", CSMS_REPORT_CONNECTION, 1

            LogAudit "V", "PARTS PICK LIST FROM ESTIMATE  - REPORTS ", cboValue
            Screen.MousePointer = 0
        Else
            ShowNoRecord
            On Error Resume Next
            cboValue.SetFocus
            Exit Sub
        End If
        Exit Sub
    Else

        Dim rsParts_Pick_ListA                         As ADODB.Recordset
        Set rsParts_Pick_ListA = New ADODB.Recordset
        Set rsParts_Pick_ListA = gconDMIS.Execute("Select * from CSMS_Ro_Det Where Rep_or = '" & cboValue.Text & "' AND TRANSTYPE = '" & "R" & "'")
        If Not rsParts_Pick_ListA.EOF And Not rsParts_Pick_ListA.BOF Then
            Screen.MousePointer = 11

            'JUN 01/05/2008
            rptParts_Pick_List.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptParts_Pick_List.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptParts_Pick_List.Formulas(2) = "Printedby = '" & LOGNAME & "'"

            PrintSQLReport rptParts_Pick_List, CSMS_REPORT_PATH & "Parts_Pick_List_FromAppointment.rpt", "{CSMS_Repor.Rep_or} = '" & cboValue.Text & "' AND {CSMS_repor.TRANSTYPE} = 'R' and {CSMS_RO_DET.livil}='2'", CSMS_REPORT_CONNECTION, 1

            LogAudit "V", "PARTS PICK LIST FROM APPOINTMENT  - REPORTS ", cboValue
            Screen.MousePointer = 0
        Else
            ShowNoRecord
            cboValue.SetFocus
            Exit Sub
        End If

        Exit Sub

    End If
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cboEstimateNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        cboValue.SetFocus
    End If
End Sub

Private Sub cboValue_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
    '    Call cmdPrint_Click
    'End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    optFromAppointment.Value = True
End Sub

Private Sub optFromAppointment_Click()
    cboValue.Clear
    Dim tmp_valueA                                     As String
    tmp_valueA = ""
    Dim rsAppointment_Number                           As ADODB.Recordset
    Set rsAppointment_Number = New ADODB.Recordset
    cboValue.Clear
    rsAppointment_Number.Open "Select REP_OR from CSMS_REPOR where TRANSTYPE = 'R' ORDER BY REP_OR", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsAppointment_Number.EOF And Not rsAppointment_Number.BOF Then
        rsAppointment_Number.MoveFirst
        Do While Not rsAppointment_Number.EOF
            'If tmp_valueA = Null2String(rsAppointment_Number!ApptNo) Then
            '    rsAppointment_Number.MoveNext
            'Else

            '    If Null2String(rsAppointment_Number!ApptNo) = "" Then
            '    Else
            cboValue.AddItem Null2String(rsAppointment_Number!rep_OR)
            '        tmp_valueA = Null2String(rsAppointment_Number!ApptNo)
            '    End If
            rsAppointment_Number.MoveNext
            'End If
        Loop
    End If
    Set rsAppointment_Number = Nothing
End Sub

Private Sub optFromEstimate_Click()
    cboValue.Clear
    Dim tmp_valueE                                     As String
    tmp_valueE = ""
    Dim rsEstimate_Number                              As ADODB.Recordset
    Set rsEstimate_Number = New ADODB.Recordset
    cboValue.Clear
    rsEstimate_Number.Open "Select ESTIMATENO from CSMS_EstHD order by ESTIMATENO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsEstimate_Number.EOF And Not rsEstimate_Number.BOF Then
        rsEstimate_Number.MoveFirst

        Do While Not rsEstimate_Number.EOF
            'If tmp_valueE = Null2String(rsEstimate_Number!EstimateNo) Then
            '    rsEstimate_Number.MoveNext
            'Else
            cboValue.AddItem Null2String(rsEstimate_Number!EstimateNo)
            '    tmp_valueE = Null2String(rsEstimate_Number!EstimateNo)
            rsEstimate_Number.MoveNext
            'End If
        Loop
    End If
    Set rsEstimate_Number = Nothing
End Sub


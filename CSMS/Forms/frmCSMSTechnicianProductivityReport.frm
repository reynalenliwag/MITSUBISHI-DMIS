VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmCSMSTechnicianProductivityReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Technician Productivity Report"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
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
      Height          =   825
      Left            =   1140
      MouseIcon       =   "frmCSMSTechnicianProductivityReport.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianProductivityReport.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   870
      Width           =   885
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
      Height          =   825
      Left            =   2130
      MouseIcon       =   "frmCSMSTechnicianProductivityReport.frx":05F1
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianProductivityReport.frx":0743
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   870
      Width           =   885
   End
   Begin Crystal.CrystalReport rptWork_In_Progress 
      Left            =   405
      Top             =   1125
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Caption         =   "Technician Productivity"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   390
      Width           =   2490
   End
End
Attribute VB_Name = "frmCSMSTechnicianProductivityReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''==============================================================
''Function/Feature: Techinician Productivity Report
''Date Started: 05/17/2007 3:00pm
''Last Update:
''Database Updates:
''Who Updated: Jonathan
''Updating Code: JAA - 05172007
''==============================================================
'
'Private Sub cboEstimateNumber_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        'Call cmdPrint_Click
'        cboEstimateNumber.SetFocus
'    End If
'End Sub
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    MoveKeyPress KeyCode
'End Sub
'
'Private Sub Form_Load()
'    Screen.MousePointer = 11
'    CenterMe frmMain, Me, 1
'    Screen.MousePointer = 0
'    FillCombo
'End Sub
'Sub cmdPrint_Click()
'    On Error GoTo ErrorCode
'    If Function_Access(LOGID, "Acess_Print") = False Then Exit Sub
'
'    If optFromEstimate.Value = True Then
'        'Print Parts Pick-List from Estimate
'        Dim rsParts_Pick_List          As ADODB.Recordset
'        Set rsParts_Pick_List = New ADODB.Recordset
'        Set rsParts_Pick_List = gconDMIS.Execute("Select * from CSMS_EstDetails Where ESTIMATENO = '" & cboValue.Text & "'")
'        If Not rsParts_Pick_List.EOF And Not rsParts_Pick_List.BOF Then
'            Screen.MousePointer = 11
'            PrintSQLReport rptParts_Pick_List, CSMS_REPORT_PATH & "Parts_Pick_List.rpt", "{CSMS_EstDetails.ESTIMATENO} = '" & cboValue.Text & "'", CSMS_REPORT_Connection, 1
'            Screen.MousePointer = 0
'        Else
'            ShowNoRecord
'            cboValue.SetFocus
'            Exit Sub
'        End If
'        Exit Sub
'    Else
'        'Print Parts Pick-List from Appointment
'        ShowNoRecord
'        cboValue.SetFocus
'        Exit Sub
'    End If
'ErrorCode:
'    ShowVBError
'    Screen.MousePointer = 0
'End Sub
'
'Sub FillCombo()
'    Dim tmp_valueA                     As String
'    tmp_valueA = ""
'    Dim rsAppointment_Number           As ADODB.Recordset
'    Set rsAppointment_Number = New ADODB.Recordset
'    rsAppointment_Number.Open "Select ApptNo from CSMS_Appointment order by ApptNo asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
'    If Not rsAppointment_Number.EOF And Not rsAppointment_Number.BOF Then
'        rsAppointment_Number.MoveFirst
'        cboValue.Clear
'        Do While Not rsAppointment_Number.EOF
'            If tmp_valueA = rsAppointment_Number!ApptNo Then
'                rsAppointment_Number.MoveNext
'            Else
'                cboValue.AddItem Null2String(rsAppointment_Number!ApptNo)
'                tmp_valueA = rsAppointment_Number!ApptNo
'                rsAppointment_Number.MoveNext
'            End If
'        Loop
'    End If
'    Set rsAppointment_Number = Nothing
'
'
'    Dim tmp_valueE                     As String
'    tmp_valueE = ""
'    Dim rsEstimate_Number              As ADODB.Recordset
'    Set rsEstimate_Number = New ADODB.Recordset
'    rsEstimate_Number.Open "Select ESTIMATENO from CSMS_EstDetails order by ESTIMATENO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
'    If Not rsEstimate_Number.EOF And Not rsEstimate_Number.BOF Then
'        rsEstimate_Number.MoveFirst
'        'cboValue.Clear
'        Do While Not rsEstimate_Number.EOF
'            If tmp_valueE = rsEstimate_Number!EstimateNo Then
'                rsEstimate_Number.MoveNext
'            Else
'                cboValue.AddItem Null2String(rsEstimate_Number!EstimateNo)
'                tmp_valueE = rsEstimate_Number!EstimateNo
'                rsEstimate_Number.MoveNext
'            End If
'        Loop
'    End If
'    Set rsEstimate_Number = Nothing
'End Sub
'

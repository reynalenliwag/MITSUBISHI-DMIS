VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_Report_ServiceAppointment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Appointment"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Report_ServiceAppointment.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport rptServiceAppointment 
      Left            =   465
      Top             =   825
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Service Appointment"
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpServiceAppointment 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd/MMMM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   390
      Left            =   2010
      TabIndex        =   0
      Top             =   105
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      Format          =   52887553
      CurrentDate     =   31392
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
      Left            =   2100
      MouseIcon       =   "Report_ServiceAppointment.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "Report_ServiceAppointment.frx":0E1C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   585
      Width           =   915
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
      Left            =   1200
      MouseIcon       =   "Report_ServiceAppointment.frx":1267
      MousePointer    =   99  'Custom
      Picture         =   "Report_ServiceAppointment.frx":13B9
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   585
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Appointment Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   405
      TabIndex        =   3
      Top             =   165
      Width           =   1950
   End
End
Attribute VB_Name = "frmCRIS_Report_ServiceAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub cmdPrint_Click()

    On Error GoTo ErrorCode

    Screen.MousePointer = 11
    Dim ApptDate                                                      As Date
    Dim rsServiceAppointment                                          As ADODB.Recordset
    Set rsServiceAppointment = New ADODB.Recordset

    ApptDate = CDate(dtpServiceAppointment.Value)

    Set rsServiceAppointment = gconDMIS.Execute("Select * from CSMS_vw_Appointment_Diary")
    If Not rsServiceAppointment.EOF And Not rsServiceAppointment.BOF Then
        Screen.MousePointer = 11
        rptServiceAppointment.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptServiceAppointment.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptServiceAppointment, CRIS_REPORT_PATH & "ServiceAppointment.rpt", "{CSMS_vw_Appointment_Diary.TranDate} =  date(" & Year(ApptDate) & ", " & Month(ApptDate) & ", " & Day(ApptDate) & ")", CRIS_REPORT_PATH, 1
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "SERVICE APPOINTMENT", "", "", "", dtpServiceAppointment, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

        'LogAudit "V", "SERVICE APPOINTMENT", DateValue(dtpServiceAppointment.Value)
        Screen.MousePointer = 0
    Else
        ShowNoRecord
    End If
    Screen.MousePointer = 0

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub dtpServiceAppointment_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdPrint_Click
    End If
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SERVICE APPOINTMENT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SERVICE APPOINTMENT", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    dtpServiceAppointment.Value = LOGDATE
End Sub


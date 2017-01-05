VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCSMSAppointmentDiary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Appointment Diary Report"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSAppointmentDiary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   4740
   Begin Crystal.CrystalReport rptAppointment_Diary 
      Left            =   90
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Appointment Diary Report"
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpAppointmentDiaryReport 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd/MMMM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   90
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   661
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
      Format          =   51380225
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
      Height          =   795
      Left            =   3900
      MouseIcon       =   "frmCSMSAppointmentDiary.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSAppointmentDiary.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   540
      Width           =   735
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
      Height          =   795
      Left            =   3180
      MouseIcon       =   "frmCSMSAppointmentDiary.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSAppointmentDiary.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter Appointment Date"
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
      Left            =   90
      TabIndex        =   3
      Top             =   150
      Width           =   2265
   End
End
Attribute VB_Name = "frmCSMSAppointmentDiary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub dtpAppointmentDiaryReport_KeyPress(KeyAscii As Integer)
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (APPOINTMENT DIARY)"
            Call frmALL_AuditInquiry.DisplayHistory("", "APPOINTMENT DIARY", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
    dtpAppointmentDiaryReport.Value = Now
End Sub

Sub cmdPrint_Click()
    Screen.MousePointer = 11
    On Error GoTo ERRORCODE
    Dim ApptDate                                       As Date
    Dim rsAppointment_Diary                            As ADODB.Recordset

    ApptDate = CDate(dtpAppointmentDiaryReport.Value)
    Set rsAppointment_Diary = New ADODB.Recordset
    Set rsAppointment_Diary = gconDMIS.Execute("Select * from CSMS_vw_Appointment_Diary")
    If Not rsAppointment_Diary.EOF And Not rsAppointment_Diary.BOF Then
        Screen.MousePointer = 11

        'JUN 01/05/2008
        rptAppointment_Diary.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptAppointment_Diary.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptAppointment_Diary.Formulas(2) = "Printedby = '" & LOGNAME & "'"

        PrintSQLReport rptAppointment_Diary, CSMS_REPORT_PATH & "Appointmentdiary.rpt", "{CSMS_Appointment.TranDate} =  date(" & Year(ApptDate) & ", " & Month(ApptDate) & ", " & Day(ApptDate) & ")", CSMS_REPORT_CONNECTION, 1

        'LogAudit "V", "APPOINTMENT DIARY - REPORTS ", dtpAppointmentDiaryReport
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "APPOINTMENT DIARY", "", "", "", dtpAppointmentDiaryReport.Value, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        Screen.MousePointer = 0
    Else
        ShowNoRecord
        Exit Sub
    End If
    Screen.MousePointer = 0
    Exit Sub

ERRORCODE:
    ShowVBError
    Screen.MousePointer = 0
End Sub


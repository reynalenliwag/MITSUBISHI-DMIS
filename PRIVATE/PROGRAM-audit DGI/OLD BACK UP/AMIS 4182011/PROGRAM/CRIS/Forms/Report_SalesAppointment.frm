VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_Report_SalesAppointment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Appointment"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   4005
   Icon            =   "Report_SalesAppointment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport rptSalesAppointment 
      Left            =   570
      Top             =   675
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Sales Appointment"
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpAppointmentDate 
      Height          =   390
      Left            =   1890
      TabIndex        =   0
      Top             =   105
      Width           =   2010
      _ExtentX        =   3545
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
      Format          =   52625409
      CurrentDate     =   39203
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
      Left            =   2145
      MouseIcon       =   "Report_SalesAppointment.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "Report_SalesAppointment.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   630
      Width           =   885
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
      Height          =   825
      Left            =   1275
      MouseIcon       =   "Report_SalesAppointment.frx":0B27
      MousePointer    =   99  'Custom
      Picture         =   "Report_SalesAppointment.frx":0C79
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   630
      Width           =   885
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
      Left            =   255
      TabIndex        =   3
      Top             =   165
      Width           =   1875
   End
End
Attribute VB_Name = "frmCRIS_Report_SalesAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:
    Screen.MousePointer = 11
    Dim SA_Date                                                       As Date
    Dim Date_Of_Appointment                                           As Date
    Dim rsSalesApppointment                                           As ADODB.Recordset
    Set rsSalesApppointment = New ADODB.Recordset
    Dim found                                                         As Integer

    SA_Date = CDate(dtpAppointmentDate.Value)
    rptSalesAppointment.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptSalesAppointment.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptSalesAppointment.Formulas(2) = "Appointment_Date = '" & SA_Date & "'"
    PrintSQLReport rptSalesAppointment, CRIS_REPORT_PATH & "SalesApointments.rpt", "Date({CRIS_SalesAppointments.StartDateTime}) = date(" & Year(SA_Date) & "," & Month(SA_Date) & "," & Day(SA_Date) & ")", CRIS_REPORT_PATH, 1
    'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
     Call NEW_LogAudit("V", "SALES APPOINTMENT", "", "", "", dtpAppointmentDate, "", "")
    'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'LogAudit "V", "SALES APPOINTMENT", SA_Date
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdPrint_Click
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES APPOINTMENT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SALES APPOINTMENT", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    dtpAppointmentDate.Value = LOGDATE
End Sub


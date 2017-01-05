VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_Report_CustomerRemindersAndTask 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Reminders And Tasks"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   4425
   Icon            =   "Report_CustomerReminderandTask.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1515
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport rptCustomerReminders 
      Left            =   510
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Customer Reminders And Tasks"
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpReminderDate 
      Height          =   390
      Left            =   1650
      TabIndex        =   0
      Top             =   90
      Width           =   2580
      _ExtentX        =   4551
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
      Format          =   52690945
      CurrentDate     =   39203
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
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
      Left            =   3435
      MouseIcon       =   "Report_CustomerReminderandTask.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "Report_CustomerReminderandTask.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   555
      Width           =   795
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
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
      Left            =   2655
      MouseIcon       =   "Report_CustomerReminderandTask.frx":0B27
      MousePointer    =   99  'Custom
      Picture         =   "Report_CustomerReminderandTask.frx":0C79
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   555
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Reminder Date:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   150
      Width           =   1635
   End
End
Attribute VB_Name = "frmCRIS_Report_CustomerRemindersAndTask"
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

    '
    '
    Dim AppointmentDate                                               As Date
    Dim rsReminder                                                    As ADODB.Recordset
    Dim d_date                                                        As Date
    Dim found                                                         As Integer
    Dim OverDue                                                       As String
    AppointmentDate = CDate(dtpReminderDate.Value)

    Set rsReminder = New ADODB.Recordset
    Set rsReminder = gconDMIS.Execute("Select * from CRIS_Reminders")

    If Not rsReminder.BOF And Not rsReminder.EOF Then
        rsReminder.MoveFirst
        Do While Not rsReminder.EOF
            d_date = Format(Null2String(rsReminder!DateTimeRemind), "mm/dd/yyyy")
            If d_date = AppointmentDate Then
                If DateDiff("d", d_date, Date) <= 0 Then
                    OverDue = "0 Days"
                Else
                    OverDue = DateDiff("d", d_date, Date) & " Days"
                End If
                rptCustomerReminders.Formulas(0) = "OverDueBy = '" & OverDue & "'"
                rptCustomerReminders.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
                rptCustomerReminders.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptCustomerReminders, CRIS_REPORT_PATH & "CustomerReminderAndTask.rpt", "Date({CRIS_Reminders.DateTimeRemind}) = date(" & Year(AppointmentDate) & "," & Month(AppointmentDate) & "," & Day(AppointmentDate) & ")", CRIS_REPORT_PATH, 1
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 Call NEW_LogAudit("V", "CUSTOMER REMINDERS AND TASKS", "", "", "", dtpReminderDate, "", "")
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

                'LogAudit "V", "CUSTOMER REMINDERS AND TASKS", AppointmentDate
                found = 1                                    'found
                Exit Do
            End If
            rsReminder.MoveNext
        Loop
    Else
        ShowNoRecord
    End If

    If found = 1 Then

    Else
        ShowNoRecord
    End If

    'End of update
    Screen.MousePointer = 0

    Exit Sub
ErrorCode:
    ShowVBError
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CUSTOMER REMINDERS AND TASKS)"
            Call frmALL_AuditInquiry.DisplayHistory("", "CUSTOMER REMINDERS AND TASKS", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    dtpReminderDate.Value = LOGDATE
End Sub


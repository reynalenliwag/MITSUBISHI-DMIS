VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReportUserAcess 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Right Access"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   Icon            =   "ReportUserAcess.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Print Blank Report Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1350
      TabIndex        =   10
      Top             =   2670
      Width           =   2205
   End
   Begin VB.CheckBox chkDataEntry 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1350
      TabIndex        =   9
      Top             =   1320
      Width           =   2205
   End
   Begin VB.CheckBox chkTransaction 
      Caption         =   "Transaction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1350
      TabIndex        =   8
      Top             =   1575
      Width           =   2205
   End
   Begin VB.CheckBox chkProcessing 
      Caption         =   "Processing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1350
      TabIndex        =   7
      Top             =   1845
      Width           =   2205
   End
   Begin VB.CheckBox chkInquiry 
      Caption         =   "Inquiry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1350
      TabIndex        =   6
      Top             =   2100
      Width           =   2205
   End
   Begin VB.CheckBox chkReport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1350
      TabIndex        =   5
      Top             =   2400
      Width           =   2205
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1290
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   960
      Width           =   2235
   End
   Begin Crystal.CrystalReport rptInternalReminder 
      Left            =   510
      Top             =   2025
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Internal Reminders"
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
      Height          =   825
      Left            =   2115
      MouseIcon       =   "ReportUserAcess.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportUserAcess.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3030
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
      Left            =   1245
      MouseIcon       =   "ReportUserAcess.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportUserAcess.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3030
      Width           =   885
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Right Access Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   2370
   End
   Begin VB.Image Image2 
      Height          =   885
      Left            =   -3180
      Picture         =   "ReportUserAcess.frx":19D0
      Top             =   0
      Width           =   7665
   End
   Begin VB.Label Label1 
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1005
      Width           =   1635
   End
End
Attribute VB_Name = "frmReportUserAcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Function Feature   : Reminder Module
'Date               : 06/26/2007
'Last Update        : 06/26/2007
'Database Update    : Added Table For Reminder Called Cris Reminders
'Who Updated        : AXP
'Upating Code       : AXP-070720071152

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode
    If chkDataEntry.Value = 0 And chkTransaction.Value = 0 And chkReport.Value = 0 And chkProcessing.Value = 0 And chkInquiry.Value = 0 Then
        MessagePop RecSaveError, "Selection Required", "Select At Least One Selection"
        Screen.MousePointer = 0
        Exit Sub
    End If
    'Updating Code: AXP-12072007 for blank report printing
    Screen.MousePointer = 11
    If Check1.Value = 1 Then
        rptInternalReminder.Formulas(0) = "COMPANYNAME='" & Company_name & "'"
        rptInternalReminder.Formulas(1) = "COMPANYADDRESS='" & Company_Address & "'"

        If chkDataEntry.Value = 1 Then
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportDataEntry.rpt", "{A.MainModuleName}='AMIS' and {A.Module_Type}='Data Entry'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportDataEntry.rpt", "{A.MainModuleName}='PMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportDataEntry.rpt", "{A.MainModuleName}='CMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportDataEntry.rpt", "{A.MainModuleName}='CSMS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportDataEntry.rpt", "{A.MainModuleName}='SMIS' and {A.Module_Type}='Data Entry'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportDataEntry.rpt", "{A.MainModuleName}='HRMS'", DMIS_REPORT_Connection, 1
        End If

        If chkTransaction.Value = 1 Then
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportTransaction.rpt", "{A.MainModuleName}='AMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportTransaction.rpt", "{A.MainModuleName}='PMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportTransaction.rpt", "{A.MainModuleName}='CMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportTransaction.rpt", "{A.MainModuleName}='CSMS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportTransaction.rpt", "{A.MainModuleName}='SMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportTransaction.rpt", "{A.MainModuleName}='HRMS'", DMIS_REPORT_Connection, 1

        End If
        If chkReport.Value = 1 Then
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportReport.rpt", "{A.MainModuleName}='AMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportReport.rpt", "{A.MainModuleName}='PMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportReport.rpt", "{A.MainModuleName}='CMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportReport.rpt", "{A.MainModuleName}='CSMS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportReport.rpt", "{A.MainModuleName}='SMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportReport.rpt", "{A.MainModuleName}='HRMS'", DMIS_REPORT_Connection, 1
        End If
        If chkProcessing.Value = 1 Then

            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportProcessing.rpt", "{A.MainModuleName}='AMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportProcessing.rpt", "{A.MainModuleName}='PMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportProcessing.rpt", "{A.MainModuleName}='CMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportProcessing.rpt", "{A.MainModuleName}='CSMS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportProcessing.rpt", "{A.MainModuleName}='SMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportProcessing.rpt", "{A.MainModuleName}='HRMS'", DMIS_REPORT_Connection, 1
        End If
        If chkInquiry.Value = 1 Then
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportInquiry.rpt", "{A.MainModuleName}='AMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportInquiry.rpt", "{A.MainModuleName}='PMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportInquiry.rpt", "{A.MainModuleName}='CMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportInquiry.rpt", "{A.MainModuleName}='CSMS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportInquiry.rpt", "{A.MainModuleName}='SMIS'", DMIS_REPORT_Connection, 1
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\BLANK\AccessReportInquiry.rpt", "{A.MainModuleName}='HRMS'", DMIS_REPORT_Connection, 1
        End If
        Screen.MousePointer = 0
    Else
        'Updating Code: AXP-070720071152


        Dim TEMPRS                      As ADODB.Recordset
        Dim UID
        Dim FILTER
        Screen.MousePointer = 11

        If Combo1.ListIndex = -1 Then
            MsgBox " Please Select User Name", vbInformation
            Screen.MousePointer = 0
            Exit Sub
        End If

        UID = Combo1.ItemData(Combo1.ListIndex)

        If gconDMIS.Execute("Select COUNT(*) From ALL_RAMS_USERSACESS WHERE USERID=" & UID).Fields(0).Value = 0 Then
            Screen.MousePointer = 0
            ShowNoRecord
            Exit Sub
        End If

        FILTER = " {U.USERID}=" & UID

        rptInternalReminder.Formulas(0) = "COMPANYNAME='" & Company_name & "'"
        rptInternalReminder.Formulas(1) = "COMPANYADDRESS='" & Company_name & "'"

        If chkDataEntry.Value = 1 Then
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\AccessReportDataEntry.rpt", FILTER, DMIS_REPORT_Connection, 1
        End If

        If chkTransaction.Value = 1 Then
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\AccessReportTransaction.rpt", FILTER, DMIS_REPORT_Connection, 1

        End If
        If chkReport.Value = 1 Then

            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\AccessReportReport.rpt", FILTER, DMIS_REPORT_Connection, 1
        End If
        If chkProcessing.Value = 1 Then

            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\AccessReportProcessing.rpt", FILTER, DMIS_REPORT_Connection, 1
        End If
        If chkInquiry.Value = 1 Then
            PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "\AccessReportInquiry.rpt", FILTER, DMIS_REPORT_Connection, 1

        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me: Exit Sub
    MoveKeyPress KeyCode
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
    If CHANGE_USER = True Then
        Call FillCombo("SELECT  USER_NAME,USERID FROM ALL_RAMS_USERS WHERE USERGROUP<>'SDM' order by USER_NAME", 1, 0, Combo1)
    Else
        Call FillCombo("SELECT  USERNAME,USERID FROM ALL_RAMS_USERS WHERE USERGROUP<>'SDM' order by USERNAME", 1, 0, Combo1)
    End If
    If Combo1.ListCount > 0 Then
        Combo1.ListIndex = 0
    End If

End Sub



VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSVehicleAgingOnProcessReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Aging Report"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3690
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSVehicleAgingOnProcessReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1785
   ScaleWidth      =   3690
   Begin Crystal.CrystalReport rptVehicle_Aging_Report 
      Left            =   135
      Top             =   1155
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Vehicle Aging On Process Report"
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpToDateVehicleAging 
      Height          =   345
      Left            =   1335
      TabIndex        =   1
      Top             =   495
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   609
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
      Format          =   20643841
      CurrentDate     =   39232
   End
   Begin MSComCtl2.DTPicker dtpFromDateVehicleAging 
      Height          =   345
      Left            =   1335
      TabIndex        =   0
      Top             =   105
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   609
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
      Format          =   20643841
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
      Height          =   765
      Left            =   2865
      MouseIcon       =   "frmCSMSVehicleAgingOnProcessReport.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSVehicleAgingOnProcessReport.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   945
      Width           =   735
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
      Height          =   765
      Left            =   2145
      MouseIcon       =   "frmCSMSVehicleAgingOnProcessReport.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSVehicleAgingOnProcessReport.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   945
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "From Date"
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
      Left            =   330
      TabIndex        =   5
      Top             =   195
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "To Date"
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
      Left            =   555
      TabIndex        =   4
      Top             =   585
      Width           =   645
   End
End
Attribute VB_Name = "frmCSMSVehicleAgingOnProcessReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    On Error GoTo Errorcode

    Dim rsVehAging                                     As ADODB.Recordset
    Dim FDate                                          As Date
    Dim TDate                                          As Date
    Dim DateAging                                      As Integer

    FDate = CDate(dtpFromDateVehicleAging.Value)
    TDate = CDate(dtpToDateVehicleAging.Value)

    Set rsVehAging = New ADODB.Recordset
    Set rsVehAging = gconDMIS.Execute("SELECT * from CSMS_Repor where status <> '" & "R" & "'")

    If Not rsVehAging.BOF And Not rsVehAging.EOF Then
        rptVehicle_Aging_Report.Formulas(2) = "FromDate = '" & FDate & "'"
        rptVehicle_Aging_Report.Formulas(4) = "ToDate = '" & TDate & "'"

        'JUN 02/05/2008
        rptVehicle_Aging_Report.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptVehicle_Aging_Report.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptVehicle_Aging_Report.Formulas(5) = "Printedby = '" & LOGNAME & "'"

        PrintSQLReport rptVehicle_Aging_Report, CSMS_REPORT_PATH & "Vehicle_Aging_OnProcess_Report.rpt", "{CSMS_Repor.DTE_RECD} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_Repor.DTE_RECD} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", CSMS_REPORT_CONNECTION, 1

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "VEHICLE AGING PROGRESS REPORT", "", "", "", dtpFromDateVehicleAging & " - " & dtpToDateVehicleAging, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    Else
        ShowNoRecord
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE AGING PROGRESS REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "VEHICLE AGING PROGRESS REPORT", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    dtpFromDateVehicleAging.Value = firstDay(LOGDATE)
    dtpToDateVehicleAging.Value = LOGDATE
End Sub


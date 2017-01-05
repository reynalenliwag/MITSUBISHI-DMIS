VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmAMIS_Schedules 
   Caption         =   "Schedule Report"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5025
   Icon            =   "frmAMIS_Schedules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOk 
      Height          =   435
      Left            =   270
      TabIndex        =   1
      Top             =   6510
      Visible         =   0   'False
      Width           =   705
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4995
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmAMIS_Schedules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Dim rptApp                                         As CRAXDRT.Application
    Dim rptRep                                         As Report
    Dim crSections                                     As CRAXDRT.Sections
    Dim crSection                                      As CRAXDRT.Section
    Dim crRepObjs                                      As CRAXDRT.ReportObjects
    Dim crSubRepObj                                    As CRAXDRT.SubreportObject
    Dim crSubReport                                    As CRAXDRT.Report
    Dim j As Integer, k                                As Integer
    Dim ellaine                                        As Integer
    CRViewer1.Width = frmMain.Width
    CRViewer1.Height = frmMain.Height
    Set rptApp = New CRAXDRT.Application
    Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\AP_ScheduleReport.Rpt", 1)
    rptRep.DiscardSavedData
    rptRep.ParameterFields.GetItemByName("CompanyName").AddCurrentValue COMPANY_NAME
    rptRep.ParameterFields.GetItemByName("CompanyAddress").AddCurrentValue COMPANY_ADDRESS
    rptRep.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "ACCOUNTS PAYABLE SCHEDULE REPORT AS OF: " & dtpAsOF
    'Call rptRep.ParameterFields(4).AddCurrentValue(CDate(Now()))

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = rptRep
        .DisplayGroupTree = False
        .DisplayTabs = False
        .DisplayToolbar = True
        .ViewReport
    End With
    Screen.MousePointer = vbDefault
    Set rptApp = Nothing
    Set rptRep = Nothing
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    cmdOk_Click
End Sub

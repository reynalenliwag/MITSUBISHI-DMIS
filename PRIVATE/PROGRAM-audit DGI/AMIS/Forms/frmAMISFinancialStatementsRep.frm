VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmAMISFinancialStatementsRep 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmAMISFinancialStatementsRep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   1635
      Left            =   -30
      TabIndex        =   0
      Top             =   -30
      Width           =   1965
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
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frmAMISFinancialStatementsRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    IncomeStatement
End Sub

Private Sub IncomeStatement()
    Dim rptApp                                              As CRAXDRT.Application
    Dim rptRep                                              As REPORT
    Dim crSections                                          As CRAXDRT.Sections
    Dim crSection                                           As CRAXDRT.Section
    Dim crRepObjs                                           As CRAXDRT.ReportObjects
    Dim crSubRepObj                                         As CRAXDRT.SubreportObject
    Dim crSubReport                                         As CRAXDRT.REPORT
    Dim j As Integer, k                                     As Integer
    Dim ellaine                                             As Integer

    Me.Height = Screen.Height
    Me.Width = Screen.Width
    CRViewer1.Height = Screen.Height - 900
    CRViewer1.Width = Screen.Width
    CRViewer1.ZOrder 0

    Set rptApp = New CRAXDRT.Application
    Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatements2.rpt", 1)
    rptRep.DiscardSavedData
    rptRep.ParameterFields.GetItemByName("CompanyName").AddCurrentValue COMPANY_NAME
    rptRep.ParameterFields.GetItemByName("CompanyAddress").AddCurrentValue COMPANY_ADDRESS
    rptRep.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "INCOME STATEMENTS"
    frmAMISFinancialStatementsRep.Caption = "INCOME STATEMENTS"
    Call rptRep.ParameterFields(4).AddCurrentValue(CDate(frmAMISFinancialStatements.dtpFrom.Value))
    Call rptRep.ParameterFields(5).AddCurrentValue(CDate(frmAMISFinancialStatements.dtpTo.Value))
    Set crSections = rptRep.Sections
    For ellaine = 1 To crSections.Count
        Set crSection = crSections.Item(ellaine)
        Set crRepObjs = crSection.ReportObjects
        For j = 1 To crRepObjs.Count
            If crRepObjs.Item(j).Kind = crSubreportObject Then
                Set crSubReport = rptRep.OpenSubreport(crRepObjs.Item(j).SubreportName)
                Call crSubReport.ParameterFields(1).ClearCurrentValueAndRange
                Call crSubReport.ParameterFields(1).AddCurrentValue(CDate(frmAMISFinancialStatements.dtpFrom.Value))
                Call crSubReport.ParameterFields(2).ClearCurrentValueAndRange
                Call crSubReport.ParameterFields(2).AddCurrentValue(CDate(frmAMISFinancialStatements.dtpTo.Value))
            End If
        Next
    Next
    With CRViewer1
        .ReportSource = rptRep
        .DisplayGroupTree = False
        .DisplayTabs = False
        .DisplayToolbar = True
        .ViewReport
    End With
    Set rptApp = Nothing
    Set rptRep = Nothing
End Sub

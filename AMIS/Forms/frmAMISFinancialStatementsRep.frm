VERSION 5.00
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

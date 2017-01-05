VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#10.4#0"; "CODEJO~2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#10.4#0"; "CO15D0~1.OCX"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Customer Relation Information System"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11715
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0ECA
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5040
      Top             =   2430
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   4230
      Top             =   2460
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin wizEncrypt.wizEnc wizEnc1 
      Left            =   2520
      Top             =   3015
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
   Begin Crystal.CrystalReport rptMain 
      Left            =   2565
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6600
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Object.ToolTipText     =   "Login Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Object.ToolTipText     =   "Login Level"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4410
            MinWidth        =   4410
            Object.ToolTipText     =   "Login Time"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "3:14 PM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Num Lock Status"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock Status"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "SCRL"
            Object.ToolTipText     =   "Scroll Lock Status"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PopupControl PopCntrl 
      Left            =   5640
      Top             =   2460
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   2
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   3435
      Top             =   2460
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   6
      DesignerControls=   "frmMain.frx":13595
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents FormSearch               As frmSMIS_Mis_SearchMaster
Attribute FormSearch.VB_VarHelpID = -1
Dim LOGACTION                           As String
Public Sub ApplyPatches()

End Sub

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case 1000                                            'Customer
            If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub
            frmAllCustomer.Show
            frmAllCustomer.ZOrder 0
        Case 1010                                            'Log Letter
            If Module_Access(LOGID, "CUSTOMERS LOGS", "DATA ENTRY") = False Then Exit Sub
            Call FormSearch.SearchForCustomers
            LOGACTION = "CUS:LETTER"
            FormSearch.Show 1
        Case 1011                                            'Log Visit
            If Module_Access(LOGID, "CUSTOMERS LOGS", "DATA ENTRY") = False Then Exit Sub
            Call FormSearch.SearchForCustomers
            LOGACTION = "CUS:VISIT"
            FormSearch.Show 1
        Case 1012                                            'Log Email
            If Module_Access(LOGID, "CUSTOMERS LOGS", "DATA ENTRY") = False Then Exit Sub
            Call FormSearch.SearchForCustomers
            LOGACTION = "CUS:EMAIL"
            FormSearch.Show 1
        Case 1024                                            'Log Call
            If Module_Access(LOGID, "CUSTOMERS LOGS", "DATA ENTRY") = False Then Exit Sub
            Call FormSearch.SearchForCustomers
            LOGACTION = "CUS:CALL"
            FormSearch.Show 1
        Case 1002                                            'Customer Reminders
            If Module_Access(LOGID, "CUSTOMERS LOGS", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_Log_CustomerReminder.AddReminder ("C")
            frmSMIS_Log_CustomerReminder.Show
            frmSMIS_Log_CustomerReminder.ZOrder 0
        Case 1001                                            'Meetings
        Case 1014                                            'Sales Calculator
            If Module_Access(LOGID, "SALES CALCULATOR", "SYSTEM") = False Then Exit Sub
            frmSMIS_Mis_AOR.ShowonlyComputation
            frmSMIS_Mis_AOR.Show
        Case 1003                                            'Customer Sales History
            If Module_Access(LOGID, "CUSTOMER SALES HISTORY", "INQUIRY") = False Then Exit Sub
            If FormExist("frmSMIS_Inquiry_CustomerSalesHistory") Then
                Unload frmSMIS_Inquiry_CustomerSalesHistory
            End If
            frmSMIS_Inquiry_CustomerSalesHistory.Show
            frmSMIS_Inquiry_CustomerSalesHistory.MyTab.SelectedItem = 0
        Case 1004                                            'Customer Service History
            frmCSMSCustomerHistory.Show
        Case 1005                                            'Customer Transaction History
            If Module_Access(LOGID, "DUPLICATE CUSTOMER", "SYSTEM") = False Then Exit Sub
            frmCRIS_Inquiry_PossibleDuplicates.Show
        Case 1006                                            'Customer Call/Visit History
            If Module_Access(LOGID, "CUSTOMER CALL/VISIT HISTORY", "INQUIRY") = False Then Exit Sub
            Screen.MousePointer = 11
            frmSMIS_Inquiry_CallVisit_History.Show
            Screen.MousePointer = 0
        Case 1007                                            'Customer Vehicle Information Detail
            If Module_Access(LOGID, "CUSTOMER SALES HISTORY", "INQUIRY") = False Then Exit Sub
            Screen.MousePointer = 11
            If FormExist("frmSMIS_Inquiry_CustomerSalesHistory") Then
                Unload frmSMIS_Inquiry_CustomerSalesHistory
            End If
            frmSMIS_Inquiry_CustomerSalesHistory.Show
            frmSMIS_Inquiry_CustomerSalesHistory.MyTab.SelectedItem = 0
            Screen.MousePointer = 0
        Case 1008                                            'CUSTOMER REMINDERS/TASK
            If Module_Access(LOGID, "CUSTOMER REMINDERS/TASK", "DATA ENTRY") = False Then Exit Sub
            Screen.MousePointer = 11
            frmCRIS_Inquiry_TaskList.ShowTaskType ("C")
            frmCRIS_Inquiry_TaskList.Show
            Screen.MousePointer = 0
        Case 1013                                            'CUSTOMER LOG INQUIRY
            If Module_Access(LOGID, "CUSTOMERS LOG INQUIRY", "INQUIRY") = False Then Exit Sub
            On Error GoTo Errorcode:
            Call FormSearch.SearchForProspects(vbNullString)
            LOGACTION = "PROS:LOGINQ"
            FormSearch.Show 1
            Exit Sub
Errorcode:
            ShowVBError
            Call FormSearch.SearchForProspects(vbNullString)
            LOGACTION = "PROS:LOGINQ"
            FormSearch.Show 1
        Case 1016                                            'VEHICLES CATALOGUE/ BROCHURE
            If Module_Access(LOGID, "VEHICLE CATALOGUE", "INQUIRY") = False Then Exit Sub
            Screen.MousePointer = 11
            frmCRIS_InquiryVehicleCatalogue.Show
            Screen.MousePointer = 0
        Case 1017                                            'VEHICLES INVENTORY INFORMATION
            If Module_Access(LOGID, "VEHICLE INVENTORY INFORMATION", "INQUIRY") = False Then Exit Sub
            Screen.MousePointer = 11
            frmSMIS_Inquiry_VehicleMaster.Show
            Screen.MousePointer = 0
        Case 1018                                            'SERVICE APPOINTMENTS
            If Module_Access(LOGID, "SERVICE APPOINTMENT", "INQUIRY") = False Then Exit Sub
            Screen.MousePointer = 11
            'frmCRIS_Inquiry_ServiceAppointment.Show
            frmCSMSAppointment.Show
            Screen.MousePointer = 0
        Case 1019                                            'SALES APPOINTMENTS
            If Module_Access(LOGID, "SALES APPOINTMENT", "INQUIRY") = False Then Exit Sub
            frmCRIS_Inquiry_SalesAppointment.Show
        Case 1020                                            'TEST DRIVE APPOINTMENT
        Case 1023                                            'DASHBOARD
            frmMainMenu.Show
        Case 1022                                            'EXIT SYSTEM
            Unload Me
        Case 1025                                            'CUSTOMER LOG REPORT
            If Module_Access(LOGID, "CUSTOMER LOG REPORT", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_Log.Show
            frmCRIS_Report_Log.optCustomerLog.Value = True
        Case 1026                                            'PROSPECT LOG REPORT
            If Module_Access(LOGID, "PROSPECT LOG REPORT", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_Log.Show
            frmCRIS_Report_Log.optProspectLog.Value = True
        Case 1027                                            'SALES APPOINTMENT
            If Module_Access(LOGID, "SALES APPOINTMENT", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_SalesAppointment.Show
        Case 1028                                            'SERVICE APPOINTMENT
            If Module_Access(LOGID, "SERVICE APPOINTMENT", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_ServiceAppointment.Show
        Case 1029                                            'CUSTOMER REMINDERS && TASKS
            If Module_Access(LOGID, "CUSTOMER REMINDERS/TASK", "REPORST") = False Then Exit Sub
            frmCRIS_Report_CustomerRemindersAndTask.Show
        Case 1030                                            'INTERNAL REMINDERS
            If Module_Access(LOGID, "INTERNAL REMINDERS/TASKS", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_InternalReminder.Show
        Case 1031                                            'CUSTOMER INFORMATION REPOR
            If Module_Access(LOGID, "CUSTOMER INFORMATION REPORT", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_CustomerInformationReport.Show
        Case Else
            'Debug.Print "CASE "; Control.ID; ""; "'"; Control.Caption; vbCrLf;
            'Stop
    End Select
    'Debug.Print "Case "; Control.ID; ""; "'"; Control.Caption; vbCrLf;
End Sub

Private Sub MDIForm_Load()

    ApplyThemes
    ConfigurePopUps
    Set FormSearch = New frmSMIS_Mis_SearchMaster
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Exit CRIS, Are You Sure?", vbExclamation + vbOKCancel, "Exit System") = vbOK Then
        Dim FRM                         As Form
        For Each FRM In Forms
            If Not (FRM Is Nothing) Then
                Unload FRM
            End If
        Next
        CommandBars1.SaveCommandBars MODULENAME, App.TITLE, "Layout"
        UnloadForm Me
    Else
        Cancel = 1
        frmMainMenu.Show
    End If

End Sub

Private Sub PopCntrl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.ID = 707 Then
        PopCntrl.Close
    End If
End Sub
Private Sub ApplyThemes()
    CommandBarsGlobalSettings.App = App
    With CommandBars1
        .EnableOffice2007Frame True
        .SetMDIClient frmMain.hwnd
        .LoadDesignerBars
        .LoadCommandBars MODULENAME, App.TITLE, "Layout"
        .PaintManager.ClearTypeTextQuality = True
        .TabWorkspace.ThemedBackColor = False
        .PaintManager.ThemedStatusBar = True
        .PaintManager.ThemedCheckBox = True
        .PaintManager.ThickCheckMark = True
        .PaintManager.FlatMenuBar = True
        .StatusBar.Visible = True
    End With
    With SkinFramework1
        .LoadSkin "C:\DMIS 2.0\Styles\royale.cjstyle", ""
        
        .ApplyWindow Me.hwnd
        .ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or Not xtpSkinApplyMetrics
    End With
    Dim ToolTipContext                  As ToolTipContext
    Set ToolTipContext = CommandBars1.ToolTipContext
    With ToolTipContext
        .ShowTitleAndDescription True, xtpToolTipIconInfo
        .SetMargin 2, 2, 2, 2
        .MaxTipWidth = 180
        If .IsBalloonStyleSupported Then
            .Style = xtpToolTipBalloon
        Else
            .Style = xtpToolTipOffice2007
        End If
        .ShowShadow = True
    End With
End Sub

Private Sub ConfigurePopUps()
    Dim Item                            As PopupControlItem
    PopCntrl.RemoveAllItems
    'PopCntrl.Icons.AddIcons ImageManager.Icons
    PopCntrl.Icons.AddIcons CommandBars1.Icons
    'PopCntrl.VisualTheme = xtpPopupThemeOffice2003
    'PopCntrl.SetSize 270, 140

    Set Item = PopCntrl.AddItem(245, 8, 265, 20, vbNullString)
    Item.Button = True
    Item.IconIndex = 899
    Item.ID = 707
    Item.Height = 20
    Item.Width = 20
    Item.CenterIcon

    Set Item = PopCntrl.AddItem(10, 10, 218, 30, vbNullString)
    Item.TextColor = RGB(15, 48, 145)
    Item.Bold = True
    Item.Font.Size = 10
    Item.Hyperlink = False
    Set Item = PopCntrl.AddItem(10, 32, 60, 50, vbNullString)
    Item.Height = 50
    Item.Width = 50
    Item.IconIndex = 0
    Item.Hyperlink = False

    Set Item = PopCntrl.AddItem(62, 32, 260, 50, vbNullString)
    Item.TextAlignment = DT_WORDBREAK Or DT_LEFT
    Item.Height = 50
    Item.ID = 655
    Item.Hyperlink = False

    Set Item = PopCntrl.AddItem(20, 85, 260, 105, vbNullString)
    Item.TextAlignment = DT_WORDBREAK Or DT_LEFT
    Item.TextColor = RGB(190, 1, 1)
    Item.Height = 50
    Item.Font.Size = 7
    Item.Hyperlink = False
End Sub

'Private Sub Timer1_Timer()
'    If TIMER_REMIND = "" Then
'        ReminderModule ""
'    Else
'        If DateDiff("n", TIMER_REMIND, Now) >= 0 Then
'
'            frmSMIS_Files_Reminders.Show
'
'        End If
'    End If
'
'End Sub


Private Sub FormSearch_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    Unload FormSearch
    Select Case LOGACTION
        Case "CUS:LETTER"
            Call frmCRIS_Log_Letter.AddLetter(0, oCusRs!CUSCDE)
            frmCRIS_Log_Letter.Show
        Case "CUS:VISIT"
            Call frmCRIS_Log_Visits.AddVisit(0, oCusRs!CUSCDE)
            frmCRIS_Log_Visits.Show
        Case "CUS:CALL"
            Call frmCRIS_Log_Call.AddCall(0, oCusRs!CUSCDE)
            frmCRIS_Log_Call.Show
        Case "CUS:EMAIL"
            Call frmCRIS_Log_Email.AddEmail(0, oCusRs!CUSCDE)
            frmCRIS_Log_Email.Show
        Case "CUS:LOGINQ"
            Call frmSMIS_Inquiry_ViewLog.SHOWCUSTOMERLOG(oCusRs!CUSCDE, oCusRs!AcctName)
            frmSMIS_Inquiry_ViewLog.Show
        Case "PROS:LOGINQ"
            Call frmSMIS_Inquiry_ViewLog.SHOWPROSPECTLOG(oCusRs!ProspectID, oCusRs!AcctName)
            frmSMIS_Inquiry_ViewLog.Show
    End Select

End Sub

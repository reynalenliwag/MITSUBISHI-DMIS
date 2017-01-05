VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#10.4#0"; "CO15D0~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#10.4#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sales Management Information System"
   ClientHeight    =   7110
   ClientLeft      =   1500
   ClientTop       =   3180
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "frmMain.frx":030A
   WindowState     =   2  'Maximized
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   3030
      Top             =   2130
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5100
      Top             =   2130
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6855
      Width           =   15240
      _ExtentX        =   26882
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
            TextSave        =   "2:37 PM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
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
   Begin wizEncrypt.wizEnc wizEnc1 
      Left            =   5760
      Top             =   1185
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
   Begin Crystal.CrystalReport rptMain 
      Left            =   4425
      Top             =   2130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "VEHICLE SALES DIRECTORY"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   3615
      Top             =   2130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1887B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18B95
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18EAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D901
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DC1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DF35
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6418B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A425
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A73F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AA59
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AD73
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B08D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B3A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B6C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B9DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6BCF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C00F
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C329
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C643
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PopupControl PopCntrl 
      Left            =   1980
      Top             =   2130
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
      Width           =   140
      Height          =   270
      AnimateDelay    =   900
      ShowDelay       =   3500
      AllowMove       =   -1  'True
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   2520
      Top             =   2160
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   5
      DesignerControls=   "frmMain.frx":723A5
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ApplyPatches()

End Sub

''''''''''''''START REGION POPUPCONTROLS''''''''''''''''''''''''''''''''''''
Private Sub ConfigurePopUps()
    Dim Item                                                          As PopupControlItem
    PopCntrl.RemoveAllItems
    'PopCntrl.Icons.AddIcons ImageManager.Icons
    PopCntrl.Icons.AddIcons CommandBars1.Icons
    'PopCntrl.VisualTheme = xtpPopupThemeMSN
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

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Exit SMIS, Are You Sure?", vbExclamation + vbOKCancel, "Exit System") = vbOK Then
        Dim FRM                                                       As Form
        For Each FRM In Forms
            If Not (FRM Is Nothing) Then
                Unload FRM
            End If
        Next
        CommandBars1.SaveCommandBars MODULENAME, App.TITLE, "Layout"
    Else
        Cancel = 1
        frmMainMenu.Show
        frmMainMenu.ZOrder 1
    End If
End Sub

Private Sub PopCntrl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.ID = 707 Then
        PopCntrl.Close
    End If

End Sub

''''''''''''''END REGION POPUPCONTROLS''''''''''''''''''''''''''''''''''''
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If TIMER_REMIND = "" Then
        ReminderModule ""
    Else
        If DateDiff("n", TIMER_REMIND, Now) >= 0 Then
            frmSMIS_Files_Reminders.Show
        End If
    End If
End Sub

'"add updating code here
Private Sub MDIForm_Load()
    ApplyThemes
    ConfigurePopUps
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

End Sub

Public Sub ApplyThemes()
    CommandBarsGlobalSettings.App = App
    With CommandBars1
        .EnableOffice2007Frame True
        .LoadDesignerBars
        '        .LoadCommandBars MODULENAME, App.TITLE, "Layout"
        .PaintManager.ClearTypeTextQuality = True
        .TabWorkspace.ThemedBackColor = False
        .StatusBar.Visible = True
        .Options.SyncFloatingToolbars = True
    End With
    With SkinFramework1
        '.LoadSkin "C:\DMIS 2.0\Styles\royale.cjstyle", ""
        .ApplyWindow Me.hwnd
        .ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or Not xtpSkinApplyMetrics
    End With
    Dim ToolTipContext                                                As ToolTipContext
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

Public Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Debug.Print Control.ID; " '"; Control.Caption; vbCrLf


    Select Case Control.ID
            '******************************************************************************************************
            '''''''''''''''''''''''''''''FILES'''''''''''''''''''''''''''''''''''''''
            '******************************************************************************************************
        Case 1233
            If Module_Access(LOGID, "PDI CHECKLIST", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_Files_PDICheckList.Show
        Case 1234
            If Module_Access(LOGID, "PDI SETUP", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_Files_PDISetup.Show
        Case 1194
            frmMainMenu.Show
        Case FILES_VEHICLESCLASS
            If Module_Access(LOGID, "VEHICLE CLASS", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_Files_VehiclesClass.Show
            frmSMIS_Files_VehiclesClass.ZOrder 0
        Case FILES_CUSTOMERINFORMATION, TOOL_CUSTOMERINFORMATION
            If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub
            Call frmAllCustomer.AddEditCustomer("")
            frmAllCustomer.Show
            frmAllCustomer.ZOrder 0
        Case FILES_VEHICLECOLOR
            If Module_Access(LOGID, "VEHICLE COLOR", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_Files_Color.Show
            frmSMIS_Files_Color.ZOrder 0
        Case FILES_VEHICLEMODEL, TOOL_VEHICLEDESCRIPTION
            If Module_Access(LOGID, "VEHICLE MODEL", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_Files_Model.Show
            frmSMIS_Files_Model.ZOrder 0
        Case FILES_FINANCINGCOMPANY
            If Module_Access(LOGID, "FINANCING COMPANY", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_Files_FinancingCo.Show
            frmSMIS_Files_FinancingCo.ZOrder 0
        Case FILES_SALESACCOUNTEXECUTIVE, TOOL_SALESACCOUNTEXECUTIVES
            If Module_Access(LOGID, "SALES ACCOUNT EXECUTIVE", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_Files_SalesAE.Show
            frmSMIS_Files_SalesAE.ZOrder 0
        Case 1252                                           'Signatories
            If Module_Access(LOGID, "SIGNATORIES", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_Files_Signatories.Show
        Case FILES_FINANCIALDOCUMENTS
            If Module_Access(LOGID, "FINANCIAL DOCUMENTS", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_Files_Document.Show
            frmSMIS_Files_Document.ZOrder 0
        Case FILE_PROSPECT
            If Module_Access(LOGID, "PROSPECT", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_Files_Prospects.Show
            frmSMIS_Files_Prospects.ZOrder 0
        Case FILES_LEADCLASSIFICATIONS
            If Module_Access(LOGID, "CLASSIFY LEADS", "DATA ENTRY") = False Then Exit Sub
            frmCRIS_ClassifyLeads.Show
            frmCRIS_ClassifyLeads.ZOrder 0
            '******************************************************************************************************
            '''''''''''''''''''''''''''''TRANSACTIONS'''''''''''''''''''''''''''''''''''''
            '******************************************************************************************************
        Case 1235
            If Module_Access(LOGID, "PDI CHECKLIST", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            frmSMIS_Trans_VehiclesCheckList.Show
            If FormExist("frmSMIS_Trans_VehiclesCheckList") Then
                frmSMIS_Trans_VehiclesCheckList.ZOrder 0
            End If
            Err.Clear
        Case 1239
            If Module_Access(LOGID, "QUOTATION", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            frmSMIS_Trans_Quotation.Show
        Case 1240                                           'Stock Transfer:
            If Module_Access(LOGID, "STOCK TRANSFER", "TRANSACTION") = False Then Exit Sub
            frmSMIS_Trans_MRR1.Show
            frmSMIS_Trans_MRR1.ZOrder 0

        Case INVENTORY_VEHICLERECIEVING, TOOL_CARINVENTORY
            If Module_Access(LOGID, "VEHICLE RECIEVING", "TRANSACTION") = False Then Exit Sub
            frmSMIS_Trans_MRR.Show
            frmSMIS_Trans_MRR.ZOrder 0
        Case TRANSACTIONS_VEHICLESSALESMONITORING
            If Module_Access(LOGID, "VEHICLES SALES MONITORING", "SYSTEM") = False Then Exit Sub
            MainForm.Show
            MainForm.ZOrder 0
        Case TRANSACTIONS_LOANAPPLICATION_INDIVIDUAL
            If Module_Access(LOGID, "INDIVIDUAL LOAN APPLICATION", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            frmSMIS_Trans_ApplicationIndividual.Show
            If FormExist("frmSMIS_Trans_ApplicationIndividual") Then
                frmSMIS_Trans_ApplicationIndividual.ZOrder 0
            End If
        Case TRANSACTIONS_LOANAPPLICATION_CORPORATE
            If Module_Access(LOGID, "CORPORATE LOAN APPLICATION", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            frmSMIS_Trans_ApplicationCorporate.Show
            If FormExist("frmSMIS_Trans_ApplicationCorporate") Then
                frmSMIS_Trans_ApplicationCorporate.ZOrder 0
            End If
        Case TRANSACTIONS_VEHICLESNVOICING, TOOL_CARINVOICING
            If Module_Access(LOGID, "VEHICLES INVOICING", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            frmSMIS_Trans_VehicleInvoice.Show
            If FormExist("frmSMIS_Trans_VehicleInvoice") Then
                frmSMIS_Trans_VehicleInvoice.ZOrder 0
            End If
        Case TRANSACTIONS_PURCHASEORDERDATAENTRY
            If Module_Access(LOGID, "PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
            frmSMIS_Trans_Ordering.Show
            frmSMIS_Trans_Ordering.ZOrder 0
        Case TRANSACTIONS_SALESORDER
            On Error Resume Next
            If Module_Access(LOGID, "SALES ORDER", "TRANSACTION") = False Then Exit Sub
            frmSMIS_Trans_SalesOrder.Show
            '******************************************************************************************************
            '''''''''''''''''''''''''''''INQUIRY''''''''''''''''''''''''''''''''''''''
            '******************************************************************************************************
        Case 1227
            If Module_Access(LOGID, "CUSTOMER SALES HISTORY", "INQUIRY") = False Then Exit Sub
            frmSMIS_Inquiry_CustomerSalesHistory.Show
            frmSMIS_Inquiry_CustomerSalesHistory.ZOrder 1
        Case 1228
            If Module_Access(LOGID, "CUSTOMER SERVICE HISTORY", "INQUIRY") = False Then Exit Sub
            frmCSMSCustomerHistory.Show
            frmCSMSCustomerHistory.ZOrder 0

        Case 1229
            If Module_Access(LOGID, "CUSTOMER CALL VISIT HISTORY", "INQUIRY") = False Then Exit Sub
            frmSMIS_Inquiry_CallVisit_History.Show
        Case 1230

        Case 1226
            If Module_Access(LOGID, "VEHICLE MASTER INQUIRY", "INQUIRY") = False Then Exit Sub
            frmSMIS_Inquiry_VehicleMaster.Show

        Case 1236
            If Module_Access(LOGID, "PENDING PO", "INQUIRY") = False Then Exit Sub
            If FormExist("frmSMIS_Inquiry_OverDuePending") Then
                Unload frmSMIS_Inquiry_OverDuePending
            End If
            frmSMIS_Inquiry_OverDuePending.ShowPendingOrders
            frmSMIS_Inquiry_OverDuePending.Show
        Case 1237
            If Module_Access(LOGID, "OVERDUE PO", "INQUIRY") = False Then Exit Sub
            If FormExist("frmSMIS_Inquiry_OverDuePending") Then
                Unload frmSMIS_Inquiry_OverDuePending
            End If
            frmSMIS_Inquiry_OverDuePending.ShowOverDueOrders
            frmSMIS_Inquiry_OverDuePending.Show
        Case 1238
            If Module_Access(LOGID, "SERVED PO", "INQUIRY") = False Then Exit Sub
            If FormExist("frmSMIS_Inquiry_OverDuePending") Then
                Unload frmSMIS_Inquiry_OverDuePending
            End If
            frmSMIS_Inquiry_OverDuePending.ShowServerdOrders
            frmSMIS_Inquiry_OverDuePending.Show
        Case INQUIRY_SALESAPPOINTMENTCALENDAR
            If Module_Access(LOGID, "SALES APPOINTMENT CALENDAR", "INQUIRY") = False Then Exit Sub
            frmSMIS_Inquiry_SalesAppointment.Show

        Case INQUIRY_ALLOCATEDCARS, TOOL_ALLOCATEDCARS
            If Module_Access(LOGID, "ALLOCATED CARS", "INQUIRY") = False Then Exit Sub
            frmSMIS_Inquiry.Show
            frmSMIS_Inquiry.optAllCars.Value = True
            frmSMIS_Inquiry.ZOrder 0
        Case INQUIRY_INVOICEDCARS, TOOL_INVOICEDCARS
            If Module_Access(LOGID, "INVOICED CARS", "INQUIRY") = False Then Exit Sub
            frmSMIS_Inquiry.Show
            frmSMIS_Inquiry.optInvCars.Value = True
            frmSMIS_Inquiry.ZOrder 0
        Case INQUIRY_TOTALVEHICLERELEASE, TOOL_TOTALVEHICLERELEASE
            If Module_Access(LOGID, "TOTAL RELEASED VEHICLES", "INQUIRY") = False Then Exit Sub

            frmSMIS_Inquiry.Show
            frmSMIS_Inquiry.optCarRelease.Value = True
            frmSMIS_Inquiry.ZOrder 0
        Case INQUIRY_VEHICLESONSTOCK, TOOL_VEHICLEONSTOCK

            If Module_Access(LOGID, "VEHICLES ON STOCK", "INQUIRY") = False Then Exit Sub

            frmSMIS_Inquiry.Show
            frmSMIS_Inquiry.optVehStock.Value = True
            frmSMIS_Inquiry.ZOrder 0
        Case INQUIRY_SALESEXECUTIVEPERFORMANCE, TOOL_SALESEXECUTIVEPERFORMANCE
            If Module_Access(LOGID, "SALES EXECUTIVE PERFORMANCE", "INQUIRY") = False Then Exit Sub
            frmSMIS_Inquiry.Show
            frmSMIS_Inquiry.optSalesPer.Value = True
            frmSMIS_Inquiry.ZOrder 0
        Case INQUIRY_PROSPECTINQUIRY
            If Module_Access(LOGID, "PROSPECT INQUIRY", "INQUIRY") = False Then Exit Sub
            frmSMIS_Inquiry_InquiryMain.optAdvSearch(0).Value = True
            Load frmSMIS_Inquiry_InquiryMain
            frmSMIS_Inquiry_InquiryMain.Show
        Case INQUIRY_SALESAPPOINTMENTBYSAE
            If Module_Access(LOGID, "SALES APPOINTMENT BY SAE", "INQUIRY") = False Then Exit Sub
            frmSMIS_Inquiry_InquiryMain.optAdvSearch(1).Value = True
            Load frmSMIS_Inquiry_InquiryMain
            frmSMIS_Inquiry_InquiryMain.Show
            frmSMIS_Inquiry_InquiryMain.ZOrder 0
            '****************************************************************************************************
            '''''''''''''''''''''''''''''REPORTS''''''''''''''''''''''''''''''''''''''''''
            '****************************************************************************************************
        Case ID_CUSTOMERDIRECTORYBYCUSTOMERTYPE
            If Module_Access(LOGID, "CUSTOMER DIRECTORY BY CUSTOMER TYPE", "REPORTS") = False Then Exit Sub
            Screen.MousePointer = 11
            rptMain.WindowTitle = "Customer Directory By Customer Type"
            rptMain.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptMain.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptMain, SMIS_REPORT_PATH & "CustomerListByCustType.rpt", "", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Case ID_VEHICLEMODEL
            If Module_Access(LOGID, "VEHICLE MODEL LIST", "REPORTS") = False Then Exit Sub
            Screen.MousePointer = 11
            rptMain.WindowTitle = "VEHICLE MODEL LIST"
            rptMain.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptMain.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

            PrintSQLReport rptMain, SMIS_REPORT_PATH & "listing/VehiclesGroupList.rpt", "", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Case 1249                                           'Customer Log Report
            If Module_Access(LOGID, "CUSTOMER LOG REPORT", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_Log.Show
        Case 1246                                           'Customer Information Report
            If Module_Access(LOGID, "CUSTOMER INFORMATION", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_CustomerInformationReport.Show
        Case 1247                                           'Customer Reminder && Task
            If Module_Access(LOGID, "CUSTOMER REMINDERS AND TASKS", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_CustomerRemindersAndTask.Show
        Case 1248                                           'Sales Appointment Report
            If Module_Access(LOGID, "SALES APPOINTMENT", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_SalesAppointment.Show
        Case 1250                                           'Service Appointment Report
            If Module_Access(LOGID, "SERVICE APPOINTMENT", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_ServiceAppointment.Show
        Case 1251                                           'Internal Reminder
            If Module_Access(LOGID, "INTERNAL REMINDERS", "REPORTS") = False Then Exit Sub
            frmCRIS_Report_InternalReminder.Show
        Case 1138
            If Module_Access(LOGID, "SALES AND STOCK TRACKING REPORT", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_GenDSSR.Show
            frmSMIS_Report_GenDSSR.ZOrder 0
        Case REPORT_INV_MONTHLYPURCHASESREPORT
            If Module_Access(LOGID, "MONTHLY PURCHASES REPORT", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_Purchases2.Show
            frmSMIS_Report_Purchases2.ZOrder 0
        Case REPORT_INV_MONTHLYINVENTORYCONTROL
            If Module_Access(LOGID, "MONTHLY INVENTORY CONTROL", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_InvControl.Show
        Case REPORT_INV_ENDINGINVENTORY
            If Module_Access(LOGID, "MONTHLY ENDING INVENTORY", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_InvControl2.Show
            frmSMIS_Report_InvControl2.ZOrder 0
        Case REPORT_INV_VEHICLEONSTOCK
            If Module_Access(LOGID, "VEHICLE ON STOCK", "REPORTS") = False Then Exit Sub
            rptMain.ReportFileName = SMIS_REPORT_PATH & "vehstock.rpt"
            rptMain.Connect = DMIS_REPORT_Connection
            rptMain.Action = 1
        Case REPORT_INV_VEHICLEINVENTORY
            If Module_Access(LOGID, "VEHICLE INVENTORY REPORT", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_VehiclesInventory.Show
            frmSMIS_Report_VehiclesInventory.ZOrder 0

        Case REPORT_INV_DELIVERYUNITSREPORT
            If Module_Access(LOGID, "DELIVERY UNITS REPORT", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_DelReport.Show
            frmSMIS_Report_DelReport.ZOrder 0
        Case REPORT_SALES_UNITSRELEASED
            If Module_Access(LOGID, "UNITS RELEASED REPORT", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_Released.Show
            frmSMIS_Report_Released.ZOrder 0
        Case REPORT_SALES_MONTHLYVEHICLEGROSSPROFIT
            If Module_Access(LOGID, "MONTHLY VEHICLE GROSS PROFIT", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_GenRep.Show
            frmSMIS_Report_GenRep.ZOrder 0
        Case REPORT_SALES_YEARLYVEHICLEGROSSPROFIT
            If Module_Access(LOGID, "YEARLY VEHICLE GROSS PROFIT", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_RepYearly.Show
        Case REPORT_SALES_VEHICLESALESREPORT
            If Module_Access(LOGID, "VEHICLE SALES", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_VSRep.Show
            frmSMIS_Report_VSRep.ZOrder 0
        Case REPORT_SALES_DISTRIBUTIONOFSALESASTOMODEOFPAYMENT
            If Module_Access(LOGID, "SALES DISTRIBUTION OF SALES AS TO MODE OF PAYMENT", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_DistModePayment.Show
            frmSMIS_Report_DistModePayment.ZOrder 0
        Case REPORT_SALES_PROJECTEDVEHICLESALES
            If Module_Access(LOGID, "PROJECTED VEHICLE SALES REPORT", "REPORTS") = False Then Exit Sub
        Case REPORT_OTHER_LISTOFUNITSREGISTERED
            If Module_Access(LOGID, "LIST OF UNITS REGISTERED", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_ListReg.Show
            frmSMIS_Report_ListReg.ZOrder 0

        Case REPORT_OTHER_SALESEXECUTIVEPERFORMANCE
            If Module_Access(LOGID, "SALES EXECUTIVE PERFORMANCE", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_SAEPer.Show
            frmSMIS_Report_SAEPer.ZOrder 0
        Case REPORT_OTHER_VEHICLESALESCUSTOMERSSUMMARY
            If Module_Access(LOGID, "VEHICLE SALES CUSTOMERS SUMMARY", "REPORTS") = False Then Exit Sub
            CUST_REPT_TYPE = "1"
            frmSMIS_Report_CustSummary.Show
            frmSMIS_Report_CustSummary.ZOrder 0
        Case REPORT_OTHER_CUSTOMERSWITHINSURANCEPOLICIES
            If Module_Access(LOGID, "CUSTOMERS WITH INSURANCE POLICIES", "REPORTS") = False Then Exit Sub
            CUST_REPT_TYPE = "2"
            frmSMIS_Report_CustSummary.Show
            frmSMIS_Report_CustSummary.ZOrder 0
        Case REPORT_MARKETTING_CUSTOMERSDIRECTORY
            If Module_Access(LOGID, "REPORT CUSTOMERS DIRECTORY", "REPORTS") = False Then Exit Sub
            rptMain.ReportFileName = SMIS_REPORT_PATH & "invoices.rpt"
            rptMain.Connect = DMIS_REPORT_Connection
            rptMain.Action = 1
        Case REPORT_MARKETTING_BIRTHDAYCELEBRANTSOFTHEMONTH
            If Module_Access(LOGID, "BIRTHDAY CELEBRANTS", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_CustomerBDays.Show
            frmSMIS_Report_CustomerBDays.ZOrder 0
        Case REPORT_GOV_BIRYEARREPORT
            If Module_Access(LOGID, "BIR YEAR REPORT", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_BIRYearEnd.Show
            frmSMIS_Report_BIRYearEnd.ZOrder 0
        Case REPORT_MONTHLYSALESANALYSISREPORT
            If Module_Access(LOGID, "MONTHLY SALES ANALYSIS", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_MonthlySalesAnalysis.Show
            frmSMIS_Report_MonthlySalesAnalysis.ZOrder 0
        Case REPORT_YEARLYSALESANALYSISREPORT
            If Module_Access(LOGID, "YEARLY SALES ANALYSIS", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_YearlySalesAnalysis.Show
            frmSMIS_Report_YearlySalesAnalysis.ZOrder 0
        Case REPORT_RANGEDSALESANALYSISREPORT
        Case REPORT_OTHER_MONTHLYSALESEXECUTIVELIST
            If Module_Access(LOGID, "MONTLY SALES EXECUTIVE LIST", "REPORTS") = False Then Exit Sub
            frmSMIS_Report_MonthlySAEList.Show
            frmSMIS_Report_MonthlySAEList.ZOrder 0
        Case REPORT_OTHER_ID_SALESEXECUTIVELISTING
            If Module_Access(LOGID, "SALES EXECITIVE LISTING", "REPORTS") = False Then Exit Sub
            Dim FILTER                                                As String
            Screen.MousePointer = 11
            With rptMain
                .ReportTitle = "SALES EXECITIVE LISTING"
                .Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                .Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                .Connect = DMIS_REPORT_Connection
                .WindowTitle = "SALES PERSONNEL LIST"
                .ReportFileName = SMIS_REPORT_PATH & "listing\sae.rpt"
                .Action = 1
                Screen.MousePointer = 0
            End With
            '****************************************************************************************************
            '''''''''''''''''''''''''''''MAINTAIN''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '****************************************************************************************************
        Case MAINTAIN_COMPANYPROFILE
            If Module_Access(LOGID, "MAINTAIN COMPANY PROFILE", "SYSTEM") = False Then Exit Sub
            frmSMIS_Files_Profile.Show
        Case MAINTAIN_PASSWORDMAINTENANCE
            frmAccMaintenance.Show
            frmAccMaintenance.ZOrder 0
        Case WINDOW_ABOUT, TOOL_ABOUTTHEAUTHOR
            frmAbout.Show
        Case WINDOW_EXIT, TOOL_EXITSYSTEM
            Unload Me
        Case TOOL_DASHBOARD
            frmMainMenu.Show
            frmMainMenu.ZOrder 0
        Case 1223                                           'Prospect Log
            frmSMIS_Log_Menu.picLogCustomer.Visible = False
            frmSMIS_Log_Menu.picLogProspect.Visible = True
            frmSMIS_Log_Menu.Show

        Case 1225                                           'Customer Log
            frmSMIS_Log_Menu.picLogCustomer.Visible = True
            frmSMIS_Log_Menu.picLogProspect.Visible = False
            frmSMIS_Log_Menu.Show
        Case 1253  'AOR/OMA/DI Set Up
            If Module_Access(LOGID, "AOR-OMA-DI MASTER FILE", "DATA ENTRY") = False Then Exit Sub
            frmSMIS_File_AORRate.Show
            frmSMIS_File_AORRate.ZOrder 0
        Case Else
Debug.Print " CASE " & Control.ID; " '"; Control.Caption; vbCrLf
            '

    End Select


End Sub


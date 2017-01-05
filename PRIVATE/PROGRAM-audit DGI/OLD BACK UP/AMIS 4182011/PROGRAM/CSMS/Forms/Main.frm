VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#10.4#0"; "CO15D0~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#10.4#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000F&
   Caption         =   "Car Service Management System"
   ClientHeight    =   7170
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15150
   Icon            =   "Main.frx":0000
   NegotiateToolbars=   0   'False
   Picture         =   "Main.frx":030A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1770
      Top             =   120
   End
   Begin Crystal.CrystalReport rptMain 
      Left            =   480
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Customer Directory Listing"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin wizEncrypt.wizEnc wizEnc1 
      Left            =   -900
      Top             =   7260
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6915
      Width           =   15150
      _ExtentX        =   26723
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
            TextSave        =   "10:45 AM"
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
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   1380
      Top             =   120
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.PopupControl PopCntrl 
      Left            =   60
      Top             =   90
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
      Animation       =   1
      AnimateDelay    =   125
      ShowDelay       =   2500
      BackgroundBitmap=   60
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   960
      Top             =   120
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   1
      DesignerControls=   "Main.frx":17BDC
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frm                                                As frmSMIS_Files_Reminders

Sub ApplyPatches()

End Sub

Private Sub MDIForm_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    ApplyThemes
    ConfigurePopUps
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Exit CSMS? Are you Sure?", vbOKCancel + vbExclamation, "Exit Application") = vbOK Then
        Dim frm                                        As Form
        For Each frm In Forms
            If Not (frm Is Nothing) Then
                Unload frm
            End If
        Next
        CommandBars1.SaveCommandBars MODULENAME, App.TITLE, "Layout"
        UnloadForm Me
    Else
        Cancel = 1
        frmMainMenu.Show
    End If
End Sub

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '    MsgBox Control.ID

    Select Case Control.ID
        Case FILES_CUSTOMERMASTERLIST, 1088
            If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub
            frmAllCustomer.Show
        Case FILES_SERVICE_SERVICEADVISOR, 1089
            If Module_Access(LOGID, "SERVICE ADVISOR", "DATA ENTRY") = False Then Exit Sub

        Case FIILES_SERVICE_TECHNICIAN, 1090
            If Module_Access(LOGID, "TECHNICIAN", "DATA ENTRY") = False Then Exit Sub

        Case FILES_SERVICE_GENERALJOBS, 1091
            If Module_Access(LOGID, "JOBS", "DATA ENTRY") = False Then Exit Sub
            frmCSMSReqJobs.Show

        Case FILES_SERVICE_PMSJOBS, 1092
            If Module_Access(LOGID, "PMS JOBS", "DATA ENTRY") = False Then Exit Sub
            frmCSMSAddPms.Show
        Case FILES_SERVICE_CANNEDLABOR, 1093
            If Module_Access(LOGID, "CANNED LABOR", "DATA ENTRY") = False Then Exit Sub
            frmCSMSCannedlabor.Show
        Case FILES_SERVICE_MODEL, 1094
            If Module_Access(LOGID, "MODEL", "DATA ENTRY") = False Then Exit Sub
            frmCSMSModel.Show

        Case TRANS_SERVICECOUNTER
            If Module_Access(LOGID, "SERVICE COUNTER", "SYSTEM") = False Then Exit Sub
            frmCSMS_ServiceCounter.Show
            
        Case TRANS_JOBESTIMATE
            If Module_Access(LOGID, "JOB ESTIMATE", "TRANSACTION") = False Then Exit Sub
            frmCSMSEstimateEntry.Show
            
        Case TRANS_REPAIRORDER
            If Module_Access(LOGID, "BILLING SYSTEM", "TRANSACTION") = False Then Exit Sub
            With frmCSMSDataEntry
                .Show
                .Caption = "Repair Order Data Entry"
                .SSTab1.TabEnabled(2) = False
                .SSTab1.TabEnabled(3) = False
                .cmdReleaseRO.Enabled = False
                '.cmdAdd.Enabled = True
                '.pic3.Visible = False
                '.pic4.Visible = False
            End With
        Case TRANS_BILLINGSYSTEM
            If Module_Access(LOGID, "BILLING SYSTEM", "TRANSACTION") = False Then Exit Sub
            With frmCSMSDataEntry
                .Show
                .Caption = "Billing System"
                .SSTab1.TabEnabled(2) = True
                .SSTab1.TabEnabled(3) = True
                '.cmdAdd.Enabled = False
                .pic3.Visible = True
                '.pic4.Visible = True
            End With
        Case INQUIRY_CUSTOMERVEHICLEINQUIRY
            If Module_Access(LOGID, "CUSTOMER VEHICLE INQUIRY", "INQUIRY") = False Then Exit Sub
            frmCSMSCustomerHistory.Show
        
        Case INQUIRY_SERVICEADVISORSWORKDETAILS
            If Module_Access(LOGID, "SERVICE ADVISOR WORK DETAILS", "INQUIRY") = False Then Exit Sub
            
            
        Case INQUIRY_LISTBYREPAIRORDER
            If Module_Access(LOGID, "LIST BY REPAIR ORDER", "INQUIRY") = False Then Exit Sub

        Case INQUIRY_LISTBYINVOICENUMBER
            If Module_Access(LOGID, "LIST BY INVOICE NUMBER", "INQUIRY") = False Then Exit Sub
            
        Case INQUIRY_LISTBYPLATENUMBER
            If Module_Access(LOGID, "LIST BY PLATE NUMBER", "INQUIRY") = False Then Exit Sub
            
        Case INQUIRY_JOBESTIMATELISTING
            If Module_Access(LOGID, "JOB ESTIMATE LISTING", "INQUIRY") = False Then Exit Sub
            frmCSMS_INQUIRY_JobEstimate.Show
        Case INQUIRY_PARTSINQUIRY
            If Module_Access(LOGID, "PARTS INQUIRY", "INQUIRY") = False Then Exit Sub
            
        Case INQUIRY_PARTSLISTING
            If Module_Access(LOGID, "PARTS LISTING", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            PARTSQUERY = 1
            CSMS_PARTSQUERY = True
        Case RPT_CUSTOMERDIRECTORYLISTING
            If Module_Access(LOGID, "CUSTOMER DIRECTORY LISTING", "REPORTS") = False Then Exit Sub
            PrintSQLReport rptMain, CSMS_REPORT_PATH & "Customer.rpt", "", CSMS_REPORT_CONNECTION, 1
        Case RPT_TECHNICIANLABOREFFICIENCY, 1081
            If Module_Access(LOGID, "TECHNICIAN LABOR EFFICIENCY", "REPORTS") = False Then Exit Sub
            frmCSMSTechnician_Efficiency.Show
        Case RPT_TRANSACTIONSFORFOLLOWUP, 1067
            If Module_Access(LOGID, "TRANSACTIONS FOR FOLLOW UP", "REPORTS") = False Then Exit Sub
            frmCSMSFor_followUp.Show
        Case RPT_VEHICLEBYMODEL
            If Module_Access(LOGID, "VEHICLE BY MODEL", "REPORTS") = False Then Exit Sub
            frmCSMSVehicleByModel.Show
        Case RPT_SERVICEREPORT
            If Module_Access(LOGID, "SERVICE REPORT", "REPORTS") = False Then Exit Sub
            frmCSMSServiceReport.Show

        Case MAINTAIN_COMPANYPROFILE
            If Module_Access(LOGID, "COMPANY PROFILE", "SYSTEM") = False Then Exit Sub
            frmCSMSProfile.Show

        Case MAINTAIN_USERMODULES
        Case MAINTAIN_PASSWORDMAINTENANCE
            frmAccMaintenance.Show
        Case WINDOW_ABOUT
            frmAbout.Show
        Case WINDWO_EXIT
            'If MsgBox("Close " & MODULENAME & " Application?", vbQuestion + vbYesNo, "Exiting System...") = vbYes Then
            If MsgBox("Exit " & MODULENAME & "? Are You Sure?", vbQuestion + vbYesNo, "Exit Application") = vbYes Then
                End
            End If
        Case 1064
            frmMainMenu.Show
            frmMainMenu.ZOrder 0
        Case 1066
            If Module_Access(LOGID, "AFTER SALES REPORT", "REPORTS") = False Then Exit Sub
            frmCSMSAfterSalesServiceReport.Show
            
        Case 1068
            If Module_Access(LOGID, "MONTHLY PERFORMANCE REPORT", "REPORTS") = False Then Exit Sub
            frmCSMSHyundaiMonthlyPerformanceReport.Show
            
        Case 1069
            If Module_Access(LOGID, "ACTUAL MANNING REPORT", "REPORTS") = False Then Exit Sub
            frmCSMSActualManningReport.Show
        Case 1070
            If Module_Access(LOGID, "WORKSHOP SALES WEEKLY PERFORMANCE REPORT", "REPORTS") = False Then Exit Sub
            frmCSMSWorkshopSalesWeeklyPerformanceReport.Show
        Case 1071
            If Module_Access(LOGID, "UNITS RECEIVE WEEKLY PERFORMANCE REPORT", "REPORTS") = False Then Exit Sub
            frmCSMSUnitsReceivedWeeklyPerformanceReport.Show
        Case 1072
            If Module_Access(LOGID, "WORKING IN PROGRRESS", "REPORTS") = False Then Exit Sub
            frmCSMSWorkInProgress.Show
        Case 1073
            If Module_Access(LOGID, "SERVICE ADVISOR REPORT", "REPORTS") = False Then Exit Sub
            frmCSMSServiceAdvisorReport.Show
        Case 1074
            If Module_Access(LOGID, "VEHICLE AGING PROGRESS REPORT", "REPORTS") = False Then Exit Sub
            frmCSMSVehicleAgingOnProcessReport.Show
        Case 1075
            MsgBox "Module Under Revision", vbInformation, "CSMS"
            'If Module_Access(LOGID, "PARTS PICK LIST", "REPORTS") = False Then Exit Sub
            'frmCSMSPartsPickList.Show
        Case 1076
            If Module_Access(LOGID, "APPOINTMENT DIARY", "REPORTS") = False Then Exit Sub
            frmCSMSAppointmentDiary.Show
        Case 1079
            If Module_Access(LOGID, "WARRANTY REPORTS", "REPORTS") = False Then Exit Sub
            frmCSMS_WarRep.Show
        Case 1080
            If Module_Access(LOGID, "TECHNICIAN REPORT", "REPORTS") = False Then Exit Sub
            frmTechnicianReport.Show
        Case 1086
            If Module_Access(LOGID, "COMPLAINTS FORM", "DATA ENTRY") = False Then Exit Sub
            FrmCSMSComplaintsForm.Show
        Case 1087
            If Module_Access(LOGID, "CONCERN RESOLUTION", "DATA ENTRY") = False Then Exit Sub
            frmCSMSConcernResolution.Show
        Case 1083
            If Module_Access(LOGID, "JOB CLOCK", "SYSTEM") = False Then Exit Sub
            frmCSMSClockINOUT.Show
        Case 1084
            If Module_Access(LOGID, "QUALITY INFORMATION REPORT", "TRANSACTION") = False Then Exit Sub
            frmCSMS_CQI.Show
        Case 1082
            MsgBox "Module Under Revision", vbInformation, "CSMS"
            Exit Sub
            
        Case 1101
            frmMainMenu.Show


    End Select
    Screen.MousePointer = 0
End Sub

Private Sub ApplyThemes()
    CommandBarsGlobalSettings.App = App
    'Office2007.cjstyles:
    '    -NORMALAQUA.INI
    '    -NORMALBLUE.INI
    'Vista.cjstyles:
    '    -NORMALBLACK.INI
    '    -NORMALBLUE.INI
    '    -NORMALSILVER.INI
    'WinXP.Luna.cjstyles:
    '    -EXTRALARGEBLUE.INI
    '    -EXTRALARGEHOMESTEAD.INI
    '    -EXTRALARGEMETALLIC.INI
    '    -LARGEBLUE.INI
    '    -LARGEHOMESTEAD.INI
    '    -LARGEMETALLIC.INI
    '    -NORMALBLUE.INI
    '    -NORMALHOMESTEAD.INI
    '    -NORMALMETALLIC.INI
    'WinXP.Royale.cjstyles:
    '    -EXTRALARGEFONTSROYALE.INI
    '    -LARGEFONTSROYALE.INI
    '    -NORMALROYALE.INI


    With CommandBars1
        .LoadDesignerBars
        .LoadCommandBars MODULENAME, App.TITLE, "Layout"
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
    Dim ToolTipContext                                 As ToolTipContext
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
    Dim Item                                           As PopupControlItem
    PopCntrl.RemoveAllItems
    PopCntrl.Icons.AddIcons CommandBars1.Icons

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

Private Sub PopCntrl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.ID = 707 Then
        PopCntrl.Close
    End If
End Sub

Private Sub Timer1_Timer()
    If gconDMIS Is Nothing Then Exit Sub
    If TIMER_REMIND = "" Then
        ReminderModule ""
    Else
        If DateDiff("n", TIMER_REMIND, Now) >= 0 Then
            Set frm = New frmSMIS_Files_Reminders
            frm.Show
        End If
    End If
End Sub


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#10.4#0"; "CO15D0~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#10.4#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Parts Management Information System"
   ClientHeight    =   6645
   ClientLeft      =   3285
   ClientTop       =   3240
   ClientWidth     =   15240
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "Main.frx":030A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin wizEncrypt.wizEnc wizEnc1 
      Left            =   -1920
      Top             =   6360
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6390
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
            TextSave        =   "12:58 AM"
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
            Object.Width           =   4313
            MinWidth        =   4304
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
   Begin MSComctlLib.ImageList IMG 
      Left            =   3870
      Top             =   1230
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
            Picture         =   "Main.frx":14CD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":14FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15440
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1575A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":160A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":164FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1C794
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1CAAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1CF00
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1D21A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1D534
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1D84E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1DCA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1E0F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1E40C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1E726
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1EA40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5205
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport rptMain 
      Left            =   4620
      Top             =   1380
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   6390
      Top             =   1440
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2880
      Top             =   1410
   End
   Begin XtremeSuiteControls.PopupControl PopCntrl 
      Left            =   3420
      Top             =   1410
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
      Left            =   5880
      Top             =   1440
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   6
      DesignerControls=   "Main.frx":247A2
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LOGDT                                              As Integer

Private Sub MDIForm_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    LOGDT = 1
    ApplyThemes
    ConfigurePopUps
    FLAG = 0                                          'Set command bars control to false
    If LOGDATE <> lastDay(LOGDATE) Then
        LOGDT = 0
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("EXIT PMIS, Are You Sure?", vbExclamation + vbOKCancel, "Exit System") = vbOK Then
        Dim FRM                                        As Form
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

Private Sub MDIForm_Resize()
    CenterMe Me, Me, 0
End Sub

Private Sub ApplyThemes()
    CommandBarsGlobalSettings.App = App
    With CommandBars1
        .LoadDesignerBars
        .LoadCommandBars MODULENAME, App.TITLE, "Layout"
        .PaintManager.ClearTypeTextQuality = True
        .TabWorkspace.ThemedBackColor = False
        .StatusBar.Visible = True
        .Options.SyncFloatingToolbars = True
    End With
    With SkinFramework1
        .LoadSkin "C:\DMIS 2.0\Styles\royale.cjstyle", ""
        '.LoadSkin App.Path & "\Royale.cjstyles", ""
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

''''''''''''''START REGION POPUPCONTROLS''''''''''''''''''''''''''''''''''''
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

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
            '***************************************************************************
        Case 1287
            If Module_Access(LOGID, "UPDATE LOCATION", "SYSTEM") = False Then Exit Sub
            FormExistsShow frmPMISUpdateLocation
        Case 1289
            If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub
            FormExistsShow frmAllCustomer
        Case 1290
            If Module_Access(LOGID, "MASTER HARIPARTS", "DATA ENTRY") = False Then Exit Sub
            FormExistsShow frmPMISMaster_DNPPEntry
        Case 1291
            If Module_Access(LOGID, "VENDORS", "DATA ENTRY") = False Then Exit Sub
            FormExistsShow frmAMISMASTERFILEVendor
        Case 1294
            If Module_Access(LOGID, "PARTS MASTER FILE", "DATA ENTRY") = False Then Exit Sub
            frmMasterFile_Parts.SetStockType ("P")
            FormExistsShow frmMasterFile_Parts
        Case 1295
            If Module_Access(LOGID, "MATERIALS", "DATA ENTRY") = False Then Exit Sub
            frmMasterFile_Material.SetStockType ("M")
            FormExistsShow frmMasterFile_Material
        Case 1542                                     'Accessories Master File
            If Module_Access(LOGID, "ACCESSORIES MASTER FILE", "DATA ENTRY") = False Then Exit Sub
            frmMasterFile_Accessories.SetStockType ("A")
            FormExistsShow frmMasterFile_Accessories
        Case 1296
            If Module_Access(LOGID, "SALESMAN MASTER FILE", "DATA ENTRY") = False Then Exit Sub
            FormExistsShow frmPMISMaster_SalesMan
        Case 1297
            If Module_Access(LOGID, "PARTS COUNTER", "DATA ENTRY") = False Then Exit Sub
            frmMasterFile_Counter_Parts.SetStockType ("P")
            FormExistsShow frmMasterFile_Counter_Parts
        Case 1299
            Unload frmPMIS_Physical_INVMenu_New
            C_TYPE = "P"
            FormExistsShow frmPMIS_Physical_INVMenu_New
        Case 1301
            If Module_Access(LOGID, "PHYSICAL COUNT", "SYSTEM") = False Then Exit Sub
            C_TYPE = "P"
            FormExistsShow frmPMIS_Physical_CreateINVDATA
        Case 1559
            Unload frmPMIS_Physical_INVMenu_New
            C_TYPE = "A"
            FormExistsShow frmPMIS_Physical_INVMenu_New
        Case 1560
            If Module_Access(LOGID, "PHYSICAL COUNT", "SYSTEM") = False Then Exit Sub
            Unload frmPMIS_Physical_CreateINVDATA
            C_TYPE = "M"
            FormExistsShow frmPMIS_Physical_CreateINVDATA
        Case 1561
            Unload frmPMIS_Physical_INVMenu_New
            C_TYPE = "M"
            FormExistsShow frmPMIS_Physical_INVMenu_New
        Case 1562
            If Module_Access(LOGID, "PHYSICAL COUNT", "SYSTEM") = False Then Exit Sub
            Unload frmPMIS_Physical_CreateINVDATA
            C_TYPE = "M"
            FormExistsShow frmPMIS_Physical_CreateINVDATA
        Case 1302
            If Module_Access(LOGID, "LOCATION", "REPORTS") = False Then Exit Sub
            frmPMISReports_Location_Parts.SETSTOCK_TYPE ("P")
            FormExistsShow frmPMISReports_Location_Parts
        Case 1467
            If Module_Access(LOGID, "PARTS REQUISTION SLIP", "TRANSACTION") = False Then Exit Sub
            WAREHOUSETYPE = "PRS"
            frmPMISTrans_PrisForms.txtTranType.Text = "PRS"
            FormExistsShow frmPMISTrans_PrisForms

        Case 1303
            If Module_Access(LOGID, "PARTS ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder
            COUNTERTYPE = "CSH"
            frmPMISTrans_CustomerOrder.txtTranType.Text = "CSH"
            FormExistsShow frmPMISTrans_CustomerOrder
        Case 1473                                     'Parts Issuance (Over the Counter) - Charge
            If Module_Access(LOGID, "PARTS ISSUANCE COUNTER CHARGE", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder
            COUNTERTYPE = "CHG"
            frmPMISTrans_CustomerOrder.txtTranType.Text = "CHG"
            FormExistsShow frmPMISTrans_CustomerOrder

        Case 1305
            If Module_Access(LOGID, "PARTS ISSUANCE SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder
            COUNTERTYPE = "RIV"
            frmPMISTrans_CustomerOrder.txtTranType.Text = "RIV"
            FormExistsShow frmPMISTrans_CustomerOrder
        Case 1306
            If Module_Access(LOGID, "PARTS DR OUT ISSUANCE", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder
            COUNTERTYPE = "DR"
            frmPMISTrans_CustomerOrder.txtTranType.Text = "DR"
            FormExistsShow frmPMISTrans_CustomerOrder
        Case 1307
            If Module_Access(LOGID, "PARTS ADVANCE BILL DATA ENTRY", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder
            COUNTERTYPE = "ADB"
            frmPMISTrans_CustomerOrder.txtTranType.Text = "ADB"
            FormExistsShow frmPMISTrans_CustomerOrder
        Case 1318
            If Module_Access(LOGID, "PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISTrans_Purchase
        Case 1319
            If Module_Access(LOGID, "PURCHASE RECEIVING STORING", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISTrans_Receiving2
        Case 1478                                     'Accessoires Requisition
            On Error Resume Next
            If Module_Access(LOGID, "ACCESSORIES REQUISITION SLIP", "TRANSACTION") = False Then Exit Sub
            Unload frmPMISAC_ARISForms
            WAREHOUSETYPE = "ARS"
            frmPMISAC_ARISForms.txtTranType.Text = "ARS"
            FormExistsShow frmPMISAC_ARISForms
        Case 1474                                     '&Accessories Issuance (Over the Counter) - Cash
            On Error Resume Next
            If Module_Access(LOGID, "ACCESSORIES ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub
            Unload frmPMISTrans_CustomerOrder_AC
            COUNTERTYPE = "CSH"
            frmPMISTrans_CustomerOrder_AC.txtTranType.Text = "CSH"
            FormExistsShow frmPMISTrans_CustomerOrder_AC
        Case 1475                                     '&Accessories Issuance (Over the Counter) - Charge
            On Error Resume Next
            If Module_Access(LOGID, "ACCESSORIES ISSUANCE COUNTER CHARGE", "TRANSACTION") = False Then Exit Sub
            Unload frmPMISTrans_CustomerOrder_AC
            COUNTERTYPE = "CHG"
            frmPMISTrans_CustomerOrder_AC.txtTranType.Text = "CHG"
            FormExistsShow frmPMISTrans_CustomerOrder_AC
        Case 1476                                     'Accessories Issuance (Service Issuance)
            If Module_Access(LOGID, "ACCESSORIES SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder_AC
            COUNTERTYPE = "RIV"
            frmPMISTrans_CustomerOrder_AC.txtTranType.Text = "RIV"
            FormExistsShow frmPMISTrans_CustomerOrder_AC
        Case 1477                                     'Accessories DR Out Issuance
            If Module_Access(LOGID, "ACCESSORIES DR OUT ISSUANCE", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder_AC
            COUNTERTYPE = "DR"
            frmPMISTrans_CustomerOrder_AC.txtTranType.Text = "DR"
            FormExistsShow frmPMISTrans_CustomerOrder_AC
        Case 1480                                     'Accessories Purchase Order Data Entry
            If Module_Access(LOGID, "ACCESSORIES PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            FormExistsShow frmPMISAC_Purchase
        Case 1481                                     'Accessories Receiving && Storing
            If Module_Access(LOGID, "ACCESSORIES RECEIVING", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISAC_Receiving
        Case 1486                                     'Material Purchase Order Data Entry
            If Module_Access(LOGID, "MATERIALS PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISMAT_Purchase
        Case 1487                                     'Materials Receiving && Storing
            If Module_Access(LOGID, "MATERIALS RECEIVING", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISMAT_Receiving
        Case 1479                                     'Materials Requisition
            On Error Resume Next
            If Module_Access(LOGID, "MATERIALS REQUISITION SLIP", "TRANSACTION") = False Then Exit Sub
            Unload frmPMISMAT_MRISForms
            WAREHOUSETYPE = "MRS"
            frmPMISMAT_MRISForms.txtTranType.Text = "MRS"
            FormExistsShow frmPMISMAT_MRISForms
        Case 1482                                     '&Materials Issuance (Over the Counter) - Cash
            On Error Resume Next
            If Module_Access(LOGID, "MATERIALS ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub
            Unload frmPMISTrans_CustomerOrder_MAT
            COUNTERTYPE = "CSH"
            frmPMISTrans_CustomerOrder_MAT.txtTranType.Text = "CSH"
            FormExistsShow frmPMISTrans_CustomerOrder_MAT
        Case 1483                                     '&Materials Issuance (Over the Counter) - Charge
            On Error Resume Next
            If Module_Access(LOGID, "MATERIALS ISSUANCE COUNTER CHARGE", "TRANSACTION") = False Then Exit Sub
            Unload frmPMISTrans_CustomerOrder_MAT
            COUNTERTYPE = "CHG"
            frmPMISTrans_CustomerOrder_MAT.txtTranType.Text = "CHG"
            FormExistsShow frmPMISTrans_CustomerOrder_MAT

        Case 1484                                     'Materials Issuance (Service Issuance)
            If Module_Access(LOGID, "MATERIALS SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder_MAT
            COUNTERTYPE = "RIV"
            frmPMISTrans_CustomerOrder_MAT.txtTranType.Text = "RIV"
            FormExistsShow frmPMISTrans_CustomerOrder_MAT
        Case 1485                                     'Materials DR Out Issuance
            If Module_Access(LOGID, "TRANSACTION DR OUT ISSUANCE", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder_MAT
            COUNTERTYPE = "DR"
            frmPMISTrans_CustomerOrder_MAT.txtTranType.Text = "DR"
            FormExistsShow frmPMISTrans_CustomerOrder_MAT
        Case 1470
            If Module_Access(LOGID, "PARTS QUALITY INFORMATION REPORT", "DATA ENTRY") = False Then Exit Sub
            FormExistsShow frmPMISTrans_PQIR
        Case 1469
            If Module_Access(LOGID, "DEALER PART INQUIRY", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISTrans_DealerPartInquiry
        Case 1308
            If Module_Access(LOGID, "PARTS INVENTORY ADJUSTMENT", "TRANSACTION") = False Then Exit Sub
            frmPMISTrans_InventoryAdjustment_Parts.SetStockType ("P")
            FormExistsShow frmPMISTrans_InventoryAdjustment_Parts
        Case 1320
            If Module_Access(LOGID, "TRANSACTION HISTORY CASH COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist
            COUNTERTYPE = "CSH"
            frmPMISInquiry_CustomerOrderHist.txtTranType.Text = "CSH"
            FormExistsShow frmPMISInquiry_CustomerOrderHist
        Case 1321
            If Module_Access(LOGID, "TRANSACTION HISTORY CHARGE COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist
            COUNTERTYPE = "CHG"
            frmPMISInquiry_CustomerOrderHist.txtTranType.Text = "CHG"
            FormExistsShow frmPMISInquiry_CustomerOrderHist
        Case 1322
            If Module_Access(LOGID, "TRANSACTION HISTORY REQUISTION ISSUANCE VOUCHER", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist
            COUNTERTYPE = "RIV"
            frmPMISInquiry_CustomerOrderHist.txtTranType.Text = "RIV"
            FormExistsShow frmPMISInquiry_CustomerOrderHist
        Case 1323
            If Module_Access(LOGID, "TRANSACTION HISTORY DR OUT ISSUANCE", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist
            COUNTERTYPE = "DR"
            frmPMISInquiry_CustomerOrderHist.txtTranType.Text = "DR"
            FormExistsShow frmPMISInquiry_CustomerOrderHist
        Case 1324
            If Module_Access(LOGID, "TRANSACTION HISTORY ADVANCE BILL DATA ENTRY", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist
            COUNTERTYPE = "ADB"
            frmPMISInquiry_CustomerOrderHist.txtTranType.Text = "ADB"
            FormExistsShow frmPMISInquiry_CustomerOrderHist
        Case 1325
            If Module_Access(LOGID, "PARTS TRANSACTION HISTORY RECEIVING AND STORING", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            FormExistsShow frmPMISInquiry_ReceivingHist
        Case 1335
            If Module_Access(LOGID, "MATERIAL INVENTORY MATERIALS LISTING", "TRANSACTION") = False Then Exit Sub
            PARTSQUERY = 9
            FormExistsShow frmPMISInquiry_Query
            frmPMISInquiry_Query.Caption = "MATERIALS LISTING"
            ''''''''''''''''''
        Case 1355
            If Module_Access(LOGID, "PROCESSING CHECK ERROR TRANSACTIONS", "PROCESSING") = False Then Exit Sub
            FormExistsShow frmPMISProcess_CheckDupTrans
        Case 1356
            If Module_Access(LOGID, "UPDATE MASTER FILE", "PROCESSING") = False Then Exit Sub
            FormExistsShow frmPMISProcess_UpdateMaster
        Case 1362
            If Module_Access(LOGID, "UPDATE ADJUSTMENT FILE", "PROCESSING") = False Then Exit Sub
            FormExistsShow frmPMISProcess_UpdateAdjustment
        Case 1359
            If Module_Access(LOGID, "MONTH-END PROCESSING", "PROCESSING") = False Then Exit Sub
            PROC_TYPE = "MONTH_END"
            If LOGDATE <> lastDay(LOGDATE) Then
                If MsgBox("It's Not Currently End of The Month. Are you Sure You Want To Process Month End ", vbInformation + vbYesNo, "Not Allowed!") = vbNo Then
                    Exit Sub
                End If
            End If
            FormExistsShow frmPMISProcess_MonthEndProc
            '***************************************************************************
            ''INQUIRY''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '***************************************************************************
        Case 1364
            If Module_Access(LOGID, "PARTS AVAILABILITY", "INQUIRY") = False Then Exit Sub
            FormExistsShow frmPMISInquiry_PartsInquiry
        Case 1543
            If Module_Access(LOGID, "COUNTER INQUIRY", "INQUIRY") = False Then Exit Sub
            frmPMIS_CounterInquiry_Parts.SETSTOCK_TYPE ("P")
            FormExistsShow frmPMIS_CounterInquiry_Parts
        Case 1366
            If Module_Access(LOGID, "PARTS COMPUTERIZED STOCKCARDS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 1
            FormExistsShow frmPMISInquiry_Query
        Case 1369
            If Module_Access(LOGID, "PO TRANSACTIONS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 3
            FormExistsShow frmPMISInquiry_Query
        Case 1370
            If Module_Access(LOGID, "MRR TRANSACTIONS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 4
            FormExistsShow frmPMISInquiry_Query
        Case 1371
            If Module_Access(LOGID, "TRANSACTIONS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 5
            FormExistsShow frmPMISInquiry_Query
        Case 1372
            If Module_Access(LOGID, "TRANSACTION DETAILS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 7
            FormExistsShow frmPMISInquiry_Query
        Case 1373
            If Module_Access(LOGID, "PARTS CHECK PREVIOUS BALANCE", "PROCESSING") = False Then Exit Sub
            frmPMISInquiry_CheckPrevBal_Parts.SetStockType ("P")
            FormExistsShow frmPMISInquiry_CheckPrevBal_Parts
        Case 1374
            If Module_Access(LOGID, "INVENTORY RANKING INQUIRY", "INQUIRY") = False Then Exit Sub
            FormExistsShow frmPMISInquiry_RankingInquiry
        Case 1375
            If Module_Access(LOGID, "DEALER SRP DNP LISTING", "INQUIRY") = False Then Exit Sub
            FormExistsShow frmPMISInquiry_PartsSRPComparison
        Case 1376
            If Module_Access(LOGID, "DEALER DISTRIBUTOR DNP COMPARISON", "INQUIRY") = False Then Exit Sub
            FormExistsShow frmPMISInquiry_PartsDNPComparison
        Case 1377
            If Module_Access(LOGID, "DEALER DISTRIBUTOR SRP COMPARISON", "INQUIRY") = False Then Exit Sub
            FormExistsShow frmPMISInquiry_PartsSRPComparison
        Case 1378
            If Module_Access(LOGID, "BROWSE ERROR FILES", "INQUIRY") = False Then Exit Sub
            FormExistsShow frmPMISInquiry_ErrorQuery
            '***************************************************************************
            ''REPORTS'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '***************************************************************************
        Case 1379
            If Module_Access(LOGID, "DAILY SALES REPORT", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISReports_DailySales
        Case 1381
            If Module_Access(LOGID, "PARTS TRANSACTION LISTING", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISReports_RCRange
        Case 1382
            If Module_Access(LOGID, "PARTS TRANSACTION LISTING", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISReports_ISRange
        Case 1383
            If Module_Access(LOGID, "PARTS MONTHLY REPORT", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISReports_Receipts
        Case 1380
            If Module_Access(LOGID, "RIV FOR WORKINPROGRESS", "REPORTS") = False Then Exit Sub
            ISSREPTYPE = "RIV_INPROCESS"
            FormExistsShow frmPMISReports_Issuances
        Case 1384
            If Module_Access(LOGID, "PARTS MONTHLY REPORT", "REPORTS") = False Then Exit Sub
            ISSREPTYPE = ""
            FormExistsShow frmPMISReports_Issuances
        Case 1385
            If Module_Access(LOGID, "STOCK STATUS REPORT", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISReports_PrintStockStat
        Case 1387
            If Module_Access(LOGID, "RANKING REPORT", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISReports_PrintRankfle
        Case 1388
            If Module_Access(LOGID, "REPORTS INTERNAL MOVEMENT CATEGORY", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISReports_MoveCat
        Case 1389
            If Module_Access(LOGID, "REPORTS INTERNAL STOCKS BELOW SAFETY STOCK LEVEL", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISReports_PrintBelowSafetyStock
        Case 1390
            If Module_Access(LOGID, "REPORTS INTERNAL SLOW MOVING PARTS FOR DISPOSAL", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISReports_SlowMoving
        Case 1391
            If Module_Access(LOGID, "REPORTS INTERNAL UNPOSTED RECEIPTS TRANSACTION", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISInquiry_UnPostedRCRange
        Case 1392
            If Module_Access(LOGID, "REPORTS INTERNAL UNPOSTED ISSUANCES TRANSACTION", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISInquiry_UnPostedRange
        Case 1393
            If Module_Access(LOGID, "REPORTS TOTAL RETAIL SALES", "REPORTS") = False Then Exit Sub
            PRR_REPORT = "RETAIL SALES"
            FormExistsShow frmPMISReports_PRRMonthly
        Case 1395
            If Module_Access(LOGID, "REPORTS PARTS RUNDOWN BEGINNING INVENTORY REPORT", "REPORTS") = False Then Exit Sub
            PRR_REPORT = "BEGINNING INVENTORY"
            FormExistsShow frmPMISReports_PRRMonthly
        Case 1396
            If Module_Access(LOGID, "REPORTS PARTS RUNDOWN TOTAL PURCHASES", "REPORTS") = False Then Exit Sub
            PRR_REPORT = "TOTAL PURCHASES"
            FormExistsShow frmPMISReports_PRRMonthly
        Case 1394
            If Module_Access(LOGID, "REPORTS TOTAL COST OF SALES", "REPORTS") = False Then Exit Sub
            PRR_REPORT = "COST OF SALES"
            FormExistsShow frmPMISReports_PRRMonthly
        Case 1397
            If Module_Access(LOGID, "REPORTS PARTS RUNDOWN INVENTORY ADJUSTMENTS", "REPORTS") = False Then Exit Sub
            PRR_REPORT = "INVENTORY ADJUSTMENTS"
            FormExistsShow frmPMISReports_PRRMonthly
        Case 1398
            If Module_Access(LOGID, "REPORTS PARTS RUNDOWN INVENTORY ADJUSTMENTS", "REPORTS") = False Then Exit Sub
            PRR_REPORT = "PARTS MAD"
            FormExistsShow frmPMISReports_PRRMonthly
        Case 1399
            If Module_Access(LOGID, "REPORTS PARTS RUNDOWN INVENTORY GROSS RETURN", "REPORTS") = False Then Exit Sub
            PRR_REPORT = "INV_GROSS_RETURN"
            FormExistsShow frmPMISReports_PRRMonthly
        Case 1400
            If Module_Access(LOGID, "REPORTS PARTS RUNDOWN FILL RATE", "REPORTS") = False Then Exit Sub
            PRR_REPORT = "FILL RATE"
            FormExistsShow frmPMISReports_PRRMonthly
        Case 1401
            If Module_Access(LOGID, "REPORTS PARTS RUNDOWN ORDERED PARTS REPORT BY CATEGORY", "REPORTS") = False Then Exit Sub
            PRR_REPORT = "ORDERED PARTS"
            FormExistsShow frmPMISReports_PRRMonthly
        Case 1402
            If Module_Access(LOGID, "REPORTS PARTS RUNDOWN PARTS BACK ORDER", "REPORTS") = False Then Exit Sub
            PRR_REPORT = "PARTS BACK ORDER"
            FormExistsShow frmPMISReports_PRRMonthly
        Case 1403
            If Module_Access(LOGID, "REPORTS PARTS RUNDOWN EXCEL", "REPORTS") = False Then Exit Sub
        Case 1352
            If Module_Access(LOGID, "MATERIAL INVENTORY MONTHLY BIR YEAR REPORT", "REPORTS") = False Then Exit Sub
            BIR_YearEnd = "MATERIALS"
            FormExistsShow frmPMISReports_BIR_YearEnd
        Case 1404
            If Module_Access(LOGID, "REPORT GOV BIR YEAR REPORT", "REPORTS") = False Then Exit Sub
            BIR_YearEnd = "PARTS"
            FormExistsShow frmPMISReports_BIR_YearEnd
            '***************************************************************************
        Case 1407
            FormExistsShow frmAccMaintenance
        Case 1410
            FormExistsShow frmAbout
        Case 1412
            Unload Me
        Case 1471
            frmMainMenu.Show
            frmMainMenu.ZOrder 0
            '================================
            'Accessories Transaction History
            '================================
        Case 1499                                     'Accessories Cash Counter Issuance
            If Module_Access(LOGID, "TRANSACTION HISTORY CASH COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist_AC
            COUNTERTYPE = "CSH"
            frmPMISInquiry_CustomerOrderHist_AC.txtTranType.Text = "CSH"
            FormExistsShow frmPMISInquiry_CustomerOrderHist_AC
        Case 1500                                     'Accessories Charge Counter Issuance
            If Module_Access(LOGID, "TRANSACTION HISTORY CHARGE COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist_AC
            COUNTERTYPE = "CHG"
            frmPMISInquiry_CustomerOrderHist_AC.txtTranType.Text = "CHG"
            FormExistsShow frmPMISInquiry_CustomerOrderHist_AC
        Case 1501                                     'Requisition Issuance Voucher
            If Module_Access(LOGID, "TRANSACTION HISTORY REQUISTION ISSUANCE VOUCHER", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist_AC
            COUNTERTYPE = "RIV"
            frmPMISInquiry_CustomerOrderHist_AC.txtTranType.Text = "RIV"
            FormExistsShow frmPMISInquiry_CustomerOrderHist_AC
        Case 1502                                     'DR Out Issuance
            If Module_Access(LOGID, "TRANSACTION HISTORY DR OUT ISSUANCE", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist_AC
            COUNTERTYPE = "DR"
            frmPMISInquiry_CustomerOrderHist_AC.txtTranType.Text = "DR"
            FormExistsShow frmPMISInquiry_CustomerOrderHist_AC
        Case 1503                                     'Accessories Advance Bill Data Entry
            If Module_Access(LOGID, "TRANSACTION HISTORY ADVANCE BILL DATA ENTRY", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist_AC
            COUNTERTYPE = "ADB"
            frmPMISInquiry_CustomerOrderHist_AC.txtTranType.Text = "ADB"
            FormExistsShow frmPMISInquiry_CustomerOrderHist_AC
        Case 1504                                     'Accessories Receiving and Storing
            If Module_Access(LOGID, "ACCESSORIES TRANSACTION HISTORY RECEIVING AND STORING", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            FormExistsShow frmPMISAC_ReceivingHist
            '================================
            'Materials Transaction History
            '================================
        Case 1507                                     'Materials Cash Counter Issuance
            If Module_Access(LOGID, "TRANSACTION HISTORY CASH COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist_MAT
            COUNTERTYPE = "CSH"
            frmPMISInquiry_CustomerOrderHist_MAT.txtTranType.Text = "CSH"
            FormExistsShow frmPMISInquiry_CustomerOrderHist_MAT
        Case 1508                                     'Materials Charge Counter Issuance
            If Module_Access(LOGID, "TRANSACTION HISTORY CHARGE COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist_MAT
            COUNTERTYPE = "CHG"
            frmPMISInquiry_CustomerOrderHist_MAT.txtTranType.Text = "CHG"
            FormExistsShow frmPMISInquiry_CustomerOrderHist_MAT
        Case 1509                                     'Materials Requisition Issuance Voucher
            If Module_Access(LOGID, "TRANSACTION HISTORY REQUISTION ISSUANCE VOUCHER", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist_MAT
            COUNTERTYPE = "RIV"
            frmPMISInquiry_CustomerOrderHist_MAT.txtTranType.Text = "RIV"
            FormExistsShow frmPMISInquiry_CustomerOrderHist_MAT
        Case 1510                                     'Materials DR Out Issuance
            If Module_Access(LOGID, "TRANSACTION HISTORY DR OUT ISSUANCE", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist_MAT
            COUNTERTYPE = "DR"
            frmPMISInquiry_CustomerOrderHist_MAT.txtTranType.Text = "DR"
            FormExistsShow frmPMISInquiry_CustomerOrderHist_MAT
        Case 1555                                     'Materials Advance Bill Data Entry
            If Module_Access(LOGID, "TRANSACTION HISTORY ADVANCE BILL DATA ENTRY", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_CustomerOrderHist_MAT
            COUNTERTYPE = "ADB"
            frmPMISInquiry_CustomerOrderHist_MAT.txtTranType.Text = "ADB"
            FormExistsShow frmPMISInquiry_CustomerOrderHist_MAT
        Case 1512                                     'Materials Receiving and Storing
            If Module_Access(LOGID, "MATERIALS TRANSACTION HISTORY RECEIVING AND STORING", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            FormExistsShow frmPMISMat_ReceivingHist
        Case 1513                                     'Order Report from HARI
            If Module_Access(LOGID, "ORDER REPORT", "REPORTS") = False Then Exit Sub
            ORDER_REPORT = "HARI"
            FormExistsShow frmPMISReports_OrderReport
        Case 1514                                     'Order Report from Other Supplier
            If Module_Access(LOGID, "ORDER REPORT", "REPORTS") = False Then Exit Sub
            ORDER_REPORT = "NON_HARI"
            FormExistsShow frmPMISReports_OrderReport
        Case 1515                                     'Lost Sales
            frmPMISReports_LostSales.Show
        Case 1516                                     'Stock Report: History of Price of Parts(Cost, SRP)
            FormExistsShow frmPMISReports_History_DNP_SRP
        Case 1517                                     'Ordered Report by Category
            FormExistsShow frmPMISReports_OrderedReport_ByCategory
        Case 1519                                     'Sales Report: Local Purchase, Imported and Consigned
            FormExistsShow frmPMISReports_SalesReport_Loc_Imp_Con
            'Add FORECASTING REPORT in Menu Bar
        Case 1521                                     'Level Of Service Report
            If Module_Access(LOGID, "FORCASTING REPORT", "REPORTS") = False Then Exit Sub
            FORECASTING_BUTTON_CLICK = 1
            FormExistsShow frmPMISReports_PrintForeCasting
        Case 1522                                     'Six Months Moving Average Report
            If Module_Access(LOGID, "FORCASTING REPORT", "REPORTS") = False Then Exit Sub
            FORECASTING_BUTTON_CLICK = 2
            FormExistsShow frmPMISReports_PrintForeCasting
        Case 1523                                     'Moving Median Reports(6 Mos.)
            If Module_Access(LOGID, "FORCASTING REPORT", "REPORTS") = False Then Exit Sub
            FORECASTING_BUTTON_CLICK = 3
            FormExistsShow frmPMISReports_PrintForeCasting
        Case 1524                                     'Linear Regression Report
            If Module_Access(LOGID, "FORCASTING REPORT", "REPORTS") = False Then Exit Sub
            FORECASTING_BUTTON_CLICK = 4
            FormExistsShow frmPMISReports_PrintForeCasting
        Case 1525                                     'Mean Absolute Deviation Report
            If Module_Access(LOGID, "FORCASTING REPORT", "REPORTS") = False Then Exit Sub
            FORECASTING_BUTTON_CLICK = 5
            FormExistsShow frmPMISReports_PrintForeCasting
        Case 1526                                     'Safety Stock Report
            If Module_Access(LOGID, "FORCASTING REPORT", "REPORTS") = False Then Exit Sub
            FORECASTING_BUTTON_CLICK = 6
            FormExistsShow frmPMISReports_PrintForeCasting
        Case 1527                                     'Suggested Order Quantity Report
            If Module_Access(LOGID, "FORCASTING REPORT", "REPORTS") = False Then Exit Sub
            FORECASTING_BUTTON_CLICK = 7
            FormExistsShow frmPMISReports_PrintForeCasting
        Case 1528
            If Module_Access(LOGID, "PARTS TRANSACTION LISTING", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISReports_Parts_PORange
        Case 1531
            If Module_Access(LOGID, "PARTS RUNDOWN REPORT", "REPORTS") = False Then Exit Sub
            PRR_BUTTON_CLICK = 1                      'PRR - Retail Sales
            FormExistsShow frmPMISReports_PrintPartsRunDown
        Case 1532
            If Module_Access(LOGID, "PARTS RUNDOWN REPORT", "REPORTS") = False Then Exit Sub
            PRR_BUTTON_CLICK = 2                      'PRR - Cost of Sales
            FormExistsShow frmPMISReports_PrintPartsRunDown
        Case 1533
            If Module_Access(LOGID, "PARTS RUNDOWN REPORT", "REPORTS") = False Then Exit Sub
            PRR_BUTTON_CLICK = 3                      'PRR - Beginning Inventory Report
            FormExistsShow frmPMISReports_PrintPartsRunDown
        Case 1534
            If Module_Access(LOGID, "PARTS RUNDOWN REPORT", "REPORTS") = False Then Exit Sub
            PRR_BUTTON_CLICK = 4                      'PRR - Total Purchases Report
            FormExistsShow frmPMISReports_PrintPartsRunDown
        Case 1535
            If Module_Access(LOGID, "PARTS RUNDOWN REPORT", "REPORTS") = False Then Exit Sub
            PRR_BUTTON_CLICK = 5                      'PRR - Inventory Adjustments
            FormExistsShow frmPMISReports_PrintPartsRunDown
        Case 1536
            If Module_Access(LOGID, "PARTS RUNDOWN REPORT", "REPORTS") = False Then Exit Sub
            PRR_BUTTON_CLICK = 6                      'PRR - Parts Moving Average Demand
            FormExistsShow frmPMISReports_PrintPartsRunDown
        Case 1537
            If Module_Access(LOGID, "PARTS RUNDOWN REPORT", "REPORTS") = False Then Exit Sub
            PRR_BUTTON_CLICK = 7                      'PRR - Inventory Gross Return
            FormExistsShow frmPMISReports_PrintPartsRunDown
        Case 1538
            If Module_Access(LOGID, "PARTS RUNDOWN REPORT", "REPORTS") = False Then Exit Sub
            PRR_BUTTON_CLICK = 8                      'PRR - Fill Rate Reports
            FormExistsShow frmPMISReports_PrintPartsRunDown
        Case 1539
            If Module_Access(LOGID, "PARTS RUNDOWN REPORT", "REPORTS") = False Then Exit Sub
            PRR_BUTTON_CLICK = 9                      'PRR - Back-Order Report
            FormExistsShow frmPMISReports_PrintPartsRunDown
        Case 1540
            If Module_Access(LOGID, "PARTS RUNDOWN REPORT", "REPORTS") = False Then Exit Sub
            PRR_BUTTON_CLICK = 10                     'PRR - Excel Reports
            FormExistsShow frmPMISReports_PrintPartsRunDown
        Case 1541                                     'Purchases for the Month - Parts
            If Module_Access(LOGID, "PURCHASE FOR THE MONTH", "REPORTS") = False Then Exit Sub
            FormExistsShow frmPMISReports_Purchase_For_The_Month

        Case 1544                                     'Accessories Stock Cards
            If Module_Access(LOGID, "ACCESSORIES LEDGER FILE", "INQUIRY") = False Then Exit Sub
            Unload frmPMISInquiry_Query
            PARTSQUERY = 1
            frmPMISInquiry_Query.SetTYPE ("A")
            FormExistsShow frmPMISInquiry_Query
        Case 1545                                     'Materials Stock Cards
            If Module_Access(LOGID, "MATERIALS QUERY", "INQUIRY") = False Then Exit Sub
            Unload frmPMISInquiry_Query
            PARTSQUERY = 1
            frmPMISInquiry_Query.SetTYPE ("M")
            FormExistsShow frmPMISInquiry_Query
        Case 1546                                     'Accessories Counter Inquiry
            If Module_Access(LOGID, "ACCESSORIES COUNTER INQUIRY", "INQUIRY") = False Then Exit Sub
            frmPMIS_CounterInquiry_Accessories.SETSTOCK_TYPE ("A")
            FormExistsShow frmPMIS_CounterInquiry_Accessories
        Case 1548                                     'Accessories Running Balance
            If Module_Access(LOGID, "ACCESSORIES CHECK PREVIOUS BALANCE", "PROCESSING") = False Then Exit Sub
            frmPMISInquiry_CheckPrevBal_Accessories.SetStockType ("A")
            FormExistsShow frmPMISInquiry_CheckPrevBal_Accessories
        Case 1549                                     'Materials Running Balance
            If Module_Access(LOGID, "MATERIALS CHECK PREVIOUS BALANCE", "PROCESSING") = False Then Exit Sub
            frmPMISInquiry_CheckPrevBal_Materials.SetStockType ("M")
            FormExistsShow frmPMISInquiry_CheckPrevBal_Materials
        Case 1550                                     'Materials Running Balance
            FormExistsShow frmSMIS_Log_Reminder
        Case 1552                                     'Parts - Purchase Order History
            FormExistsShow frmPMISInquiry_Purchase_Hist
        Case 1553                                     'Accessories - Purchase Order History
            FormExistsShow frmPMISAC_Purchase_Hist
        Case 1554                                     'Materials - Purchase Order History
            FormExistsShow frmPMISMAT_Purchase_Hist
        Case 1557
            If Module_Access(LOGID, "DEALER PARTS INQUIRY", "INQUIRY") = False Then Exit Sub
            FormExistsShow frmPMISReports_DealerPartInquiry    'Dealer Part Inquiry Report
        Case 1558
            If Module_Access(LOGID, "PARTS QUALITY INFORMATION REPORT", "DATA ENTRY") = False Then Exit Sub
            FormExistsShow frmPMISReports_PQIRReport  'Parts Quality Information Report
        Case 1564
            If Module_Access(LOGID, "PARTS ISSUED TO CUSTOMER", "REPORTS") = False Then Exit Sub
            PARTS_ISSUED_TO_CUSTOMER_TYPE = "P"
            FormExistsShow frmPMISReports_PartsIssuedToCustomer
        Case 1565
            If Module_Access(LOGID, "PARTS ISSUED TO CUSTOMER", "REPORTS") = False Then Exit Sub
            PARTS_ISSUED_TO_CUSTOMER_TYPE = "A"
            FormExistsShow frmPMISReports_PartsIssuedToCustomer
        Case 1566
            If Module_Access(LOGID, "PARTS ISSUED TO CUSTOMER", "REPORTS") = False Then Exit Sub
            PARTS_ISSUED_TO_CUSTOMER_TYPE = "M"
            FormExistsShow frmPMISReports_PartsIssuedToCustomer
        Case 1567
            If Module_Access(LOGID, "BACK-ORDER ALLOCATION", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISTrans_POConfirmationProcess
        Case 1547                                     '&Materials Counter Inquiry
            If Module_Access(LOGID, "MATERIALS COUNTER INQUIRY", "INQUIRY") = False Then Exit Sub
            frmPMIS_CounterInquiry_Materials.SETSTOCK_TYPE ("M")
            FormExistsShow frmPMIS_CounterInquiry_Materials
        Case 1568                                     'A&ccessories PO Transactions
            If Module_Access(LOGID, "ACCESSORIES PO TRANSACTIONS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 3
            frmPMISInquiry_Query.SetTYPE ("A")
            FormExistsShow frmPMISInquiry_Query
        Case 1569                                     'Acc&essories MRR Transactions

            If Module_Access(LOGID, "ACCESSORIES MRR TRANSACTIONS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 4
            frmPMISInquiry_Query.SetTYPE ("A")
            FormExistsShow frmPMISInquiry_Query
        Case 1570                                     'Accessor&ies Issuances Transactions


            If Module_Access(LOGID, "ACCESSORIES TRANSACTIONS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 5
            frmPMISInquiry_Query.SetTYPE ("A")
            FormExistsShow frmPMISInquiry_Query

        Case 1571                                     'Access&ories Transaction Details


            If Module_Access(LOGID, "ACCESSORIES TRANSACTION DETAILS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 7
            frmPMISInquiry_Query.SetTYPE ("A")
            FormExistsShow frmPMISInquiry_Query


        Case 1572                                     'Materials PO Tra&nsactions

            If Module_Access(LOGID, "MATERIAL PO TRANSACTIONS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 3
            frmPMISInquiry_Query.SetTYPE ("P")
            FormExistsShow frmPMISInquiry_Query


        Case 1573                                     'Material&s MRR Transactions



            If Module_Access(LOGID, "MATERIAL MRR TRANSACTIONS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 4
            frmPMISInquiry_Query.SetTYPE ("M")
            FormExistsShow frmPMISInquiry_Query

        Case 1574                                     'Materi&als Issuances Transactions
            If Module_Access(LOGID, "MATERIAL TRANSACTIONS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 5
            frmPMISInquiry_Query.SetTYPE ("M")
            FormExistsShow frmPMISInquiry_Query
        Case 1575                                     'Materia&ls Transaction Details


            If Module_Access(LOGID, "MATERIAL TRANSACTION DETAILS", "INQUIRY") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISInquiry_Query
            PARTSQUERY = 7
            frmPMISInquiry_Query.SetTYPE ("M")
            FormExistsShow frmPMISInquiry_Query
        Case 1576                                     'Accessories Inventory Adjustment
            If Module_Access(LOGID, "ACCESSORIES INVENTORY ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
            frmPMISTrans_InventoryAdjustment_Accessories.SetStockType ("A")
            FormExistsShow frmPMISTrans_InventoryAdjustment_Accessories
        Case 1577                                     'Materials Inventory Adjustment
            If Module_Access(LOGID, "MATERIALS INVENTORY ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
            frmPMISTrans_InventoryAdjustment_Materials.SetStockType ("M")
            FormExistsShow frmPMISTrans_InventoryAdjustment_Materials
        Case 1405                                     'Company Profile
            If Module_Access(LOGID, "COMPANY PROFILE", "DATA ENTRY") = False Then Exit Sub
            frmPMISProfile.Show
        Case Else
Debug.Print " CASE " & Control.ID & " '" & Control.Caption


    End Select
End Sub


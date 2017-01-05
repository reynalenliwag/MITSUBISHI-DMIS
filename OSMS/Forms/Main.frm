VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "WIZENCRYPT.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#10.4#0"; "CO15D0~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#10.4#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000F&
   Caption         =   "Office Supplies Management System"
   ClientHeight    =   7020
   ClientLeft      =   165
   ClientTop       =   270
   ClientWidth     =   11880
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Main.frx":000C
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6840
      Top             =   4080
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   4350
      Top             =   4020
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin wizEncrypt.wizEnc wizEnc1 
      Left            =   180
      Top             =   2760
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
   Begin Crystal.CrystalReport rptMain 
      Left            =   6330
      Top             =   4020
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   5676
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":577D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":57AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":57E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":58124
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5843E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":58758
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":58A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":58D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":591DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":59630
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5994A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":59D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5A1EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5A508
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6765
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
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
            TextSave        =   "10:50 PM"
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
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   5234
      Top             =   4020
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   6
      DesignerControls=   "Main.frx":5A822
   End
   Begin XtremeSuiteControls.PopupControl PopCntrl 
      Left            =   4792
      Top             =   4020
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'UPdate by AXP-062620071225
'Function Feature   : Reminder Module
'Date               : 06/26/2007
'Last Update        : 06/26/2007
'Database Update    : Added Table For Reminder Called Cris Reminders
'Who Updated        : AXP
'Upating Code       :AXP-062620071225
Private Sub MDIForm_Load()
    ApplyThemes
    ConfigurePopUps

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
      If MsgBox("Exit OSMS, Are You Sure?", vbExclamation + vbOKCancel, "Exit System") = vbOK Then
        Dim frm                              As Form
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

    Select Case Control.Id
    Case FILE_DEPARTMENT
        Screen.MousePointer = 11
        frmOSMSFilesDepartment.Show
        frmOSMSFilesDepartment.ZOrder 0
        Screen.MousePointer = 0
    Case FILE_EMPLOYEE
        frmOSMSFilesEmployee.Show
        frmOSMSFilesEmployee.ZOrder 0
    Case FILE_SIGNATORIES
        frmOSMSFilesSignatories.Show
        frmOSMSFilesSignatories.ZOrder 0
    Case FILE_SUPPLIER
        frmOSMSFilesSupplier.Show
        frmOSMSFilesSupplier.ZOrder 0
    Case FILE_SUPPLY
        frmOSMSFilesSupply.Show
        frmOSMSFilesSupply.ZOrder 0
    Case FILE_UNIT
        frmOSMSFilesUnit.Show
        frmOSMSFilesUnit.ZOrder 0
    Case INQUIRY_SUPPLIESISSUED
        frmOSMSInquiryIssued.Show
        frmOSMSInquiryIssued.ZOrder 0
    Case INQUIRY_SUPPLIESRECEIVED
        frmOSMSInquiryReceiving.Show
        frmOSMSInquiryReceiving.ZOrder 0
    Case INQUIRY_SUPPLYINVENTORY
        frmOSMSInquirySupply.Show
        frmOSMSInquirySupply.ZOrder 0

    Case MAINTAINENANCE_PASSWORDMAINTENANCE
        frmAccMaintenance.Show
        frmAccMaintenance.ZOrder 0
    Case PROCEESING_MONTHEND_BATCHPOSTING
        frmOSMSProcessBatchPosting.Show
    Case PROCESSING_CHECKPREVIOUSBALANCE
        frmOSMSProcessCheckPrevBal.Show
    Case PROCESSING_MONTHEND_CREATESTOCKSTATUS
        PROC_TYPE = "STKSTAT"
        frmOSMSProcessMonthEndProc.Show
    Case PROCESSING_MONTHEND_GENERATERANKFILE
        PROC_TYPE = "RANKING"
        frmOSMSProcessMonthEndProc.Show
    Case PROCESSING_MONTHEND_MONTHPROCESS
        PROC_TYPE = "MONTH_END"
        frmOSMSProcessMonthEndProc.Show
    Case PROCESSING_UPDATEINVENTORYADJUSTMENTS
        frmOSMSProcessUpdateAdjustment.Show
    Case PROCESSING_UPDATESUPPLIESMASTERFILE
        frmOSMSProcessUpdateMaster.Show

    Case REPORTS_MONTHLYISSUANCEBYDEPARTMENT
        frmOSMSReportDepartment.Show
        frmOSMSReportDepartment.ZOrder 0
    Case REPORTS_MONTHLYRECEIPTSBYSUPPLIER
        frmOSMSReportSupplier.Show
        frmOSMSReportDepartment.ZOrder 0
    Case REPORTS_SUPPLYINVENTORYRANKING

    Case REPORTS_SUPPLYSTATUSREPORT

    Case TRANSACTION_RECEIVINGSUPPLY
        frmOSMSTransactionReceivingSupply.Show
        frmOSMSTransactionReceivingSupply.ZOrder 0
    Case TRANSACTION_SUPPLIESINVENTORYADJUSTMENT
        frmOSMSTransactionInvAdjustment.Show
        frmOSMSTransactionInvAdjustment.ZOrder 0
    Case TRANSACTION_SUPPLYISSUANCE
        frmOSMSTransactionIssuance.Show
        frmOSMSTransactionIssuance.ZOrder 0
    Case WINDOW_ABOUT
        frmAbout.Show
        frmAbout.ZOrder 0
    Case WINDOW_EXITOSMS
        If MsgBoxXP("Are you sure you want to exit OSMS?", "Exit OSMS", XP_YesNo, msg_Question) = True Then
            End
        End If
    
    Case 1132 'Dashboard
        frmMainMenu.Show
'    Case Else
'        Debug.Print Control.Id; "'"; Control.Caption
'        Stop
    
    End Select

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
        .LoadSkin SKIN_PATH, ""
        .ApplyWindow Me.hwnd
        .ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or Not xtpSkinApplyMetrics
        '.ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or xtpSkinApplyMetrics
    End With
    Dim ToolTipContext As ToolTipContext
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
    Dim Item As PopupControlItem
    PopCntrl.RemoveAllItems
    PopCntrl.Icons.AddIcons CommandBars1.Icons
    PopCntrl.VisualTheme = xtpPopupThemeOffice2003
    PopCntrl.SetSize 270, 140

    Set Item = PopCntrl.AddItem(245, 8, 265, 20, vbNullString)
    Item.Button = True
    Item.IconIndex = 899
    Item.Id = 707
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
    Item.Id = 655
    Item.Hyperlink = False

    Set Item = PopCntrl.AddItem(20, 85, 260, 105, vbNullString)
    Item.TextAlignment = DT_WORDBREAK Or DT_LEFT
    Item.TextColor = RGB(190, 1, 1)
    Item.Height = 50
    Item.Font.Size = 7
    Item.Hyperlink = False
End Sub
Private Sub PopCntrl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.Id = 707 Then
        PopCntrl.Close
    End If

End Sub

Private Sub Timer1_Timer()
    If TIMER_REMIND = "" Then
        ReminderModule ""
    Else
        If DateDiff("n", TIMER_REMIND, Now) >= 0 Then
            frmSMIS_Files_Reminders.Show
            
        End If
    End If
End Sub

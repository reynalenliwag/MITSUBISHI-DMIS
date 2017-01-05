VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COC288~1.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#10.4#0"; "CO15D0~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#10.4#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DMIS 2.0 System Administration"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "mainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picClient 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   5145
      Left            =   30
      ScaleHeight     =   5145
      ScaleWidth      =   9720
      TabIndex        =   0
      Top             =   540
      Width           =   9720
      Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
         Height          =   5265
         Left            =   9960
         TabIndex        =   1
         Top             =   120
         Width           =   105
         _Version        =   655364
         _ExtentX        =   185
         _ExtentY        =   9287
         _StockProps     =   64
         VisualTheme     =   3
         Animation       =   1
         ItemLayout      =   2
         HotTrackStyle   =   1
         ColumnWidth     =   50
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DATABASE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5790
         MouseIcon       =   "mainMenu.frx":3332
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Tag             =   "MODULE"
         Top             =   4380
         Width           =   1170
      End
      Begin VB.Label LABDATABASE 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6990
         MouseIcon       =   "mainMenu.frx":3484
         TabIndex        =   13
         Tag             =   "MODULE"
         Top             =   4380
         Width           =   3045
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SERVER:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         MouseIcon       =   "mainMenu.frx":35D6
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Tag             =   "MODULE"
         Top             =   4380
         Width           =   1170
      End
      Begin VB.Label LABSERVER 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         MouseIcon       =   "mainMenu.frx":3728
         TabIndex        =   11
         Tag             =   "MODULE"
         Top             =   4380
         Width           =   2985
      End
      Begin VB.Label lblCopyRightAccess 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copy Right Access Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1530
         MouseIcon       =   "mainMenu.frx":387A
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Tag             =   "COPYUSERSETTING"
         Top             =   1710
         Width           =   2565
      End
      Begin VB.Label lblAuditInquiry 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audit Inquiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5640
         MouseIcon       =   "mainMenu.frx":39CC
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Tag             =   "AUDITINQUIRY"
         Top             =   2310
         Width           =   1215
      End
      Begin VB.Label lblAuditReport 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audit Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1530
         MouseIcon       =   "mainMenu.frx":3B1E
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Tag             =   "AUDITREPORT"
         Top             =   2310
         Width           =   1185
      End
      Begin VB.Label lblImportModules 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Import Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5640
         MouseIcon       =   "mainMenu.frx":3C70
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Tag             =   "IMPORTMODULE"
         Top             =   2940
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblUserRightAccessReport 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Right Acess Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5640
         MouseIcon       =   "mainMenu.frx":3DC2
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Tag             =   "RIGHTACESSREPORT"
         Top             =   1710
         Width           =   2265
      End
      Begin VB.Label lblServerSetting 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5640
         MouseIcon       =   "mainMenu.frx":3F14
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Tag             =   "SERVERSETTING"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblUserModule 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Modules"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1530
         MouseIcon       =   "mainMenu.frx":4066
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Tag             =   "USERMODULE"
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label lblModule 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modules"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1530
         MouseIcon       =   "mainMenu.frx":41B8
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Tag             =   "MODULE"
         Top             =   420
         Width           =   795
      End
      Begin VB.Label lblUserMaintain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Maintenance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5640
         MouseIcon       =   "mainMenu.frx":430A
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Tag             =   "USERMAINTAIN"
         Top             =   420
         Width           =   1710
      End
      Begin VB.Image imgCopyRightAccess 
         Height          =   480
         Left            =   930
         MouseIcon       =   "mainMenu.frx":445C
         MousePointer    =   99  'Custom
         Picture         =   "mainMenu.frx":45AE
         Top             =   1590
         Width           =   480
      End
      Begin VB.Image imgAuditInquiry 
         Height          =   480
         Left            =   5010
         MouseIcon       =   "mainMenu.frx":4C0F
         MousePointer    =   99  'Custom
         Picture         =   "mainMenu.frx":4D61
         Top             =   2190
         Width           =   480
      End
      Begin VB.Image imgAuditReport 
         Height          =   480
         Left            =   960
         MouseIcon       =   "mainMenu.frx":5409
         MousePointer    =   99  'Custom
         Picture         =   "mainMenu.frx":555B
         Top             =   2190
         Width           =   480
      End
      Begin VB.Image imgImportMoudle 
         Height          =   480
         Left            =   5040
         MouseIcon       =   "mainMenu.frx":5B03
         MousePointer    =   99  'Custom
         Picture         =   "mainMenu.frx":5C55
         Top             =   2820
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgUserRightAcessReport 
         Height          =   480
         Left            =   5010
         MouseIcon       =   "mainMenu.frx":62AA
         MousePointer    =   99  'Custom
         Picture         =   "mainMenu.frx":63FC
         Top             =   1590
         Width           =   480
      End
      Begin VB.Image imgServerSettings 
         Height          =   480
         Left            =   5010
         MouseIcon       =   "mainMenu.frx":6A80
         MousePointer    =   99  'Custom
         Picture         =   "mainMenu.frx":6BD2
         Top             =   960
         Width           =   480
      End
      Begin VB.Image imgUserModule 
         Height          =   480
         Left            =   930
         MouseIcon       =   "mainMenu.frx":726D
         MousePointer    =   99  'Custom
         Picture         =   "mainMenu.frx":73BF
         Top             =   960
         Width           =   480
      End
      Begin VB.Image imModules 
         Height          =   480
         Left            =   930
         MouseIcon       =   "mainMenu.frx":7A3C
         MousePointer    =   99  'Custom
         Picture         =   "mainMenu.frx":7B8E
         Top             =   300
         Width           =   480
      End
      Begin VB.Image imgUserMaintain 
         Height          =   480
         Left            =   5010
         MouseIcon       =   "mainMenu.frx":81B5
         MousePointer    =   99  'Custom
         Picture         =   "mainMenu.frx":8307
         Top             =   420
         Width           =   480
      End
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   180
      Top             =   -1200
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   2370
      Top             =   -1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMenu.frx":89BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMenu.frx":8FF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMenu.frx":9667
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMenu.frx":9CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMenu.frx":A38F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMenu.frx":AA00
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMenu.frx":B0AB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   1800
      Top             =   -1170
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "mainMenu.frx":B763
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   1110
      Top             =   -1260
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.PopupControl PopCntrl 
      Bindings        =   "mainMenu.frx":4D963
      Left            =   615
      Top             =   -1200
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

Private Sub ApplyThemes()
    SKIN_PATH = App.Path & "\STYLES\THEME.cjstyles"
    With CommandBars1
        .EnableOffice2007Frame True
        .PaintManager.ClearTypeTextQuality = True
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
    End With

End Sub

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case 1201                                            '&New User
            frmfiles_Users.Show

            frmfiles_Users.cmdAdd.Value = True
        Case 1203                                            'Preview & Print Audit Log Report
            frmReportAuditReport.Show
        Case 1204                                            'Preview & Print User Right Acess Form

        Case 1205                                            'Preview and Print Users Right Acess Report...
            frmReportUserAcess.Show
        Case 1207                                            'Modules

            frmFiles_Modules.Show
        Case 1208                                            'User Modules

            frmFiles_AcessManagement.Show
        Case 1209                                            'Browser Modules
            '  frmModulesSheet.Show
        Case 1210                                            'Import Modules
            frmImportModules.Show
        Case 1211                                            'Users Maintainenace
            frmfiles_Users.Show
        Case 1212                                            'Audit Inquiry
            frmInquiry_Audit.Show
        Case 1213                                            'Server Settings
            frmFiles_ServerSetting.intSteps = 0
            frmFiles_ServerSetting.ShowLogin = False
            frmFiles_ServerSetting.Show
        Case 1214                                            'Exit
            Unload Me
    End Select
End Sub

Private Sub CommandBars1_Resize()
    On Error Resume Next
    Dim Left                            As Long
    Dim Top                             As Long
    Dim Right                           As Long
    Dim Bottom                          As Long
    CommandBars1.GetClientRect Left, Top, Right, Bottom
    picClient.Move Left, Top, Right - Left, Bottom - Top
    wndTaskPanel.Height = picClient.ScaleHeight
End Sub

''''''''''''''START REGION POPUPCONTROLS''''''''''''''''''''''''''''''''''''
Private Sub ConfigurePopUps()
    Dim item                            As PopupControlItem
    PopCntrl.RemoveAllItems
    PopCntrl.Icons.AddIcons ImageManager1.Icons
    PopCntrl.VisualTheme = xtpPopupThemeOffice2003
    PopCntrl.SetSize 270, 140

    Set item = PopCntrl.AddItem(245, 8, 265, 20, vbNullString)
    item.Button = True
    item.IconIndex = 899
    item.Id = 707
    item.Height = 20
    item.Width = 20
    item.CenterIcon
    Set item = PopCntrl.AddItem(10, 10, 218, 30, vbNullString)
    item.TextColor = RGB(15, 48, 145)
    item.Bold = True
    item.Font.Size = 10
    item.Hyperlink = False
    Set item = PopCntrl.AddItem(10, 32, 60, 50, vbNullString)
    item.Height = 50
    item.Width = 50
    item.IconIndex = 0
    item.Hyperlink = False

    Set item = PopCntrl.AddItem(62, 32, 260, 50, vbNullString)
    item.TextAlignment = DT_WORDBREAK Or DT_LEFT
    item.Height = 50
    item.Id = 655
    item.Hyperlink = False

    Set item = PopCntrl.AddItem(20, 85, 260, 105, vbNullString)
    item.TextAlignment = DT_WORDBREAK Or DT_LEFT
    item.TextColor = RGB(190, 1, 1)
    item.Height = 50
    item.Font.Size = 7
    item.Hyperlink = False
End Sub

Private Sub CreateRibbonBar()
    Dim Control                         As CommandBarControl
    Dim RibbonBar                       As RibbonBar
    Set RibbonBar = CommandBars1.AddRibbonBar("The Ribbon")
    RibbonBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    Dim ControlFile As CommandBarPopup, ControlTheme As CommandBarPopup, ControlAbout As CommandBarControl
    Dim PopupBar                        As CommandBar
    Set ControlFile = RibbonBar.AddSystemButton()
    ControlFile.IconId = 1200
    ControlFile.Caption = "&File"
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, 1201, "&New User"
        Set Control = .Add(xtpControlSplitButtonPopup, 1202, "&Print")
        Control.BeginGroup = True
        Set PopupBar = CommandBars1.CreateCommandBar("CXTPRibbonSystemPopupBarPage")
        Set Control.CommandBar = PopupBar
        Set Control = PopupBar.Controls.Add(xtpControlLabel, 0, "Report & Forms")
        Control.Width = 296
        Control.DefaultItem = True
        Control.Style = xtpButtonCaption
        PopupBar.DefaultButtonStyle = xtpButtonCaptionAndDescription
        PopupBar.SetIconSize 32, 32
        PopupBar.ShowGripper = False
        PopupBar.Controls.Add xtpControlButton, 1203, "Preview & Print Audit Log Report"
        PopupBar.Controls.Add xtpControlButton, 1204, "Preview & Print User Right Acess Form"
        PopupBar.Controls.Add xtpControlButton, 1205, "Preview and Print Users Right Acess Report..."
        Set Control = .Add(xtpControlSplitButtonPopup, 1206, "&Modules")
        Control.BeginGroup = True
        Set PopupBar = CommandBars1.CreateCommandBar("CXTPRibbonSystemPopupBarPage")
        Set Control.CommandBar = PopupBar
        Set Control = PopupBar.Controls.Add(xtpControlLabel, 0, "Modules")
        Control.Width = 296
        Control.DefaultItem = True
        Control.Style = xtpButtonCaption
        PopupBar.Controls.Add xtpControlButton, 1207, "Modules"
        PopupBar.Controls.Add xtpControlButton, 1208, "User Modules"
        PopupBar.Controls.Add xtpControlButton, 1209, "Browser Modules"
        PopupBar.Controls.Add xtpControlButton, 1210, "Import Modules"
        PopupBar.DefaultButtonStyle = xtpButtonCaptionAndDescription
        PopupBar.SetIconSize 32, 32
        PopupBar.ShowGripper = False
        .Add xtpControlButton, 1211, "Users Maintainenace"
        .Add xtpControlButton, 1212, "Audit Inquiry"
        .Add xtpControlButton, 1213, "Server Settings"
        Set Control = CommandBars1.CreateCommandBarControl("CXTPRibbonControlSystemPopupBarButton")
        Control.Caption = "Exit"
        Control.Id = 1214
        .AddControl Control
        ControlFile.CommandBar.SetIconSize 32, 32
    End With
End Sub

Sub CreateTaskPanel()
    Dim Group                           As TaskPanelGroup
    Set Group = wndTaskPanel.Groups.Add(0, "Modules")
    Group.Tooltip = "Main Modules"
    Group.Special = True
    Group.Items.Add 9001, "Modules", xtpTaskItemTypeLink, 1
    Group.Items.Add 9002, "Browser Modules", xtpTaskItemTypeLink, 2
    Group.Items.Add 9003, "Import Modules", xtpTaskItemTypeLink, 3
    Set Group = wndTaskPanel.Groups.Add(0, "Settings")
    Group.Tooltip = "Configure Settings for DMIS 2.0 Modules"
    Group.Items.Add 9004, "Add New User", xtpTaskItemTypeLink, 5
    Group.Items.Add 9005, "Copy Right Access Setting", xtpTaskItemTypeLink, 6
    Group.Items.Add 9006, "Server Settings", xtpTaskItemTypeLink, 7
    Group.Items.Add 9007, "Auditing Inquiry", xtpTaskItemTypeLink, 8
    wndTaskPanel.SetImageList imlTaskPanelIcons
    'wndTaskPanel.SetMargins 5, 5, 5, 5, 5
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            lblModule_Click
        Case vbKeyF2
            lblUserModule_Click
        Case vbKeyF3
            lblCopyRightAccess_Click
        Case vbKeyF4
            lblAuditReport_Click
            'Case vbKeyF5
            '    lblUserMaintain_Click
        Case vbKeyF6
            lblServerSetting_Click
        Case vbKeyF7
            lblUserRightAccessReport_Click
        Case vbKeyF8
            lblAuditInquiry_Click

    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo adder:
    CommandBarsGlobalSettings.Office2007Images = App.Path & "\STYLES\theme.dll"
    CommandBars1.Icons = ImageManager1.Icons
    CreateRibbonBar
    CommandBars1.PaintManager.RefreshMetrics
    CommandBars1.RecalcLayout
    ConfigurePopUps
    ApplyThemes
    picClient.Refresh
    RibbonBar.EnableFrameTheme
    Dim ContextMenu As CommandBar, Control As CommandBarControl
    Set ContextMenu = CommandBars1.ContextMenus.Add(400, "Context Menu")
    Dim ToolTipContext                  As ToolTipContext
    CommandBars1.EnableCustomization False
    CommandBars1.LoadCommandBars "DMIS ADMINISTRATION", App.title, "Layout"

    Set ToolTipContext = CommandBars1.ToolTipContext
    ToolTipContext.Style = xtpToolTipOffice2007
    ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
    ToolTipContext.SetMargin 2, 2, 2, 2
    ToolTipContext.MaxTipWidth = 180
    CreateTaskPanel

    Exit Sub
adder:
    MsgBox Err.Description
End Sub
''to know the server setting
Sub SETCAPTION()
    LABDATABASE = Database
    LABSERVER = ServerName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CommandBars1.SaveCommandBars "DMIS ADMINISTRATION", App.title, "Layout"
    Dim frm                             As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub

Private Sub imgAuditInquiry_Click()
    lblAuditInquiry_Click
End Sub

Private Sub imgAuditReport_Click()
    lblAuditReport_Click
End Sub

Private Sub imgCopyRightAccess_Click()
    lblCopyRightAccess_Click
End Sub

Private Sub imgImportMoudle_Click()
    lblImportModules_Click
End Sub

Private Sub imgServerSettings_Click()
    lblServerSetting_Click
End Sub

Private Sub imgUserMaintain_Click()
    lblUserMaintain_Click
End Sub

Private Sub imgUserModule_Click()
    lblUserModule_Click
End Sub

Private Sub imgUserRightAcessReport_Click()
    lblUserRightAccessReport_Click
End Sub

Private Sub imModules_Click()
    lblModule_Click
End Sub



Private Sub lblAuditInquiry_Click()
    frmInquiry_Audit.Show 1
End Sub

Private Sub lblAuditReport_Click()
    frmReportAuditReport.Show 1
End Sub

Private Sub lblCopyRightAccess_Click()
    frmFiles_CopySettings.Show 1
End Sub

Private Sub lblImportModules_Click()
    frmImportModules.Show 1
End Sub

Private Sub lblModule_Click()
    frmFiles_Modules.Show 1
End Sub

Private Sub lblServerSetting_Click()
    frmFiles_ServerSetting.intSteps = 0
    frmFiles_ServerSetting.Show 1
    SETCAPTION
End Sub

Private Sub lblUserMaintain_Click()
    frmfiles_Users.Show
    frmfiles_Users.ZOrder 0
End Sub

Private Sub lblUserModule_Click()
    frmMain.Hide
    frmFiles_AcessManagement.Show 1
    frmMain.Show
End Sub

Private Sub lblUserRightAccessReport_Click()
    frmReportUserAcess.Show 1
End Sub

Private Sub picClient_Paint()
    CommandBars1.PaintManager.FillWorkspace picClient.hdc, 0, 0, picClient.Width, picClient.Height
End Sub

Private Sub PopCntrl_ItemClick(ByVal item As XtremeSuiteControls.IPopupControlItem)
    If item.Id = 707 Then
        PopCntrl.Close
    End If
End Sub

Private Function RibbonBar() As RibbonBar
    Set RibbonBar = CommandBars1.ActiveMenuBar
End Function

Private Sub wndTaskPanel_ItemClick(ByVal item As XtremeTaskPanel.ITaskPanelGroupItem)

    Select Case item.Id
        Case 9001                                            'MODULES
            'add edit modules

        Case 9002                                            'BROWSE MODULE

        Case 9003                                            'import module

        Case 9004                                            'Add New User

        Case 9005                                            'Copy Right Access Setting
            On Error Resume Next

        Case 9006                                            'Server Settings

        Case 9007
            frmInquiry_Audit.Show
    End Select

End Sub


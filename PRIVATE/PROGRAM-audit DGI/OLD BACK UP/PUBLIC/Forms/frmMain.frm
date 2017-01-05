VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#10.4#0"; "COAAD5~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#10.4#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COADC6~1.OCX"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "WIZENCRYPT.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   10320
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin wizEncrypt.wizEnc wizEnc1 
      Left            =   5550
      Top             =   2760
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   5010
      Top             =   4530
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7425
      Width           =   10320
      _ExtentX        =   18203
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
            TextSave        =   "5:46 PM"
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
      Left            =   4440
      Top             =   4530
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   6
      DesignerControls=   "frmMain.frx":1B0AF
   End
   Begin XtremeSuiteControls.PopupControl PopCntrl 
      Left            =   3870
      Top             =   4545
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
      Animation       =   2
      AnimateDelay    =   125
      ShowDelay       =   2500
      BackgroundBitmap=   60
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    ApplyThemes
    ConfigurePopUps

End Sub
Private Sub ApplyThemes()
    CommandBarsGlobalSettings.App = App
    With CommandBars1
        .LoadDesignerBars
        '.PaintManager.ClearTypeTextQuality = True
        .TabWorkspace.ThemedBackColor = False
        .PaintManager.ThemedStatusBar = True
        .PaintManager.ThemedCheckBox = True
        .StatusBar.Visible = True
    End With
    With SkinFramework1
        .LoadSkin "\\server\HMI\SKINS\Royale.cjstyles", ""
        .ApplyWindow Me.hwnd
        '.ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or xtpSkinApplyMetrics
        .ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or Not xtpSkinApplyMetrics
    End With
    Dim ToolTipContext                       As ToolTipContext
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
    Dim Item                                 As PopupControlItem
    PopCntrl.RemoveAllItems
    'PopCntrl.Icons.AddIcons ImageManager.Icons
    PopCntrl.Icons.AddIcons CommandBars1.Icons
    PopCntrl.VisualTheme = xtpPopupThemeMSN
    'PopCntrl.SetSize 270, 140

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

Private Sub MDIForm_Unload(Cancel As Integer)
Dim frm As Form
    For Each frm In Forms
    Unload frm
    Next
End Sub

Private Sub PopCntrl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.Id = 707 Then
        PopCntrl.Close
    End If

End Sub
''''''''''''''END REGION POPUPCONTROLS''''''''''''''''''''''''''''''''''''



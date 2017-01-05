VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSystemConfiguration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Configurations"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SystemConfiguration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   7035
   Begin VB.TextBox txtSkinPath 
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   29
      Tag             =   "PATH"
      ToolTipText     =   "SKIN"
      Top             =   5280
      Width           =   4875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5010
      TabIndex        =   28
      Top             =   5250
      Width           =   645
   End
   Begin VB.CommandButton cmdHRMSRptPath 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5010
      TabIndex        =   19
      Top             =   4500
      Width           =   645
   End
   Begin VB.TextBox txtHRMS 
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      Tag             =   "HRMS"
      ToolTipText     =   "REPORTPATH"
      Top             =   4530
      Width           =   4875
   End
   Begin VB.CommandButton cmdSMISRptPath 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5010
      TabIndex        =   11
      Top             =   3780
      Width           =   645
   End
   Begin VB.TextBox txtSMIS 
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Tag             =   "SMIS"
      ToolTipText     =   "REPORTPATH"
      Top             =   3780
      Width           =   4875
   End
   Begin VB.CommandButton cmdPMISRptPath 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5010
      TabIndex        =   9
      Top             =   3120
      Width           =   645
   End
   Begin VB.TextBox txtPMIS 
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Tag             =   "PMIS"
      ToolTipText     =   "REPORTPATH"
      Top             =   3090
      Width           =   4875
   End
   Begin VB.CommandButton cmdCRISRptPath 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5010
      TabIndex        =   7
      Top             =   2400
      Width           =   645
   End
   Begin VB.TextBox txtCRIS 
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Tag             =   "CRIS"
      ToolTipText     =   "REPORTPATH"
      Top             =   2370
      Width           =   4875
   End
   Begin VB.CommandButton cmdCMISRptPath 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5010
      TabIndex        =   5
      Top             =   1680
      Width           =   645
   End
   Begin VB.TextBox txtCMIS 
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Tag             =   "CMIS"
      ToolTipText     =   "REPORTPATH"
      Top             =   1680
      Width           =   4875
   End
   Begin VB.CommandButton cmdCSMSRptPath 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5010
      TabIndex        =   3
      Top             =   960
      Width           =   645
   End
   Begin VB.CommandButton cmdAMISRptPath 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5010
      TabIndex        =   2
      Top             =   270
      Width           =   645
   End
   Begin VB.TextBox txtCSMS 
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "CSMS"
      ToolTipText     =   "REPORTPATH"
      Top             =   960
      Width           =   4875
   End
   Begin VB.TextBox txtAmis 
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Tag             =   "AMIS"
      ToolTipText     =   "REPORTPATH"
      Top             =   270
      Width           =   4875
   End
   Begin VB.PictureBox picFree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   3450
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   4635
      ScaleWidth      =   3360
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   3390
      Begin VB.FileListBox File1 
         Height          =   1440
         Left            =   90
         TabIndex        =   31
         Top             =   2670
         Width           =   3255
      End
      Begin VB.DirListBox Dir1 
         Height          =   1875
         Left            =   60
         TabIndex        =   27
         Top             =   720
         Width           =   3225
      End
      Begin VB.DriveListBox Drive1 
         Height          =   345
         Left            =   60
         TabIndex        =   26
         Top             =   360
         Width           =   3285
      End
      Begin VB.CommandButton cmdClosePicFREE 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2970
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.CommandButton cmdCancelFree 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   4170
         Width           =   795
      End
      Begin VB.CommandButton cmdOkFree 
         Caption         =   "OK"
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   4170
         Width           =   795
      End
      Begin XtremeShortcutBar.ShortcutCaption capFree 
         Height          =   330
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   3375
         _Version        =   655364
         _ExtentX        =   5953
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "Select Your Report Path"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   -2147483630
      End
   End
   Begin VB.Label Label 
      Caption         =   "Skin Path"
      Height          =   345
      Index           =   7
      Left            =   120
      TabIndex        =   30
      Top             =   5010
      Width           =   1605
   End
   Begin VB.Label Label 
      Caption         =   "HRMS  Report Path"
      Height          =   345
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   4260
      Width           =   1605
   End
   Begin VB.Label Label 
      Caption         =   "SMIS Report Path"
      Height          =   345
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   3540
      Width           =   1605
   End
   Begin VB.Label Label 
      Caption         =   "CRIS Report Path"
      Height          =   345
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   2130
      Width           =   1605
   End
   Begin VB.Label Label 
      Caption         =   "PMIS Report Path"
      Height          =   345
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   2820
      Width           =   1605
   End
   Begin VB.Label Label 
      Caption         =   "CMIS Report Path"
      Height          =   345
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1605
   End
   Begin VB.Label Label 
      Caption         =   "CSMS Report Path"
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   1605
   End
   Begin VB.Label Label 
      Caption         =   "AMIS Report Path"
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   30
      Width           =   1605
   End
End
Attribute VB_Name = "frmSystemConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents SelectedTextBox         As TextBox
Attribute SelectedTextBox.VB_VarHelpID = -1
Private Sub cmdAMISRptPath_Click()
    ShowHidePictureBox2 picFree, True
    capFree.Caption = "Select Your Report Path for AMIS"
    Dir1.Height = 3405
    Set SelectedTextBox = txtAmis
End Sub

Private Sub cmdCancelFree_Click()
    ShowHidePictureBox2 picFree, False
    cmdOkFree.Enabled = False
End Sub
Private Sub cmdClosePicFREE_Click()
    ShowHidePictureBox2 picFree, False
    File1.Visible = False
    cmdOkFree.Enabled = False
End Sub

Private Sub cmdCMISRptPath_Click()
    ShowHidePictureBox2 picFree, True
    Set SelectedTextBox = txtCMIS
    Dir1.Height = 3405
    capFree.Caption = "Select Your Report Path for CMIS"
End Sub

Private Sub cmdCRISRptPath_Click()
    ShowHidePictureBox2 picFree, True
    Set SelectedTextBox = txtCRIS
    Dir1.Height = 3405
    capFree.Caption = "Select Your Report Path for CRIS"
End Sub

Private Sub cmdCSMSRptPath_Click()
    ShowHidePictureBox2 picFree, True
    Set SelectedTextBox = txtCSMS
    Dir1.Height = 3405
    capFree.Caption = "Select Your Report Path for CSMS"
End Sub


Private Sub cmdHRMSRptPath_Click()
    ShowHidePictureBox2 picFree, True
    Set SelectedTextBox = txtHRMS
    Dir1.Height = 3405
    capFree.Caption = "Select Your Report Path for HRMS"
End Sub

Private Sub cmdOkFree_Click()


    If SelectedTextBox.Name = "txtSkinPath" Then
        SelectedTextBox.Text = File1.Path & "\" & File1.FileName



    Else
        SelectedTextBox.Text = Dir1.Path & "\"

    End If

    ShowHidePictureBox2 picFree, False
    cmdOkFree.Enabled = False
    Set SelectedTextBox = Nothing
End Sub

Private Sub cmdPMISRptPath_Click()
    ShowHidePictureBox2 picFree, True
    Set SelectedTextBox = txtPMIS
    Dir1.Height = 3405
    capFree.Caption = "Select Your Report Path for PMIS"
End Sub

Private Sub cmdSMISRptPath_Click()
    ShowHidePictureBox2 picFree, True
    Set SelectedTextBox = txtSMIS
    Dir1.Height = 3405
    capFree.Caption = "Select Your Report Path for SMIS"
End Sub


Private Sub Command1_Click()
    ShowHidePictureBox2 picFree, True
    Set SelectedTextBox = txtSkinPath
    File1.Pattern = "*.cjstyles"
    Dir1.Height = 1875
    File1.Visible = True
    capFree.Caption = "Select Your Skin Path For Application"
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Dir1_Click()
    cmdOkFree.Enabled = True
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOkFree.Value = True
    End If
End Sub

Private Sub Drive1_Change()
    On Error GoTo ADDER:
    Dir1.Path = Drive1.Drive
    Exit Sub
ADDER:
    Err.Clear

End Sub

Private Sub File1_Click()
    cmdOkFree.Enabled = True
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    txtAmis = GetSetting("DMIS", "REPORTPATH", "AMIS", vbNullString)
    txtCMIS = GetSetting("DMIS", "REPORTPATH", "CMIS", vbNullString)
    txtCRIS = GetSetting("DMIS", "REPORTPATH", "CRIS", vbNullString)
    txtCSMS = GetSetting("DMIS", "REPORTPATH", "CSMS", vbNullString)
    txtHRMS = GetSetting("DMIS", "REPORTPATH", "HRMS", vbNullString)
    txtPMIS = GetSetting("DMIS", "REPORTPATH", "PMIS", vbNullString)
    txtSMIS = GetSetting("DMIS", "REPORTPATH", "SMIS", vbNullString)
    txtSkinPath = GetSetting("DMIS", "SKIN", "PATH", vbNullString)
End Sub
Private Sub SelectedTextBox_Change()
    Call SaveSetting("DMIS", SelectedTextBox.ToolTipText, SelectedTextBox.Tag, SelectedTextBox.Text)

    AMIS_REPORT_PATH = GetSetting("DMIS", "REPORTPATH", "AMIS", vbNullString)
    CMIS_REPORT_PATH = GetSetting("DMIS", "REPORTPATH", "CMIS", vbNullString)
    CRIS_REPORT_PATH = GetSetting("DMIS", "REPORTPATH", "CRIS", vbNullString)
    CSMIOS_REPORT_PATH = GetSetting("DMIS", "REPORTPATH", "CSMS", vbNullString)
    HRMS_REPORT_PATH = GetSetting("DMIS", "REPORTPATH", "HRMS", vbNullString)
    PMIOS_REPORT_PATH = GetSetting("DMIS", "REPORTPATH", "PMIS", vbNullString)
    SMIS_REPORT_PATH = GetSetting("DMIS", "REPORTPATH", "SMIS", vbNullString)
    SKIN_PATH = GetSetting("DMIS", "SKIN", "PATH", vbNullString)

    If SelectedTextBox.Name = "txtSkinPath" Then
        frmMain.ApplyThemes
    End If

End Sub


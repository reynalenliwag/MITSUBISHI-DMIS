VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_MISLogInSE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In"
   ClientHeight    =   4650
   ClientLeft      =   4005
   ClientTop       =   3375
   ClientWidth     =   4740
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   Icon            =   "MISLogInSE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4650
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2850
      Left            =   45
      Picture         =   "MISLogInSE.frx":74F2
      ScaleHeight     =   2850
      ScaleWidth      =   2130
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   870
      Width           =   2130
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3750
      MaskColor       =   &H00FFFFFF&
      Picture         =   "MISLogInSE.frx":A62D
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   3930
      Width           =   885
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Picture         =   "MISLogInSE.frx":A96B
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Log-In"
      Top             =   3930
      Width           =   885
   End
   Begin VB.TextBox txtUserPass 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   22
      PasswordChar    =   "l"
      TabIndex        =   3
      Top             =   3300
      Width           =   2445
   End
   Begin MSComctlLib.ListView lstUserName 
      Height          =   1890
      Left            =   2205
      TabIndex        =   2
      Top             =   900
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   3334
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "MISLogInSE.frx":AC06
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "USERNAME"
         Object.Width           =   5644
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7665
      _Version        =   655364
      _ExtentX        =   13520
      _ExtentY        =   1535
      _StockProps     =   14
      Caption         =   "           Agent Login Window"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
End
Attribute VB_Name = "frmSMIS_MISLogInSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo ErrorCode

    If lstUserName.SelectedItem Is Nothing Then: Exit Sub
    Dim RSLOGSAE                                                      As ADODB.Recordset
    Dim VALIDPASS                                                     As Integer
    If RTrim(LTrim(txtUserPass)) <> "" Then

        Set RSLOGSAE = gconDMIS.Execute("SELECT SAECODE FROM  SMIS_vw_Srep WHERE SAECODE='" & lstUserName.SelectedItem.ListSubItems(1).Text & "'")

        If Not RSLOGSAE.EOF Or Not RSLOGSAE.BOF Then
            VALIDPASS = gconDMIS.Execute("SELECT COUNT(*) FROM  SMIS_SALESTEAM WHERE SAECODE='" & RSLOGSAE("SAECODE") & "' AND  PASSWORD ='" & txtUserPass & "'").Fields(0).Value
            If VALIDPASS = 0 Then
                MsgBox "INVALID PASSWORD ", vbCritical
                On Error Resume Next
                txtUserPass.SetFocus
            Else
                LOGSAE = RSLOGSAE("SAECODE")
                SAENAME = lstUserName.SelectedItem.Text
                LogAudit "A", "AGENT LOGIN " & "LOGIN NAME:" & SAENAME
                Unload Me
                MainSAE.Show
            End If
        Else
            MsgBox "INVALID SALES EXECUTIVE NAME ", vbInformation
            On Error Resume Next
            lstUserName.SetFocus

        End If
    Else
        SAECODE = ""
        Exit Sub
    End If


    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_Activate()
    lstUserName.SetFocus
    If lstUserName.ListItems.Count > 0 Then
        lstUserName.ListItems(1).Selected = True
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim rsALL_vw_RAMS_PAccess                                         As ADODB.Recordset
    Set rsALL_vw_RAMS_PAccess = New ADODB.Recordset
    rsALL_vw_RAMS_PAccess.Open "SELECT [NAME] , SAECODE  FROM SMIS_vw_Srep ORDER BY [NAME]", gconACCESS, adOpenKeyset
    If Not rsALL_vw_RAMS_PAccess.EOF And Not rsALL_vw_RAMS_PAccess.BOF Then
        Listview_Loadval lstUserName.ListItems, rsALL_vw_RAMS_PAccess
    Else
        MsgBox " There No Sales Agent Configured. This Window Will Close", vbCritical
        Unload Me
    End If
    txtUserPass.Enabled = False
End Sub

Private Sub lstUserName_DblClick()
    On Error Resume Next
    txtUserPass.Enabled = True
    txtUserPass.SetFocus
End Sub

Private Sub lstUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtUserPass.Enabled = True
        On Error Resume Next
        txtUserPass.SetFocus
    End If
End Sub

Private Sub txtUserPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(LTrim(RTrim(txtUserPass))) > 0 Then
        SendKeys "{TAB}"
    End If

End Sub


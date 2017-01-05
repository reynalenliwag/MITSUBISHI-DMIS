VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAISEMPLOYMENT 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2985
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7740
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   7740
   Begin VB.PictureBox picTRAIN_SAVE 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   5310
      ScaleHeight     =   825
      ScaleWidth      =   2295
      TabIndex        =   8
      Top             =   2100
      Width           =   2295
      Begin VB.CommandButton cmdCANCEL 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1440
         MouseIcon       =   "frmAISEMPLOYMENT.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmAISEMPLOYMENT.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancel Entry"
         Top             =   -30
         Width           =   735
      End
      Begin VB.CommandButton cmdDELETE 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   720
         MouseIcon       =   "frmAISEMPLOYMENT.frx":0490
         MousePointer    =   99  'Custom
         Picture         =   "frmAISEMPLOYMENT.frx":05E2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Delete Entry"
         Top             =   -30
         Width           =   735
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   0
         MouseIcon       =   "frmAISEMPLOYMENT.frx":090D
         MousePointer    =   99  'Custom
         Picture         =   "frmAISEMPLOYMENT.frx":0A5F
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Save Entry"
         Top             =   -30
         Width           =   735
      End
   End
   Begin VB.TextBox txtEMP_FROM 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2190
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1560
      Width           =   5325
   End
   Begin VB.TextBox txtEMP_POS 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2190
      MaxLength       =   35
      TabIndex        =   2
      Top             =   1170
      Width           =   5325
   End
   Begin VB.TextBox txtCOMP_ADD 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2190
      MaxLength       =   40
      TabIndex        =   1
      Top             =   780
      Width           =   5325
   End
   Begin VB.TextBox txtCOMP_Name 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2190
      MaxLength       =   30
      TabIndex        =   0
      Top             =   390
      Width           =   5325
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name Of Company"
      Height          =   240
      Index           =   20
      Left            =   300
      TabIndex        =   7
      Top             =   390
      Width           =   1815
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From - To"
      Height          =   240
      Index           =   7
      Left            =   1110
      TabIndex        =   6
      Top             =   1530
      Width           =   990
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      Height          =   240
      Index           =   6
      Left            =   1320
      TabIndex        =   5
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Address"
      Height          =   240
      Index           =   22
      Left            =   360
      TabIndex        =   4
      Top             =   780
      Width           =   1755
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7845
      _Version        =   655364
      _ExtentX        =   13838
      _ExtentY        =   503
      _StockProps     =   14
      Caption         =   "       "
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12.01
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      ForeColor       =   0
   End
End
Attribute VB_Name = "frmAISEMPLOYMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FUNCTION /FEATURE:Change all the Button and the designed of the form make it clear
'DATE STARTED:07/03/2007
'LAST UPDATE:
'DATABASE UPDATE:
'WHO UPDATE:HardNard
'UPDATING CODE:BTT - 07/03/2007
'**********************************************************************************

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    'If Function_Access(LOGID, "ACESS_DELETE", "APPLICANT INFO") = False Then Exit Sub
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    If MsgBox("Delete Employement Record", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
        gconDMIS.Execute ("Delete From HRMS_APPLICANT_EMPLOYMENT_RECORD Where Applicant_id = " & _
                          APPLICANT_ID & " And Entry_ID = " & EMP_ENTRY_ID & "")

        Unload Me
        Call frmAISApplications.DisplayEmploymentInListView
    End If

    Exit Sub
Errorcode:
    ShowVBError

End Sub

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:16
Private Sub cmdSave_Click()
    Dim vtxtCOMP_NAME As String, vtxtCOMP_ADD As String, vtxtEMP_POS As String, vtxtEMP_FROM As String
    Dim ID                                                            As Integer

    On Error GoTo Errorcode:

    If txtCOMP_Name.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Employement Records"
        On Error Resume Next
        txtCOMP_Name.SetFocus
        Exit Sub
    End If
    If txtCOMP_ADD.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Employement Records"
        On Error Resume Next
        txtCOMP_ADD.SetFocus
        Exit Sub
    End If
    If txtEMP_POS.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Employement Records"
        On Error Resume Next
        txtEMP_POS.SetFocus
        Exit Sub
    End If
    If txtEMP_FROM.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Employement Records"
        On Error Resume Next
        txtEMP_FROM.SetFocus
        Exit Sub
    End If

    vtxtCOMP_NAME = N2Str2Null(txtCOMP_Name)
    vtxtCOMP_ADD = N2Str2Null(txtCOMP_ADD)
    vtxtEMP_POS = N2Str2Null(txtEMP_POS)
    vtxtEMP_FROM = N2Str2Null(txtEMP_FROM)


    frmMain.MousePointer = 11
    If SAVE_OR_EDIT_EMP = "SAVE" Then
        Call GenerateNewID("HRMS_APPLICANT_EMPLOYMENT_RECORD", ID)

        gconDMIS.Execute ("Insert Into HRMS_APPLICANT_EMPLOYMENT_RECORD Values(" & APPLICANT_ID & _
                          "," & ID & "," & vtxtCOMP_NAME & "," & vtxtCOMP_ADD & "," & vtxtEMP_POS & "," & vtxtEMP_FROM & ")")
    Else
        gconDMIS.Execute ("Update HRMS_APPLICANT_EMPLOYMENT_RECORD Set NameOfCompany = " & vtxtCOMP_NAME & _
                          ",Address = " & vtxtCOMP_ADD & _
                          ",Posisyon = " & vtxtEMP_POS & _
                          ",From_To = " & vtxtEMP_FROM & _
                        " Where Applicant_ID = " & APPLICANT_ID & " And Entry_ID = " & EMP_ENTRY_ID & "")
    End If

    Unload Me

    Call frmAISApplications.DisplayEmploymentInListView
    frmMain.MousePointer = 11

    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 11
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISApplications.Enabled = True
    frmAISApplications.picSaves.Visible = True
    On Error Resume Next
    frmAISApplications.SetFocus
End Sub


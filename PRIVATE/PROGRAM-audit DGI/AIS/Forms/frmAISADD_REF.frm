VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAISADD_REF 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3150
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6330
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6330
   Begin VB.PictureBox picTRAIN_SAVE 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   3840
      ScaleHeight     =   765
      ScaleWidth      =   2400
      TabIndex        =   7
      Top             =   2280
      Width           =   2400
      Begin VB.CommandButton cmdREF_CANCEL 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1560
         Picture         =   "frmAISADD_REF.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdREF_DELETE 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   780
         Picture         =   "frmAISADD_REF.frx":0552
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Delete Entry"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdREF_SAVE 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   0
         Picture         =   "frmAISADD_REF.frx":0C3C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Save Entry"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.TextBox txtREF_NAME 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1050
      MaxLength       =   35
      TabIndex        =   0
      Top             =   420
      Width           =   5115
   End
   Begin VB.TextBox txtREF_POS 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1050
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1380
      Width           =   5115
   End
   Begin VB.TextBox txtREF_TEL 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1050
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1830
      Width           =   2925
   End
   Begin VB.TextBox txtREF_ADD 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1050
      MaxLength       =   50
      TabIndex        =   1
      Top             =   900
      Width           =   5100
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   240
      Index           =   0
      Left            =   330
      TabIndex        =   11
      Top             =   540
      Width           =   540
   End
   Begin VB.Label lblCAP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tel No."
      Height          =   240
      Index           =   20
      Left            =   240
      TabIndex        =   10
      Top             =   1950
      Width           =   705
   End
   Begin VB.Label lblCAP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      Height          =   240
      Index           =   23
      Left            =   150
      TabIndex        =   9
      Top             =   1470
      Width           =   765
   End
   Begin VB.Label lblCAP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   240
      Index           =   28
      Left            =   180
      TabIndex        =   8
      Top             =   1020
      Width           =   780
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   270
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6345
      _Version        =   655364
      _ExtentX        =   11192
      _ExtentY        =   476
      _StockProps     =   14
      Caption         =   "       "
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.99
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
Attribute VB_Name = "frmAISADD_REF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function CheckIfComplete() As Boolean
    If txtREF_NAME.Text = "" Then
        MsgBox "Incompete Entry", vbExclamation, "Character Reference"
        On Error Resume Next
        txtREF_NAME.SetFocus
        Exit Function
    End If
    If txtREF_ADD.Text = "" Then
        MsgBox "Incompete Entry", vbExclamation, "Character Reference"
        On Error Resume Next
        txtREF_ADD.SetFocus
        Exit Function
    End If
    If txtREF_POS.Text = "" Then
        MsgBox "Incompete Entry", vbExclamation, "Character Reference"
        On Error Resume Next
        txtREF_POS.SetFocus
        Exit Function
    End If
    If txtREF_TEL.Text = "" Then
        MsgBox "Incompete Entry", vbExclamation, "Character Reference"
        On Error Resume Next
        txtREF_TEL.SetFocus
        Exit Function
    End If
    CheckIfComplete = True
End Function

Private Sub cmdREF_CANCEL_Click()
    Unload Me
End Sub

Private Sub cmdREF_DELETE_Click()
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_DELETE", "APPLICANT INFO") = False Then Exit Sub
    frmMain.MousePointer = 11
    If MsgBox("Delete Chracter Reference", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
        gconDMIS.Execute ("Delete From HRMS_APPLICANT_REFERENCE Where Applicant_id = " & APPLICANT_ID & _
                        " And Entry_ID = " & REFERENCE_ENTRY_ID & "")

        Call LogAudit("X", "DELETE APPLICANT REFERENCE", APPLICANT_ID)
        Unload Me
        Call frmAISApplications.DisplayReferenceInListView
    End If
    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

'Upating Code       : AXP-0707200711:13
Private Sub cmdREF_SAVE_Click()
    Dim vtxtREF_NAME As String, vtxtREF_ADD, vtxtREF_POS As String, vtxtREF_TEL As String
    Dim ID                                                            As Integer

    On Error GoTo Errorcode:

    If CheckIfComplete = False Then
        Exit Sub
    End If

    vtxtREF_NAME = N2Str2Null(txtREF_NAME)
    vtxtREF_ADD = N2Str2Null(txtREF_ADD)
    vtxtREF_POS = N2Str2Null(txtREF_POS)
    vtxtREF_TEL = N2Str2Null(txtREF_TEL)

    'On Error GoTo ERROR
    frmMain.MousePointer = 11
    If SAVE_OR_EDIT_REF = "SAVE" Then
        Call GenerateNewID("HRMS_APPLICANT_REFERENCE", ID)

        gconDMIS.Execute ("Insert Into HRMS_APPLICANT_REFERENCE Values(" & APPLICANT_ID & _
                          "," & ID & "," & vtxtREF_NAME & "," & vtxtREF_ADD & "," & vtxtREF_POS & "," & vtxtREF_TEL & ")")

        Call LogAudit("A", "ADD APPLICANT REFERENCE", APPLICANT_ID)
    Else
        gconDMIS.Execute ("Update HRMS_APPLICANT_REFERENCE Set Name = " & vtxtREF_NAME & _
                          ",Address = " & vtxtREF_ADD & _
                          ",Posisyon = " & vtxtREF_POS & _
                          ",TelNo = " & vtxtREF_TEL & _
                        " Where Applicant_Id = " & APPLICANT_ID & _
                        " And Entry_ID = " & REFERENCE_ENTRY_ID & "")

        Call LogAudit("E", "UPDATE APPLICANT REFERENCE", APPLICANT_ID)
    End If

    Unload Me
    Call frmAISApplications.DisplayReferenceInListView
    frmMain.MousePointer = 0

    Exit Sub

ERROR:
    MsgBox "ERROR", vbCritical, "Character Reference"
    frmMain.MousePointer = 0

    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
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


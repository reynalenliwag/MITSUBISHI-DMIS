VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAISADD_TRAIN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4650
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6750
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
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
   Moveable        =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6750
   Begin VB.TextBox txtTRAIN_MonthYear 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1335
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1800
      Width           =   5190
   End
   Begin VB.TextBox txtTRAIN_SPONSOR 
      Appearance      =   0  'Flat
      Height          =   585
      Left            =   1350
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2850
      Width           =   5205
   End
   Begin VB.TextBox txtTRAIN_PLACE 
      Appearance      =   0  'Flat
      Height          =   585
      Left            =   1350
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2190
      Width           =   5205
   End
   Begin VB.TextBox txtTRAIN_TRAIN 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   1350
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   5145
   End
   Begin VB.PictureBox picTRAIN_SAVE 
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
      Height          =   735
      Left            =   4080
      ScaleHeight     =   735
      ScaleWidth      =   2475
      TabIndex        =   7
      Top             =   3690
      Width           =   2475
      Begin VB.CommandButton cmdTRAIN_CANCEL 
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
         Height          =   735
         Left            =   1620
         Picture         =   "frmADD_TRAIN.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exit WIndow"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTRAIN_DELETE 
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
         Height          =   735
         Left            =   840
         Picture         =   "frmADD_TRAIN.frx":0552
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Delete Trainings and Seminars"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTRAIN_SAVE 
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
         Height          =   735
         Left            =   60
         Picture         =   "frmADD_TRAIN.frx":0C3C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Save Trainings and Seminars"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Label lblCAP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month-Year"
      Height          =   240
      Index           =   28
      Left            =   90
      TabIndex        =   11
      Top             =   1875
      Width           =   1170
   End
   Begin VB.Label lblCAP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Place"
      Height          =   240
      Index           =   23
      Left            =   750
      TabIndex        =   10
      Top             =   2235
      Width           =   525
   End
   Begin VB.Label lblCAP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sponsor"
      Height          =   240
      Index           =   20
      Left            =   465
      TabIndex        =   9
      Top             =   2880
      Width           =   795
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Training"
      Height          =   240
      Index           =   0
      Left            =   420
      TabIndex        =   8
      Top             =   360
      Width           =   780
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   390
      Left            =   0
      TabIndex        =   12
      Top             =   -120
      Width           =   6765
      _Version        =   655364
      _ExtentX        =   11933
      _ExtentY        =   688
      _StockProps     =   14
      Caption         =   "       "
      ForeColor       =   0
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
      ForeColor       =   0
   End
End
Attribute VB_Name = "frmAISADD_TRAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function CheckEntryIfComplete() As Boolean
    If txtTRAIN_TRAIN.Text = "" Then
        CheckEntryIfComplete = False
        MsgBox "Incomplete Entry", vbExclamation, "Training and Seminar Attend"
        On Error Resume Next
        txtTRAIN_TRAIN.SetFocus
        Exit Function
    End If
    If txtTRAIN_SPONSOR.Text = "" Then
        CheckEntryIfComplete = False
        MsgBox "Incomplete Entry", vbExclamation, "Training and Seminar Attend"
        On Error Resume Next
        txtTRAIN_SPONSOR.SetFocus
        Exit Function
    End If
    If txtTRAIN_PLACE.Text = "" Then
        CheckEntryIfComplete = False
        MsgBox "Incomplete Entry", vbExclamation, "Training and Seminar Attend"
        On Error Resume Next
        txtTRAIN_PLACE.SetFocus
        Exit Function
    End If
    If txtTRAIN_MonthYear.Text = "" Then
        CheckEntryIfComplete = False
        MsgBox "Incomplete Entry", vbExclamation, "Training and Seminar Attend"
        On Error Resume Next
        txtTRAIN_MonthYear.SetFocus
        Exit Function
    End If
    CheckEntryIfComplete = True
End Function

Private Sub cmdTRAIN_CANCEL_Click()
    Unload Me
End Sub

Private Sub cmdTRAIN_DELETE_Click()
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_DELETE", "APPLICANT INFO") = False Then Exit Sub
    frmMain.MousePointer = 11

    If MsgBox("Are You Sure", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Dependent Child") = vbYes Then
        gconDMIS.Execute ("Delete From HRMS_APPLICANT_TRAIN Where Applicant_ID = " & _
                          APPLICANT_ID & " And Entry_ID = " & TRAINING_ENTRY_ID & "")

        Call LogAudit("X", "DELETE APPLICANT TRAININGS ATTEND", APPLICANT_ID)

        Unload Me
        Call frmAISApplications.DisplayTrainInListView
    End If
    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

'Upating Code       : AXP-0707200711:13
Private Sub cmdTRAIN_SAVE_Click()
    Dim ID                                                            As Integer
    Dim SQL                                                           As String
    Dim VtxtTRAIN_TRAIN As String, VtxtTRAIN_MONTHYEAR                As String
    Dim VtxtTRAIN_PLACE                                               As String
    Dim VtxtTRAIN_SPONSOR                                             As String

    On Error GoTo Errorcode:

    If CheckEntryIfComplete = False Then
        Exit Sub
    End If

    VtxtTRAIN_TRAIN = N2Str2Null(txtTRAIN_TRAIN)
    VtxtTRAIN_MONTHYEAR = N2Str2Null(txtTRAIN_MonthYear)
    VtxtTRAIN_SPONSOR = N2Str2Null(txtTRAIN_SPONSOR)
    VtxtTRAIN_PLACE = N2Str2Null(txtTRAIN_PLACE)

    frmMain.MousePointer = 11
    If SAVE_OR_EDIT_TRAINING = "SAVE" Then                    'NEW
        Call GenerateNewID("HRMS_APPLICANT_TRAIN", ID)
        TRAINING_ENTRY_ID = ID

        SQL = "Insert Into HRMS_APPLICANT_TRAIN Values(" & _
              APPLICANT_ID & "," & _
              TRAINING_ENTRY_ID & "," & _
              VtxtTRAIN_TRAIN & "," & _
              VtxtTRAIN_MONTHYEAR & "," & _
              VtxtTRAIN_PLACE & "," & _
              VtxtTRAIN_SPONSOR & ")"

        Call LogAudit("A", "ADD APPLICANT TRAININGS ATTEND", APPLICANT_ID)
    Else
        SQL = "Update HRMS_APPLICANT_TRAIN Set Training = " & VtxtTRAIN_TRAIN & _
              ",MonthYear = " & VtxtTRAIN_MONTHYEAR & _
              ",place = " & VtxtTRAIN_PLACE & _
              ",Sponsor = " & VtxtTRAIN_SPONSOR & _
            " Where Applicant_ID = " & APPLICANT_ID & _
            " AND Entry_ID = " & TRAINING_ENTRY_ID & ""

        Call LogAudit("E", "UPDATE APPLICANT TRAININGS ATTEND", APPLICANT_ID)
    End If

    gconDMIS.Execute (SQL)

    Unload Me
    Call frmAISApplications.DisplayTrainInListView


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


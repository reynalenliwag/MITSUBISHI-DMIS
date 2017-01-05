VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAIS_SCHEDULE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SCHEDULE OD EXAM / INTERVIEW"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10650
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10650
   Begin VB.CommandButton cmdEXIT 
      Caption         =   "EXIT"
      Height          =   915
      Left            =   9540
      TabIndex        =   2
      Top             =   5970
      Width           =   975
   End
   Begin MSComCtl2.MonthView mvwDATE 
      Height          =   2520
      Left            =   180
      TabIndex        =   1
      Top             =   4380
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4445
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   55050241
      TitleBackColor  =   0
      TitleForeColor  =   255
      TrailingForeColor=   32768
      CurrentDate     =   39317
   End
   Begin MSComctlLib.ListView lsvSCHED 
      Height          =   3675
      Left            =   150
      TabIndex        =   0
      Top             =   540
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6482
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TIME"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ACTIVITY"
         Object.Width           =   14111
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      Caption         =   "DATE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   180
      Width           =   495
   End
   Begin VB.Menu mnuSCHED 
      Caption         =   "SCHEDULE"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuSCHED_I 
         Caption         =   "Schedule an Interview"
         Index           =   0
      End
      Begin VB.Menu mnuSCHED_E 
         Caption         =   "Schedule An Exam"
         Index           =   0
      End
   End
   Begin VB.Menu mnuRESCHED1 
      Caption         =   "RESCHEDULE1"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuRESCHED1_I 
         Caption         =   "View Result Of Interview"
         Index           =   0
      End
   End
   Begin VB.Menu mnuRESCHED2 
      Caption         =   "RESCHEDULE2"
      Visible         =   0   'False
      Begin VB.Menu mnuRESCHED2_I 
         Caption         =   "View Result Of Interview"
         Index           =   0
      End
      Begin VB.Menu mnuRESCHED2_I 
         Caption         =   "Reschedule Interview"
         Index           =   1
      End
   End
   Begin VB.Menu mnuRESCHED3 
      Caption         =   "RESCHEDULE3"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuRESCHED1_E 
         Caption         =   "View Result Of Exam"
         Index           =   0
      End
   End
   Begin VB.Menu mnuRESCHED4 
      Caption         =   "RESCHEDULE4"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuRESCHED2_E 
         Caption         =   "View Result Of Exam"
         Index           =   0
      End
      Begin VB.Menu mnuRESCHED2_E 
         Caption         =   "Reschedule Exam"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmAIS_SCHEDULE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEXIT_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape:
            
    
    End Select
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

Private Sub lsvSCHED_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Index As Double
    
    If Not lsvSCHED.ListItems.Count = 0 Then
        Index = lsvSCHED.SelectedItem.Index
        
        With lsvSCHED
            If Button = vbRightButton Then
                If .ListItems(Index).SubItems(2) = "" Then
                    PopupMenu mnuSCHED
                Else
                    If Date > mvwDATE.Value Then
                        PopupMenu mnuRESCHED1
                    Else
                        PopupMenu mnuRESCHED2
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub mnuRESCHED_I_Click(Index As Integer)

End Sub

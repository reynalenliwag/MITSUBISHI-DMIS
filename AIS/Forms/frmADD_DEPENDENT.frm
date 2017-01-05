VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAISADD_DEPENDENT 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1965
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6915
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   6915
   Begin VB.CheckBox chkDependent 
      Caption         =   "Dependent Child"
      Height          =   285
      Left            =   3900
      TabIndex        =   2
      Top             =   750
      Width           =   2775
   End
   Begin VB.PictureBox picCHILD_SAVE 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2460
      ScaleHeight     =   495
      ScaleWidth      =   4155
      TabIndex        =   8
      Top             =   1260
      Width           =   4215
      Begin VB.CommandButton cmdCHILD_DELETE 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1410
         TabIndex        =   4
         Top             =   0
         Width           =   1365
      End
      Begin VB.CommandButton cmdCHILD_SAVE 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   30
         TabIndex        =   3
         Top             =   0
         Width           =   1365
      End
      Begin VB.CommandButton cmdCHILD_CANCEL 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2790
         TabIndex        =   5
         Top             =   0
         Width           =   1365
      End
   End
   Begin VB.TextBox txtCHILD_FULL 
      Height          =   315
      Left            =   1650
      MaxLength       =   30
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker DTPCHILD 
      Height          =   345
      Left            =   1650
      TabIndex        =   1
      Top             =   690
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   609
      _Version        =   393216
      Format          =   53477377
      CurrentDate     =   39125
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date"
      Height          =   240
      Index           =   21
      Left            =   360
      TabIndex        =   7
      Top             =   810
      Width           =   990
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      Height          =   240
      Index           =   24
      Left            =   330
      TabIndex        =   6
      Top             =   360
      Width           =   945
   End
End
Attribute VB_Name = "frmAISADD_DEPENDENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCHILD_CANCEL_Click()
    Unload Me
End Sub

Private Sub cmdCHILD_DELETE_Click()
'    If MsgBox("Are You Sure", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Dependent Child") = vbYes Then
'        gconDMIS.Execute ("Delete From HRMS_APPLICANT_CHILD Where Applicant_ID = " & _
'                           APPLICANT_ID & " And Entry_ID = " & CHILD_ENTRY_ID & "")
'
'        Unload Me
'    End If
End Sub

'Private Sub cmdCHILD_SAVE_Click()
''    Dim VchkCHILD_DEPEND As String
''    Dim ID               As Integer
''    Dim Sql              As String
''    Dim vtxtCHILD_FULL As String, vDTPCHILD As String
''
''    If chkDependent.Value Then VchkCHILD_DEPEND = N2Str2Null("YES") Else VchkCHILD_DEPEND = N2Str2Null("NO")
''    vtxtCHILD_FULL = N2Str2Null(txtCHILD_FULL)
' '   vDTPCHILD = N2Str2Null(DTPCHILD)
'
'    If Not txtCHILD_FULL.Text = "" Then
'        If SAVE_OR_EDIT_DEPENDENT_CHILD = "SAVE" Then                   'NEW
'            Call GenerateNewID("HRMS_APPLICANT_CHILD", ID)
'            CHILD_ENTRY_ID = ID
'
'            Sql = "Insert Into HRMS_APPLICANT_CHILD Values(" & _
'                    APPLICANT_ID & "," & _
'                    CHILD_ENTRY_ID & "," & _
'                    vtxtCHILD_FULL & "," & _
'                    vDTPCHILD & "," & _
'                    VchkCHILD_DEPEND & ")"
'        Else
'            Sql = "Update HRMS_APPLICANT_CHILD Set FullName = " & vtxtCHILD_FULL & _
'                    ",Birthdate = " & vDTPCHILD & _
'                    ",Dependent_Child = " & VchkCHILD_DEPEND & _
'                    " Where Applicant_ID = " & APPLICANT_ID & _
'                    " AND Entry_ID = " & CHILD_ENTRY_ID & ""
'        End If
'
'        gconDMIS.Execute (Sql)
'
'        Unload Me
''        frmAISApplications.DisplayChildInListView
'    End If
'End Sub

'Private Sub Form_Load()
'    Call CenterMe(frmMain, Me, 1)
'    If SAVE_OR_EDIT_DEPENDENT_CHILD = "SAVE" Then cmdCHILD_DELETE.Enabled = False
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISApplications.Enabled = True
    frmAISApplications.SetFocus
End Sub

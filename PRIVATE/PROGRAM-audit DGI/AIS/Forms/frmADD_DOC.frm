VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAISADD_DOC 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1680
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6300
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
   Moveable        =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6300
   Begin VB.ComboBox cboDOC 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1890
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   270
      Width           =   4305
   End
   Begin VB.PictureBox picCHILD_SAVE 
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
      Height          =   765
      Left            =   3780
      ScaleHeight     =   765
      ScaleWidth      =   2445
      TabIndex        =   4
      Top             =   780
      Width           =   2445
      Begin VB.CommandButton cmdCANCEL 
         Caption         =   "E&xit"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1620
         Picture         =   "frmADD_DOC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
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
         Height          =   765
         Left            =   840
         Picture         =   "frmADD_DOC.frx":0552
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Delete Entry"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdADD 
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
         Height          =   765
         Left            =   60
         Picture         =   "frmADD_DOC.frx":0C3C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Save Entry"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Document Name:"
      Height          =   240
      Index           =   6
      Left            =   150
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   210
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6375
      _Version        =   655364
      _ExtentX        =   11245
      _ExtentY        =   370
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
Attribute VB_Name = "frmAISADD_DOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function GenerateDocID() As Integer
    Dim RSTMP                                                         As ADODB.Recordset
    Dim ID                                                            As Integer

    Set RSTMP = gconDMIS.Execute("Select PaperID From HRMS_APPLICANT_PAPER Where Applicant_ID = " & APPLICANT_ID & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            ID = RSTMP!PaperID
            RSTMP.MoveNext
        Loop
    End If
    GenerateDocID = ID
End Function

Function CheckDocumentAlreadyPass() As Boolean
    Dim RSTMP                                                         As ADODB.Recordset
    Dim DOC_TYPE                                                      As String

    DOC_TYPE = N2Str2Null(cboDOC)
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_APPLICANT_PAPER Where Applicant_ID = " & APPLICANT_ID & _
                    " And PaperPass = " & DOC_TYPE & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckDocumentAlreadyPass = True
    Else
        CheckDocumentAlreadyPass = False
    End If
End Function

'Upating Code       : AXP-0707200711:13
Private Sub cmdAdd_Click()
    Dim ID                                                            As Integer
    Dim DOC_TYPE                                                      As String

    On Error GoTo Errorcode:

    DOC_TYPE = N2Str2Null(cboDOC)
    If Not cboDOC.Text = "" Then
        Call CheckIfDocumentAlreadySave
        frmMain.MousePointer = 11
        If SAVE_OR_EDIT_PAPERS = "SAVE" Then
            If CheckDocumentAlreadyPass = False Then
                ID = 0
                ID = GenerateDocID
                ID = ID + 1

                gconDMIS.Execute ("Insert Into HRMS_APPLICANT_PAPER Values(" & APPLICANT_ID & _
                                  "," & ID & "," & DOC_TYPE & ")")

                Call LogAudit("A", "ADD APPLICANT PASS DOC", APPLICANT_ID)

                Unload Me
                frmAISApplications.DisplayPapersInListView
            Else
                MsgBox "Document Type Already Pass", vbInformation, "Document Pass"
                On Error Resume Next
                cboDOC.SetFocus
            End If
        Else
            If CheckDocumentAlreadyPass = False Then
                ID = 0
                ID = GenerateDocID
                ID = ID + 1

                gconDMIS.Execute ("Update HRMS_APPLICANT_PAPER Set PaperPass = " & DOC_TYPE & _
                                " Where PaperID = " & PAPERS_ENTRY_ID & " And Applicant_ID = " & _
                                  APPLICANT_ID & "")

                Call LogAudit("E", "UPDATE APPLICANT PASS DOC", APPLICANT_ID)

                Unload Me
                frmAISApplications.DisplayPapersInListView
            Else
                MsgBox "Document Type Already Pass", vbInformation, "Document Pass"
                On Error Resume Next
                cboDOC.SetFocus
            End If
        End If
    Else
        MsgBox "Choose or Enter a Document Type", vbExclamation, "Document Pass"
        On Error Resume Next
        cboDOC.SetFocus
    End If
    frmMain.MousePointer = 0

    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub CheckIfDocumentAlreadySave()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim DOC_TYPE                                                      As String

    DOC_TYPE = N2Str2Null(cboDOC)

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_DOCUMENT Where DocumentType = " & DOC_TYPE & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
    Else
        gconDMIS.Execute ("Insert Into HRMS_DOCUMENT Values(" & DOC_TYPE & ")")
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:
    frmMain.MousePointer = 11

    If MsgBox("Remove This Document Pass", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
        gconDMIS.Execute ("Delete From HRMS_APPLICANT_PAPER Where Applicant_ID = " & APPLICANT_ID & _
                        " And PaperID = " & PAPERS_ENTRY_ID & "")

        Call LogAudit("X", "DELETE APPLICANT PASS DOC", APPLICANT_ID)
        Unload Me
        frmAISApplications.DisplayPapersInListView
    End If
    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    Call fillDocumentType
End Sub

Private Sub fillDocumentType()
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_DOCUMENT Order By DocumentType ASC")
    cboDOC.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboDOC.AddItem Null2String(RSTMP!DocumentType)
            RSTMP.MoveNext
        Loop
        cboDOC.ListIndex = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISApplications.Enabled = True
    frmAISApplications.picSaves.Visible = True
    On Error Resume Next
    frmAISApplications.SetFocus
End Sub


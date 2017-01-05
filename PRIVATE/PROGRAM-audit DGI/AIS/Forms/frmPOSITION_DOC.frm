VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAISPOSITION_DOC 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3045
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6405
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
   ScaleHeight     =   3045
   ScaleWidth      =   6405
   Begin VB.TextBox txtNOTE 
      Appearance      =   0  'Flat
      Height          =   1305
      Left            =   1980
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   4305
   End
   Begin VB.PictureBox picCHILD_SAVE 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   3900
      ScaleHeight     =   825
      ScaleWidth      =   2415
      TabIndex        =   5
      Top             =   2220
      Width           =   2415
      Begin VB.CommandButton cmdCANCEL 
         Caption         =   "E&xit"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1620
         Picture         =   "frmPOSITION_DOC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdDELETE 
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
         Height          =   795
         Left            =   840
         Picture         =   "frmPOSITION_DOC.frx":0552
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Delete Entry"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdADD 
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
         Height          =   795
         Left            =   60
         Picture         =   "frmPOSITION_DOC.frx":0C3C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Save Entry"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.ComboBox cboDOC 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1980
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   300
      Width           =   4305
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
      Height          =   240
      Index           =   7
      Left            =   1380
      TabIndex        =   7
      Top             =   690
      Width           =   465
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Document Name:"
      Height          =   240
      Index           =   6
      Left            =   180
      TabIndex        =   6
      Top             =   330
      Width           =   1695
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   210
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6405
      _Version        =   655364
      _ExtentX        =   11298
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
Attribute VB_Name = "frmAISPOSITION_DOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function GenerateDocID() As Integer
    Dim RSTMP                                                         As ADODB.Recordset
    Dim ID                                                            As Integer

    Set RSTMP = gconDMIS.Execute("Select EntryID From HRMS_POSITION_DOCUMENTS Where POS_ID = " & POSITION_ID & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            ID = RSTMP!Entryid
            RSTMP.MoveNext
        Loop
    End If
    GenerateDocID = ID
End Function

Function CheckDocumentAlreadyPass() As Boolean
    Dim RSTMP                                                         As ADODB.Recordset
    Dim DOC_TYPE                                                      As String

    DOC_TYPE = N2Str2Null(cboDOC)
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_POSITION_DOCUMENTS Where POS_ID = " & POSITION_ID & _
                    " And dOCUMENTtYPE = " & DOC_TYPE & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckDocumentAlreadyPass = True
    Else
        CheckDocumentAlreadyPass = False
    End If
End Function

Private Sub cmdAdd_Click()
    Dim ID                                                            As Integer
    Dim DOC_TYPE As String, vtxtNOTE                                  As String

    On Error GoTo Errorcode:

    DOC_TYPE = N2Str2Null(cboDOC)
    vtxtNOTE = N2Str2Null(txtNOTE)
    ID = 0

    frmMain.MousePointer = 11
    If Not cboDOC.Text = "" Then
        Call CheckIfDocumentAlreadySave
        If SAVE_OR_EDIT_PAPERS = "SAVE" Then
            If CheckDocumentAlreadyPass = False Then
                ID = GenerateDocID
                ID = ID + 1

                gconDMIS.Execute ("Insert Into HRMS_POSITION_DOCUMENTS Values(" & POSITION_ID & _
                                  "," & ID & _
                                  "," & DOC_TYPE & _
                                  "," & vtxtNOTE & ")")
                Unload Me
                frmAISPOSITION.DisplayPOSITION_DOCUMENT
            Else
                MsgBox "Document Type Already Pass", vbInformation, "Document Pass"
                On Error Resume Next
                cboDOC.SetFocus
            End If
        Else
            Call CheckIfDocumentAlreadySave
            gconDMIS.Execute ("Update HRMS_POSITION_DOCUMENTS Set DocumentType = " & DOC_TYPE & _
                              ",Notes = " & vtxtNOTE & _
                            " Where POS_ID = " & POSITION_ID & _
                            " And EntryID = " & POSITION_DOC_ENTRY_ID & "")

            Unload Me
            frmAISPOSITION.DisplayPOSITION_DOCUMENT
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
    On Error GoTo Errorcode:

    If MsgBox("Remove This Document Pass", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
        gconDMIS.Execute ("Delete From HRMS_POSITION_DOCUMENTS Where POS_ID = " & POSITION_ID & _
                        " And EntryID = " & POSITION_DOC_ENTRY_ID & "")

        Unload Me
        frmAISPOSITION.DisplayPOSITION_DOCUMENT
    End If
    Exit Sub
Errorcode:
    ShowVBError
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
    End If
    cboDOC.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISPOSITION.picSave.Visible = True
    frmAISPOSITION.Enabled = True
    On Error Resume Next
    frmAISPOSITION.SetFocus
End Sub


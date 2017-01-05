VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAISPOSITION_EDUC 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2925
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7680
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
   ScaleHeight     =   2925
   ScaleWidth      =   7680
   Begin VB.ComboBox cboFIELDS 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   750
      Width           =   6195
   End
   Begin VB.TextBox txtNOTE 
      Appearance      =   0  'Flat
      Height          =   795
      Left            =   1350
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   6195
   End
   Begin VB.ComboBox cboDEGREE 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   6195
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
      Height          =   825
      Left            =   5100
      ScaleHeight     =   825
      ScaleWidth      =   2415
      TabIndex        =   6
      Top             =   2100
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
         Picture         =   "frmPOSITION_EDUC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmPOSITION_EDUC.frx":0552
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Delete Entry"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdSAVE 
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
         Picture         =   "frmPOSITION_EDUC.frx":0C3C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Save Entry"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Study Fields"
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Degree"
      Height          =   240
      Index           =   20
      Left            =   570
      TabIndex        =   8
      Top             =   360
      Width           =   690
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
      Height          =   240
      Index           =   7
      Left            =   780
      TabIndex        =   7
      Top             =   1170
      Width           =   465
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   225
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7695
      _Version        =   655364
      _ExtentX        =   13573
      _ExtentY        =   397
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
Attribute VB_Name = "frmAISPOSITION_EDUC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function GenerateEntryID() As Integer
    Dim RSTMP                                                         As ADODB.Recordset
    Dim ID                                                            As Integer

    Set RSTMP = gconDMIS.Execute("Select EntryID From HRMS_POSITION_EDUCATION Where POS_ID = " & POSITION_ID & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            ID = RSTMP!Entryid
            RSTMP.MoveNext
        Loop
    End If

    GenerateEntryID = ID
End Function

Function CheckIfDegreeAlreadySave() As Boolean
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select Degree From HRMS_POSITION_EDUCATION Where Degree = '" & cboDEGREE & _
                      "' And Pos_ID = " & POSITION_ID & "")

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckIfDegreeAlreadySave = True
    Else
        CheckIfDegreeAlreadySave = False
    End If
End Function

Sub FillCboFIELDS()
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select Fields From HRMS_FIELDS Order By Fields ASC")
    cboFIELDS.Clear
    cboFIELDS.AddItem "-"
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboFIELDS.AddItem RSTMP!FIELDS
            RSTMP.MoveNext
        Loop
    End If
    cboFIELDS.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:

    frmMain.MousePointer = 11
    If MsgBox("Remove This Document Pass", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
        gconDMIS.Execute ("Delete From HRMS_POSITION_EDUCATION Where POS_ID = " & POSITION_ID & _
                        " And EntryID = " & POSITION_EDU_ENTRY_ID & "")

        Unload Me
        frmAISPOSITION.DisplayPOSITION_EDUCATION
    End If
    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    Dim vcboDEGREE As String, vtxtNOTE As String, vcboFIELDS          As String
    Dim ID                                                            As Integer

    On Error GoTo Errorcode:
    frmMain.MousePointer = 11

    ID = 0
    vcboDEGREE = N2Str2Null(cboDEGREE)
    vtxtNOTE = N2Str2Null(txtNOTE)
    vcboFIELDS = N2Str2Null(cboFIELDS)

    If cboDEGREE.Text = "High School Diploma" Then
        If CheckIfDegreeAlreadySave = True Then
            MsgBox "High School Degree Cannot be Duplicate", vbInformation, "Position"
            On Error Resume Next
            cboDEGREE.SetFocus
            Exit Sub
        End If
    End If

    If POSITION_SAVE_OR_EDIT_EDU = "SAVE" Then
        ID = GenerateEntryID()
        ID = ID + 1

        gconDMIS.Execute ("Insert Into HRMS_POSITION_EDUCATION Values(" & POSITION_ID & _
                          "," & ID & _
                          "," & vcboDEGREE & _
                          "," & vcboFIELDS & _
                          "," & vtxtNOTE & ")")
        Unload Me
        frmAISPOSITION.DisplayPOSITION_EDUCATION
    Else
        gconDMIS.Execute ("Update HRMS_POSITION_EDUCATION Set Degree = " & vcboDEGREE & _
                          ",Fields = " & vcboFIELDS & _
                          ",Notes = " & vtxtNOTE & _
                        " Where EntryID = " & POSITION_EDU_ENTRY_ID & "")

        Unload Me
        frmAISPOSITION.DisplayPOSITION_EDUCATION
    End If
    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    frmMain.MousePointer = 11

    Call FillEducationalDegree
    Call FillCboFIELDS

    frmMain.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISPOSITION.picSave.Visible = True
    frmAISPOSITION.Enabled = True
    On Error Resume Next
    frmAISPOSITION.SetFocus
End Sub

Private Sub FillEducationalDegree()
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_DEGREE Order By Degree ASC")
    cboDEGREE.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboDEGREE.AddItem Null2String(RSTMP!DEGREE)
            RSTMP.MoveNext
        Loop
    End If
    cboDEGREE.ListIndex = 0
End Sub


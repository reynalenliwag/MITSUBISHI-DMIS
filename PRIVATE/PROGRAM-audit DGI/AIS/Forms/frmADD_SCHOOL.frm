VERSION 5.00
Begin VB.Form frmAISADD_SCHOOL 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4440
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
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
   ScaleHeight     =   4440
   ScaleWidth      =   6645
   Begin VB.ComboBox cboFIELDS 
      Height          =   330
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1170
      Width           =   4515
   End
   Begin VB.TextBox txtSCHOOL_GRADE 
      Height          =   345
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1650
      Width           =   1515
   End
   Begin VB.TextBox txtSCHOOL_TYEAR 
      Height          =   345
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   6
      Top             =   3060
      Width           =   1515
   End
   Begin VB.ComboBox cboEDUC 
      Height          =   330
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   4515
   End
   Begin VB.TextBox txtSCHOOL_NAME 
      Height          =   345
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   1
      Top             =   660
      Width           =   4515
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
      Height          =   735
      Left            =   2670
      ScaleHeight     =   675
      ScaleWidth      =   3705
      TabIndex        =   13
      Top             =   3570
      Width           =   3765
      Begin VB.CommandButton cmdSCHOOL_DELETE 
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
         Height          =   675
         Left            =   1260
         TabIndex        =   15
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSCHOOL_CANCEL 
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
         Height          =   675
         Left            =   2490
         TabIndex        =   8
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSCHOOL_SAVE 
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
         Height          =   675
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.TextBox txtSCHOOL_ADD 
      Height          =   345
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2100
      Width           =   4485
   End
   Begin VB.TextBox txtSCHOOL_FYEAR 
      Height          =   345
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2580
      Width           =   1515
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(yyyy)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   3570
      TabIndex        =   19
      Top             =   3150
      Width           =   660
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(yyyy)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   3540
      TabIndex        =   18
      Top             =   2670
      Width           =   660
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grade"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   1230
      TabIndex        =   17
      Top             =   1770
      Width           =   570
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Study Fields"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   600
      TabIndex        =   16
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1020
      TabIndex        =   14
      Top             =   3150
      Width           =   780
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Educ. Attainment"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   12
      Top             =   270
      Width           =   1725
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   20
      Left            =   810
      TabIndex        =   11
      Top             =   2670
      Width           =   1005
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   23
      Left            =   300
      TabIndex        =   10
      Top             =   2220
      Width           =   1515
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   28
      Left            =   540
      TabIndex        =   9
      Top             =   750
      Width           =   1275
   End
End
Attribute VB_Name = "frmAISADD_SCHOOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSCHOOL_CANCEL_Click()
    Unload Me
End Sub

Private Sub cmdSCHOOL_DELETE_Click()
    If MsgBox("Are You Sure", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Educational Type Attend") = vbYes Then
        gconDMIS.Execute ("Delete From HRMS_APPLICANT_EDUC Where Applicant_ID = " & _
                           APPLICANT_ID & " And Entry_ID = " & SCHOOL_ENTRY_ID & "")

        Unload Me
        Call frmAISApplications.DisplayEDUCInListView
    End If
End Sub

Function CheckIfCompleteEducationalEntry() As Boolean
    If txtSCHOOL_NAME.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Educational Background"
        txtSCHOOL_NAME.SetFocus
        CheckIfCompleteEducationalEntry = True
        Exit Function
    End If
    If txtSCHOOL_GRADE.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Educational Background"
        txtSCHOOL_GRADE.SetFocus
        CheckIfCompleteEducationalEntry = True
        Exit Function
    End If
    If IsNumeric(txtSCHOOL_GRADE.Text) = False Then
        MsgBox "Invalid Entry", vbExclamation, "Educational Background"
        txtSCHOOL_GRADE.SetFocus
        CheckIfCompleteEducationalEntry = True
        Exit Function
    End If
    If txtSCHOOL_FYEAR.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Educational Background"
        txtSCHOOL_FYEAR.SetFocus
        CheckIfCompleteEducationalEntry = True
        Exit Function
    End If
    If IsNumeric(txtSCHOOL_FYEAR.Text) = False Then
        MsgBox "Invalid Entry", vbExclamation, "Educational Background"
        txtSCHOOL_FYEAR.SetFocus
        CheckIfCompleteEducationalEntry = True
        Exit Function
    End If
    If txtSCHOOL_TYEAR.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Educational Background"
        txtSCHOOL_TYEAR.SetFocus
        CheckIfCompleteEducationalEntry = True
        Exit Function
    End If
    If IsNumeric(txtSCHOOL_TYEAR.Text) = False Then
        MsgBox "Invalid Entry", vbExclamation, "Educational Background"
        txtSCHOOL_TYEAR.SetFocus
        CheckIfCompleteEducationalEntry = True
        Exit Function
    End If
End Function

Private Sub cmdSCHOOL_SAVE_Click()
    Dim ID                  As Integer
    Dim Sql, vcboEDUC       As String, VtxtEDU_SFYEAR   As String
    Dim VtxtEDU_SNAME       As String, VtxtEDU_SADD     As String
    Dim VtxtEDU_STYEAR      As String, vcboEDU_FIELDS   As String
    Dim rsTmp               As ADODB.Recordset
    Dim vtxtGRADE           As Integer
    
    If CheckIfCompleteEducationalEntry = True Then
        Exit Sub
    End If
    
    vcboEDUC = N2Str2Null(cboEDUC)
    VtxtEDU_SNAME = N2Str2Null(txtSCHOOL_NAME)
    VtxtEDU_SADD = N2Str2Null(txtSCHOOL_ADD)
    VtxtEDU_SFYEAR = N2Str2Null(txtSCHOOL_FYEAR)
    VtxtEDU_STYEAR = N2Str2Null(txtSCHOOL_TYEAR)
    vcboEDU_FIELDS = N2Str2Null(cboFIELDS)
    vtxtGRADE = CInt(txtSCHOOL_GRADE)

    If Not cboEDUC.Text = "" Then
        If SAVE_OR_EDIT_SCHOOL = "SAVE" Then                  'NEW
            Call GenerateNewID("HRMS_APPLICANT_EDUC", ID)
            SCHOOL_ENTRY_ID = ID
            
            If cboEDUC.Text = "High School Diploma" Then
                If CheckIfApplicantAlrwadyGotHSandELEm = True Then
                    MsgBox "Educational Degree Already on The List", vbInformation, "Education Background"
                    cboEDUC.SetFocus
                    Exit Sub
                End If
            End If
            
            Sql = "Insert Into HRMS_APPLICANT_EDUC Values(" & _
                    APPLICANT_ID & "," & _
                    SCHOOL_ENTRY_ID & "," & _
                    vcboEDUC & "," & _
                    vcboEDU_FIELDS & "," & _
                    vtxtGRADE & "," & _
                    VtxtEDU_SNAME & "," & _
                    VtxtEDU_SADD & "," & _
                    VtxtEDU_SFYEAR & "," & _
                    VtxtEDU_STYEAR & ")"
        Else
            Set rsTmp = GetRS("Select * From HRMS_APPLICANT_EDUC Where ID = " & APPLICANT_ID & _
                                " And Entry_ID = " & SCHOOL_ENTRY_ID & "")
            If Not (rsTmp.BOF And rsTmp.EOF) Then
                If Not rsTmp!SchoolType = cboEDUC.Text Then
                    If cboEDUC.Text = "ELEM" Or cboEDUC.Text = "HS" Then
                        If CheckIfApplicantAlrwadyGotHSandELEm = True Then
                            MsgBox "Educational Degree Already on The List", vbInformation, "Education Background"
                            cboEDUC.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            End If
                                
            Sql = "Update HRMS_APPLICANT_EDUC Set SchoolType = " & vcboEDUC & _
                    ",StudyFields  = " & vcboEDU_FIELDS & _
                    ",SchoolName = " & VtxtEDU_SNAME & _
                    ",SchoolADD = " & VtxtEDU_SADD & _
                    ",Grade = " & vtxtGRADE & _
                    ",FYear = " & VtxtEDU_SFYEAR & _
                    ",TYear = " & VtxtEDU_STYEAR & _
                    " Where Applicant_ID = " & APPLICANT_ID & _
                    " AND Entry_ID = " & SCHOOL_ENTRY_ID & ""
        End If

        gconDMIS.Execute (Sql)

        Unload Me
        frmAISApplications.DisplayEDUCInListView
    Else
        MsgBox "Choose a Degree", vbExclamation, "Education Background"
        cboEDUC.SetFocus
    End If
End Sub

Function CheckIfApplicantAlrwadyGotHSandELEm() As Boolean
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetRS("Select * From HRMS_APPLICANT_EDUC Where Applicant_ID = " & APPLICANT_ID & " And SchoolType = '" & _
                        Trim(cboEDUC) & "'")
    If Not (rsTmp.EOF And rsTmp.BOF) Then
        CheckIfApplicantAlrwadyGotHSandELEm = True
    End If
End Function

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    If SAVE_OR_EDIT_SCHOOL = "SAVE" Then cmdSCHOOL_DELETE.Enabled = False
    Call FillEducationalDegree
    Call FillStudyFields
End Sub

Private Sub FillStudyFields()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetRS("Select * From HRMS_FIELDS Order By Fields ASC")
    cboFIELDS.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            cboFIELDS.AddItem rsTmp!Fields
            rsTmp.MoveNext
        Loop
    End If
    cboFIELDS.ListIndex = 0
End Sub

Private Sub FillEducationalDegree()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetRS("Select * From HRMS_DEGREE Order By Degree ASC")
    cboEDUC.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            cboEDUC.AddItem Null2String(rsTmp!Degree)
            rsTmp.MoveNext
        Loop
    End If
    cboEDUC.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISApplications.Enabled = True
    frmAISApplications.SetFocus
End Sub

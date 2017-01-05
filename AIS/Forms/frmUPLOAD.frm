VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAISUPLOAD 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload Applicant"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
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
   ScaleHeight     =   4665
   ScaleWidth      =   9555
   Begin VB.Frame Frame1 
      Caption         =   "Applicant Information"
      Height          =   4455
      Left            =   90
      TabIndex        =   13
      Top             =   90
      Width           =   3195
      Begin VB.TextBox txtSEARCH 
         Height          =   375
         Left            =   60
         TabIndex        =   0
         Top             =   390
         Width           =   3045
      End
      Begin MSComctlLib.ListView lsvAPP 
         Height          =   3495
         Left            =   90
         TabIndex        =   1
         Top             =   870
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   6165
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Full Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   9
         EndProperty
      End
   End
   Begin VB.CommandButton cmdUPDATE 
      Caption         =   "&UPLOAD"
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
      Left            =   7470
      Picture         =   "frmUPLOAD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdEXIT 
      Caption         =   "E&XIT"
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
      Left            =   8460
      Picture         =   "frmUPLOAD.frx":0731
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Applicant Information"
      Height          =   2295
      Left            =   3390
      TabIndex        =   8
      Top             =   90
      Width           =   6045
      Begin VB.ComboBox cboPOSITION 
         Height          =   360
         Left            =   1845
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1290
         Width           =   2805
      End
      Begin VB.ComboBox cboTYPE 
         Height          =   360
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1770
         Width           =   2805
      End
      Begin VB.Label lblCAP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         Height          =   240
         Index           =   2
         Left            =   975
         TabIndex        =   12
         Top             =   1380
         Width           =   765
      End
      Begin VB.Label lblCAP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Type"
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   11
         Top             =   1860
         Width           =   1500
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   240
         Index           =   0
         Left            =   810
         TabIndex        =   10
         Top             =   930
         Width           =   945
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applicant No."
         Height          =   240
         Index           =   6
         Left            =   420
         TabIndex        =   9
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label lblINFO 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1860
         TabIndex        =   6
         Top             =   420
         Width           =   1125
      End
      Begin VB.Label lblINFO 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1860
         TabIndex        =   7
         Top             =   870
         Width           =   4005
      End
   End
End
Attribute VB_Name = "frmAISUPLOAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEXIT_Click()
    Unload Me
End Sub

Private Sub cmdUPDATE_Click()
    frmAISUPLOAD.Enabled = False
    frmAISPOSITION_APPLY.Show

'    Dim rsTmp As ADODB.Recordset
'    Dim ETYPE As String
'
'    If Not cboTYPE.Text = "" Then
'        If MsgBox("Upload Applicant", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
'            If cboTYPE.Text = "Contractual" Then ETYPE = N2Str2Null("C")
'            If cboTYPE.Text = "Allowance Base" Then ETYPE = N2Str2Null("A")
'            If cboTYPE.Text = "Probationary" Then ETYPE = N2Str2Null("E")
'
'            Set rsINT = gconDMIS.Execute("Select * From HRMS_APPLICANT_INTERVIEW_SCHEDULE Where Remarks = '" & _
'                "Passed" & "' And Applicant_ID = " & rsTmp!APPLICANT_ID & "")
'            If Not (rsINT.BOF And rsINT.EOF) Then
'                If MsgBox("Upload Applicant", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
'                    gconDMIS.Execute ("Update HRMS_APPLICANT_PERSONAL Set Type = " & ETYPE & " And Hired = '" & "YES" & "' Where ID = " & APPLICANT_ID & "")
'                End If
'            Else
'                If MsgBox("Applicant Not Yet Pass the Interview", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
'                    gconDMIS.Execute ("Update HRMS_APPLICANT_PERSONAL Set Type = " & ETYPE & " And Hired = '" & "YES" & "' Where ID = " & APPLICANT_ID & "")
'                End If
'            End If
'        End If
'    Else
'        MsgBox "Choose a Employee Type", vbExclamation, "Upload Applicant"
'        cboTYPE.SetFocus
'    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3:
            txtSEARCH.Text = ""
            txtSEARCH.SetFocus
    
    End Select
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    Call FillCBOType
    Call FillCboPosition
End Sub

Sub FillCboPosition()
    Dim rsTmp As ADODB.Recordset
    Dim SZERO As String
    
    Set rsTmp = gconDMIS.Execute("Select * From HRMS_POSITION Where PositionAvailable > PositionTaken And DateInactive >= '" & Date & "' Order By PositionDesc ASC")
    cboPOSITION.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            If Len(rsTmp!POS_ID) = 1 Then SZERO = "0"
            
            cboPOSITION.AddItem rsTmp!PositionDesc & " - " & SZERO & rsTmp!POS_ID
            rsTmp.MoveNext
        Loop
    End If
    cboPOSITION.ListIndex = 0
End Sub

Sub FillCBOType()
    cboTYPE.AddItem "Contractual"
    cboTYPE.ItemData(cboTYPE.NewIndex) = 0
    cboTYPE.AddItem "Allowance Base"
    cboTYPE.ItemData(cboTYPE.NewIndex) = 1
    cboTYPE.AddItem "Probationary"
    cboTYPE.ItemData(cboTYPE.NewIndex) = 2
    cboTYPE.ListIndex = 0
End Sub

Private Sub LsvAPP_DblClick()
    Dim Index As Long
    
    If Not lsvAPP.ListItems.Count = 0 Then
        Index = lsvAPP.SelectedItem.Index
        With lsvAPP
            lblINFO(0).Caption = .ListItems(Index).SubItems(1)
            lblINFO(1).Caption = .ListItems(Index).Text
            
            cboPOSITION.SetFocus
        End With
    End If
End Sub

Private Sub txtSEARCH_Change()
    Dim rsTmp As ADODB.Recordset, rsINT As ADODB.Recordset
    Dim Keyword As String
    Dim ITEM As ListItem
    
    Keyword = Trim(txtSEARCH.Text)
    
    Set rsTmp = gconDMIS.Execute("Select LastName,FirstName,Applicant_ID From HRMS_APPLICANT_PERSONAL Where LastName Like '%" & Keyword & "%' Or FirstName Like '%" & Keyword & _
        "%' And Hired = '" & "NO" & "' Order by Applicant_ID ASC")
    lsvAPP.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF

                Set ITEM = lsvAPP.ListItems.Add(, , rsTmp!LastName & ", " & rsTmp!FirstName)
                ITEM.SubItems(1) = rsTmp!APPLICANT_ID
            'End If
            
            'Set rsINT = Nothing
            rsTmp.MoveNext
        Loop
    Else
        lsvAPP.ListItems.Clear
    End If
    
    If txtSEARCH.Text = "" Then lsvAPP.ListItems.Clear
End Sub

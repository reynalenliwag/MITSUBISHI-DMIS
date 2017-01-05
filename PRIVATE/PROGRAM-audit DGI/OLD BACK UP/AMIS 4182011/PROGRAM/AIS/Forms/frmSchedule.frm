VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAISSchedule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scheduling of Interview"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   1140
   ClientWidth     =   12735
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
   ScaleHeight     =   5775
   ScaleWidth      =   12735
   Begin VB.CommandButton cmdUPDATE 
      Caption         =   "UPDATE INTERVIEW"
      Height          =   615
      Left            =   930
      TabIndex        =   10
      Top             =   5010
      Width           =   1395
   End
   Begin VB.CommandButton cmdEXIT 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   2400
      TabIndex        =   11
      Top             =   5010
      Width           =   1395
   End
   Begin VB.Frame fmeCHOOSE 
      Caption         =   "CHOOSE"
      Height          =   4425
      Left            =   60
      TabIndex        =   16
      Top             =   90
      Width           =   3735
      Begin VB.ComboBox cboPOSITION 
         Height          =   360
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   3435
      End
      Begin MSComctlLib.ListView lsvAPPLICANT 
         Height          =   3465
         Left            =   150
         TabIndex        =   9
         Top             =   810
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   6112
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
            Text            =   "ID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin VB.Frame fmeSCHED 
      Caption         =   "SCHEDULE"
      Enabled         =   0   'False
      Height          =   5535
      Left            =   3870
      TabIndex        =   12
      Top             =   90
      Width           =   8715
      Begin VB.CommandButton cmdDELETE 
         Caption         =   "DELETE"
         Enabled         =   0   'False
         Height          =   555
         Left            =   7200
         TabIndex        =   7
         Top             =   4860
         Width           =   1395
      End
      Begin VB.CommandButton cmdCANCEL 
         Caption         =   "CANCEL"
         Height          =   435
         Left            =   6900
         TabIndex        =   5
         Top             =   1650
         Width           =   1575
      End
      Begin VB.CommandButton cmdSET 
         Caption         =   "SET SCHEDULE"
         Height          =   435
         Left            =   4560
         TabIndex        =   4
         Top             =   1650
         Width           =   2205
      End
      Begin VB.PictureBox picAPP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   6900
         ScaleHeight     =   1245
         ScaleWidth      =   1515
         TabIndex        =   15
         Top             =   240
         Width           =   1545
         Begin VB.Image imgAPP 
            Height          =   1185
            Left            =   30
            Stretch         =   -1  'True
            Top             =   30
            Width           =   1455
         End
      End
      Begin VB.ComboBox cboTIME 
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1710
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPDATE 
         Height          =   345
         Left            =   150
         TabIndex        =   2
         Top             =   1710
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   609
         _Version        =   393216
         Format          =   3801089
         CurrentDate     =   39132
      End
      Begin MSComctlLib.ListView lsvSCHED 
         Height          =   2505
         Left            =   150
         TabIndex        =   6
         Top             =   2220
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   4419
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Time"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblAPPID 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   180
         TabIndex        =   17
         Top             =   4980
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblINFO 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   1
         Left            =   1500
         TabIndex        =   1
         Top             =   1080
         Width           =   5070
      End
      Begin VB.Label lblINFO 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   0
         Left            =   1500
         TabIndex        =   0
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label lblCAP 
         Caption         =   "Full Name"
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   14
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Applicant ID"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   720
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmAISSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboPOSITION_Change()
    Dim rsTmp As ADODB.Recordset, rsPER As ADODB.Recordset
    Dim ITEM As ListItem
    
    Set rsTmp = GetRS("Select * From HRMS_APPLICANT_HISTORY Where PositionID = " & Right(cboPOSITION, 3) & "")
    lsvAPPLICANT.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            '----------------------------------------------------------------------------------------
            Set rsPER = GetRS("Select LastName, FirstName From HRMS_APPLICANT_PERSONAL Where Applicant_ID = " & _
                rsTmp!APPLICANT_ID & " And Hired = '" & "NO" & "'")
            If Not (rsPER.BOF And rsPER.EOF) Then
                Set ITEM = lsvAPPLICANT.ListItems.Add(, , rsTmp!APPLICANT_ID)
                ITEM.SubItems(1) = rsPER!LastName & "," & rsPER!FirstName
            End If
            '----------------------------------------------------------------------------------------
            rsTmp.MoveNext
        Loop
    End If
End Sub

Private Sub cboPOSITION_Click()
    Call cboPOSITION_Change
End Sub

Private Sub EnableFrame(COND As Boolean)
    FmeSCHED.Enabled = COND
    fmeCHOOSE.Enabled = Not COND
End Sub

Private Sub CleanApplicantInfo()
    lblINFO(0).Caption = ""
    lblINFO(1).Caption = ""
End Sub

Private Sub cmdCancel_Click()
    cmdDELETE.Enabled = False
    lsvSCHED.ListItems.Clear
    
    Call EnableFrame(False)
    Call CleanApplicantInfo
    
    lsvAPPLICANT.SetFocus
End Sub

Private Sub cmdDELETE_Click()
    Dim rsTmp As ADODB.Recordset
    
    If MsgBox("Delete Interview", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
        '-----------------------------------------------------------------------------------------------
        Set rsTmp = GetRS("Select Int_ID From HRMS_INTERVIEW_SCHED Where DateINT = '" & CDate(dtpDATE) & _
                        "' And POS_ID = " & CInt(Right(cboPOSITION, 3)) & "")
        If Not (rsTmp.BOF And rsTmp.EOF) Then
            gconDMIS.Execute ("Delete From HRMS_APPLICANT_INTERVIEW_SCHED Where Applicant_ID = " & Val(lblAPPID.Caption) & _
                " And Int_ID = " & rsTmp!Int_ID & "")
            
            Call DTPDATE_Change
        End If
        '-----------------------------------------------------------------------------------------------
        cmdDELETE.Enabled = False
    Else
        lsvSCHED.SetFocus
    End If
End Sub

Private Sub cmdEXIT_Click()
    Unload Me
End Sub

Private Sub cmdSET_Click()
    Dim SCHED_ID As Integer
    Dim rsTmp As ADODB.Recordset, rsTIME As ADODB.Recordset, rsSCHED As ADODB.Recordset
    Dim rsALL As ADODB.Recordset
    
    Set rsTmp = GetRS("Select * From HRMS_INTERVIEW_SCHED Where POS_ID = " & CInt(Right(cboPOSITION, 3)) & _
        " And DateINT = '" & CDate(dtpDATE) & "'")
    
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        SCHED_ID = rsTmp!Int_ID
    Else
        Set rsSCHED = GetRS("Select INT_ID From HRMS_INTERVIEW_SCHED Order By Int_ID ASC")
        If Not (rsSCHED.BOF And rsSCHED.EOF) Then
            Do While Not rsSCHED.EOF
                SCHED_ID = rsSCHED!Int_ID
                rsSCHED.MoveNext
            Loop
        End If
        SCHED_ID = SCHED_ID + 1
        
        gconDMIS.Execute ("Insert Into HRMS_INTERVIEW_SCHED Values(" & SCHED_ID & _
                            "," & Right(cboPOSITION, 3) & ",'" & CDate(dtpDATE) & "')")
    End If
           
           
    Set rsALL = GetRS("Select * From HRMS_APPLIcANT_INTERVIEW_SCHED Where Applicant_ID = " & Val(lblINFO(0)) & _
                    " And Int_ID = " & SCHED_ID & "")
                
    If Not (rsALL.BOF And rsALL.EOF) Then
        MsgBox "Applicant Already Schedule in this Day", vbInformation, "Schedule of Interview"
        dtpDATE.SetFocus
        Exit Sub
    Else
        Set rsTIME = GetRS("Select * From HRMS_APPLICANT_INTERVIEW_SCHED Where TIMEID = " & _
                        CInt(cboTIME.ListIndex) + 1 & " And Int_ID = " & SCHED_ID & "")
        If Not (rsTIME.BOF And rsTIME.EOF) Then
            MsgBox "Interview Time Already Occupied", vbInformation, "Schedule of Interview"
            cboTIME.SetFocus
            Exit Sub
        Else
            gconDMIS.Execute ("Insert Into HRMS_APPLICANT_INTERVIEW_SCHED Values(" & Val(lblINFO(0)) & _
                "," & SCHED_ID & "," & CInt(cboTIME.ListIndex) + 1 & ",'" & "" & "')")
        End If
    End If
    
    
    cmdDELETE.Enabled = False
    Call EnableFrame(False)
    
    Call DTPDATE_Change
    lsvAPPLICANT.SetFocus
End Sub

Private Sub cmdUPDATE_Click()
    frmAISSchedule.Enabled = False
    frmAISSchedule_EDIT.Show
End Sub

Private Sub DTPDATE_Change()
    '''''''CODE HERE FILTER THE INTERVIEW ON THIS DATE BY KIND OF POSITION.....
    Dim rsTmp As ADODB.Recordset, rsINT As ADODB.Recordset, rsPER As ADODB.Recordset, rsTIME As ADODB.Recordset
    Dim DATEINT As String
    Dim ITEM As ListItem
    
    DATEINT = Trim(dtpDATE)
    'A==============================================================================================================
    Set rsTmp = GetRS("Select * From HRMS_INTERVIEW_SCHED Where DATEINT = '" & DATEINT & "' And POS_ID = " & _
        CInt(Right(cboPOSITION, 3)) & "")
    
    lsvSCHED.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            'B======================================================================================================
            Set rsINT = GetRS("Select * From HRMS_APPLICANT_INTERVIEW_SCHED Where INT_ID = " & rsTmp!Int_ID & "")
            If Not (rsINT.BOF And rsINT.EOF) Then
                Do While Not rsINT.EOF
                    'C==================================================================================================
                    Set rsPER = GetRS("Select FirstName,LastName,Applicant_ID From HRMS_APPLICANT_PERSONAL Where Applicant_ID = " & _
                        rsINT!APPLICANT_ID & "")
                    If Not (rsPER.BOF And rsPER.EOF) Then
                        Set ITEM = lsvSCHED.ListItems.Add(, , rsPER!APPLICANT_ID)
                        ITEM.SubItems(1) = rsPER!LastName & "," & rsPER!FirstName
                        ITEM.SubItems(2) = rsTmp!DATEINT
                    
                        Set rsTIME = GetRS("Select * From HRMS_INTERVIEW_SCHED_TIME Where TIMEID = " & rsINT!TIMEID & "")
                        If Not (rsTIME.BOF And rsTIME.EOF) Then
                            ITEM.SubItems(3) = rsTIME!TimeINT
                        End If
                    End If
                    'C==================================================================================================
                    rsINT.MoveNext
                Loop
            End If
            'B======================================================================================================
            rsTmp.MoveNext
        Loop
    End If
    'A==============================================================================================================
End Sub

Private Sub DTPDATE_Click()
    Call DTPDATE_Change
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    Call FillCboPosition
    Call FillCBOTime(cboTIME)
End Sub

Private Sub FillCboPosition()
    Dim rsTmp As ADODB.Recordset
    Dim SZERO As String
    
    Set rsTmp = GetRS("Select * From HRMS_POSITION Order By PositionDesc ASC")
    
    cboPOSITION.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            If Len(rsTmp!POS_ID) = 1 Then SZERO = "00"
            If Len(rsTmp!POS_ID) = 2 Then SZERO = "0"
            
            cboPOSITION.AddItem rsTmp!PositionDesc & " - " & SZERO & rsTmp!POS_ID
            
            rsTmp.MoveNext
        Loop
    End If
    cboPOSITION.ListIndex = 0
End Sub

Private Sub lsvAPPLICANT_Click()
    Dim rsTmp As ADODB.Recordset
    Dim INDEX As Long
    
    If Not lsvAPPLICANT.ListItems.Count = 0 Then
        INDEX = CLng(lsvAPPLICANT.SelectedItem.INDEX)
        With lsvAPPLICANT
            lblINFO(0).Caption = .ListItems(INDEX).Text
            lblINFO(1).Caption = .ListItems(INDEX).SubItems(1)
            
            Set rsTmp = GetRS("Select * From HRMS_APPLICANT_IMAGE_LOCATION Where Applicant_ID = " & _
                CInt(lblINFO(0).Caption) & "")
            If Not (rsTmp.BOF And rsTmp.EOF) Then
                If Null2String(rsTmp!ImageLocation) <> "" Then
                    On Error Resume Next
                    LoadPic imgAPP, Null2String(rsTmp!ImageLocation)
                Else
                    LoadPic imgAPP, ""
                End If
            Else
                LoadPic imgAPP, ""
            End If
        End With
    End If
End Sub

Private Sub lsvAPPLICANT_DblClick()
    Dim INDEX As Integer
    If Not lsvAPPLICANT.ListItems.Count = 0 Then
        INDEX = lsvAPPLICANT.SelectedItem.INDEX
        With lsvAPPLICANT
            lblINFO(0).Caption = .ListItems(INDEX).Text
            lblINFO(1).Caption = .ListItems(INDEX).SubItems(1)
            
            fmeCHOOSE.Enabled = False
            FmeSCHED.Enabled = True
            dtpDATE.SetFocus
            
            Call DTPDATE_Change
        End With
    End If
End Sub

Private Sub lsvSCHED_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    Dim INDEX As Integer
    
    If Not lsvSCHED.ListItems.Count = 0 Then
        cmdDELETE.Enabled = True
        INDEX = CInt(lsvSCHED.SelectedItem.INDEX)
        With lsvSCHED
            lblAPPID.Caption = lsvSCHED.ListItems(INDEX).Text
            
        End With
    End If
End Sub

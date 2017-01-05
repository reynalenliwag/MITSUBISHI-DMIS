VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSMIS_FILE_jobMasterFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Master File"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSMIS_FILE_jobMasterFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetails 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   5
      Top             =   1110
      Width           =   5715
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   90
         MaxLength       =   35
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   570
         Width           =   5535
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "&Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton optCode 
         Caption         =   "&Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3030
         TabIndex        =   6
         Top             =   180
         Width           =   1245
      End
      Begin MSComctlLib.ListView lstjob 
         Height          =   2325
         Left            =   60
         TabIndex        =   9
         Top             =   960
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   4101
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmSMIS_FILE_jobMasterFile.frx":08CA
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "JobCode"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   10
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      Begin VB.TextBox txtdesc 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   600
         Width           =   4440
      End
      Begin VB.TextBox txtcode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   660
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   630
         TabIndex        =   3
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.PictureBox picSaves 
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
      Height          =   885
      Left            =   4200
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   20
      Top             =   4530
      Width           =   1800
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   750
         MouseIcon       =   "frmSMIS_FILE_jobMasterFile.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_jobMasterFile.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   60
         MouseIcon       =   "frmSMIS_FILE_jobMasterFile.frx":0EBC
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_jobMasterFile.frx":100E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picAdds 
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
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   11
      Top             =   4500
      Width           =   6075
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   5010
         MouseIcon       =   "frmSMIS_FILE_jobMasterFile.frx":135E
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_jobMasterFile.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   5010
         MouseIcon       =   "frmSMIS_FILE_jobMasterFile.frx":1816
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_jobMasterFile.frx":1968
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   4320
         MouseIcon       =   "frmSMIS_FILE_jobMasterFile.frx":1CCE
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_jobMasterFile.frx":1E20
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   3630
         MouseIcon       =   "frmSMIS_FILE_jobMasterFile.frx":214B
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_jobMasterFile.frx":229D
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   2940
         MouseIcon       =   "frmSMIS_FILE_jobMasterFile.frx":25F9
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_jobMasterFile.frx":274B
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   2250
         MouseIcon       =   "frmSMIS_FILE_jobMasterFile.frx":2A5E
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_jobMasterFile.frx":2BB0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   1560
         MouseIcon       =   "frmSMIS_FILE_jobMasterFile.frx":2EAA
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_jobMasterFile.frx":2FFC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   870
         MouseIcon       =   "frmSMIS_FILE_jobMasterFile.frx":3354
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_jobMasterFile.frx":34A6
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.Label labid 
         Alignment       =   2  'Center
         Caption         =   "Label4"
         Height          =   435
         Left            =   90
         TabIndex        =   23
         Top             =   -390
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmSMIS_FILE_jobMasterFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UPDATE_MODE                                                       As Boolean
Dim RS                                                                As New ADODB.Recordset

Sub SaveJobs()
    Dim SQL                                                           As String
    Dim theCode                                                       As String
    Dim theDesc                                                       As String

    theCode = Trim(TXTCODE.Text)
    theDesc = Trim(txtdesc.Text)

    'If theCode = "" Then
    '    MsgBox "Please"
    '    Exit Sub
    'End If

    If theDesc = "" Then
        MsgBox "Please input description!!", vbExclamation, "WARNING"
        txtdesc.SetFocus
        Exit Sub
    End If


    If UPDATE_MODE = False Then

        SQL_STATEMENT = "INSERT INTO SMIS_Jobs (jobCode,description) VALUES('" & theCode & "','" & theDesc & "')"
        '*********NEW LOG AUDIT***********
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "A", "VEHICLE MAKE", SQL_STATEMENT, FindTransactionID(N2Str2Null(TXTCODE), "JobCode", "SMIS_JOBS"), "", "Code :" & TXTCODE, "", ""
        '*********NEW LOG AUDIT***********


        LogAudit "A", "JOB MASTER FILE", txtdesc
    Else

        SQL_STATEMENT = "UPDATE SMIS_JOBS set Description='" & theDesc & "' where ID='" & labid & "'"
        '*********NEW LOG AUDIT***********
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "E", "VEHICLE MAKE", SQL_STATEMENT, N2Str2Null(labid), "", "Code :" & TXTCODE, "", ""
        '*********NEW LOG AUDIT***********
        LogAudit "E", "JOB MASTER FILE", txtdesc
    End If
    MsgBox "All information has been save..", vbInformation, "Information"
    If UPDATE_MODE = True Then
        RS.Find "ID=" & labid & ""
    End If
    DisplayJob
    StoreMemVars
    picAdds.Visible = True
    picSaves.Visible = False
End Sub

Sub InitMemVars()
    TXTCODE.Text = ""
    txtdesc.Text = ""
    txtSEARCH.Text = ""
End Sub

Sub DisplayJob()
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset
    Dim arnie                                                         As ListItem


    SQL = "SELECT * FROM SMIS_JOBS"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    lstjob.ListItems.Clear

    Do While Not RS.EOF
        Set arnie = lstjob.ListItems.Add(, , RS!jobcode)
        arnie.SubItems(1) = Null2String(RS!Description)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM SMIS_jobs", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()

    If Not RS.EOF And Not RS.BOF Then
        labid = Null2String(RS!ID)
        TXTCODE.Text = Null2String(RS!jobcode)
        txtdesc.Text = Null2String(RS!Description)
    End If


End Sub

Private Sub cmdAdd_Click()
    picSaves.Visible = True
    picAdds.Visible = False
    UPDATE_MODE = False
    InitMemVars
End Sub

Private Sub cmdCancel_Click()
    picSaves.Visible = False
    picAdds.Visible = True
End Sub

Private Sub cmdDelete_Click()
    Dim ans                                                           As String
    Dim SQL                                                           As String

    If TXTCODE = "" Then
        MsgBox "Nothing to Delete!!", vbInformation, "WARNING"
        Exit Sub
    End If

    ans = MsgBox("Are you sure do you want to delete this record?", vbQuestion + vbYesNo)

    If ans = vbYes Then

        SQL_STATEMENT = "DELETE FROM SMIS_Jobs where jobcode='" & TXTCODE.Text & "'"

        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "X", "VEHICLE MAKE", SQL_STATEMENT, FindTransactionID(N2Str2Null(TXTCODE), "JobCode", "SMIS_JOBS"), "", "Code :" & TXTCODE, "", ""


        LogAudit "X", "JOB MASTER FILE", txtdesc
        gconDMIS.Execute (SQL)
        MsgBox "Record has been deleted..", vbInformation, "INFORMATION"
        DisplayJob
        StoreMemVars
    End If


End Sub

Private Sub cmdEdit_Click()
    UPDATE_MODE = True
    picSaves.Visible = True
    picAdds.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    LogAudit "V", "JOB MASTER FILE", txtdesc
End Sub

Private Sub cmdSave_Click()
    SaveJobs
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picAdds.Visible = True Then;
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show 1
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE MAKE)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labid), "VEHICLE MAKE")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    picSaves.Visible = False
    rsRefresh
    StoreMemVars
    UPDATE_MODE = False
    DisplayJob
    rsRefresh
End Sub

Private Sub lstjob_Click()
    On Error Resume Next
    If lstjob.SelectedItem Is Nothing Then Exit Sub
    TXTCODE.Text = lstjob.ListItems(lstjob.SelectedItem.Index)
    txtdesc.Text = lstjob.SelectedItem.SubItems(1)
End Sub


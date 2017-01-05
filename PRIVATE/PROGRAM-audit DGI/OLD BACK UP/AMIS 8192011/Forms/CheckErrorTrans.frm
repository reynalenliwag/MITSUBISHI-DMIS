VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frm_TOOLS_DuplicateTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Error Files"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   ControlBox      =   0   'False
   FillColor       =   &H0049B049&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "CheckErrorTrans.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5775
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4980
      MouseIcon       =   "CheckErrorTrans.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "CheckErrorTrans.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Exit Window"
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4260
      MouseIcon       =   "CheckErrorTrans.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "CheckErrorTrans.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Process Checking of Error Files"
      Top             =   720
      Width           =   735
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   60
      Width           =   5715
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   1
         Top             =   750
         Width           =   3615
         Begin VB.Label labProcessing 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   2
            Top             =   -30
            Width           =   3525
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   4
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   5
            Top             =   0
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   609
            TX              =   "cmd1"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "CheckErrorTrans.frx":0BAF
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   300
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         Picture         =   "CheckErrorTrans.frx":0BCB
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "CheckErrorTrans.frx":0BE7
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.Label labCPB 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   30
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frm_TOOLS_DuplicateTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                  As ADODB.Recordset
Dim rsJournal_Det                                 As ADODB.Recordset
Dim rsChartAccount                                As ADODB.Recordset

Sub CheckDupTrans()
    Dim ORD_HDkey, AcctCodekey                    As String

    Dim varDupJTYPE, varDupVoucherNo, varDupFileNeym As String
    Dim varDuprecno1, varDuprecno2                As Long
    Dim varDupstatus, DupSql                      As String
    Dim I                                         As Long

    gconDMIS.Execute "delete from AMIS_Duplicat"
    gconDMIS.Execute "delete from AMIS_No_Mstr"
    gconDMIS.Execute "delete from AMIS_NoHeader"
    gconDMIS.Execute "delete from AMIS_NoDetail"

    MsgSpeech "Checking Duplicate Records..."
    Me.Caption = "Checking Duplicate Records..."
    DoEvents
    Screen.MousePointer = 11
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_HD where status <> 'C' order by Jtype,Voucherno asc", gconDMIS
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        I = 0
        MsgSpeech "Checking for Duplicate Issuances..."
        Me.Caption = "Checking for Duplicate Issuances..."
        rsJournal_HD.MoveFirst
        ORD_HDkey = rsJournal_HD!jtype & rsJournal_HD!VOUCHERNO
        Do While Not rsJournal_HD.EOF
            varDuprecno1 = rsJournal_HD!ID
            DoEvents
            If rsJournal_HD.EOF = True Then
                Exit Do
            Else
                If I < rsJournal_HD.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsJournal_HD.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            labProcessing.Caption = "Processing: " & Null2String(rsJournal_HD!jtype) & " #" & Null2String(rsJournal_HD!VOUCHERNO)
            DoEvents
            ORD_HDkey = rsJournal_HD!jtype & rsJournal_HD!VOUCHERNO
            rsJournal_HD.MoveNext
            If rsJournal_HD.EOF = True Then Exit Do
            Do While rsJournal_HD!jtype & rsJournal_HD!VOUCHERNO = ORD_HDkey
                varDupJTYPE = N2Str2Null(rsJournal_HD!jtype)
                varDupVoucherNo = N2Str2Null(rsJournal_HD!VOUCHERNO)
                varDupFileNeym = "'H'"
                varDuprecno2 = rsJournal_HD!ID
                varDupstatus = N2Str2Null(rsJournal_HD!Status)
                If rsJournal_HD!Status = "P" Then
                    DupSql = "insert into AMIS_Duplicat " & _
                             "([TYPE],TRANTYPE,TRANNO,fileneym,recno1,recno2,status)" & _
                             " values ('H', " & varDupJTYPE & ", " & varDupVoucherNo & ", " & varDupFileNeym & _
                             ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                    gconDMIS.Execute DupSql
                Else
                    gconDMIS.Execute ("Delete from AMIS_Journal_HD Where id = " & rsJournal_HD!ID)
                End If
                ORD_HDkey = rsJournal_HD!jtype & rsJournal_HD!VOUCHERNO
                I = I + 1
                progCPB.Value = (I / rsJournal_HD.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsJournal_HD.MoveNext
                If rsJournal_HD.EOF = True Then Exit Do
            Loop
            DoEvents
            If rsJournal_HD.EOF = True Then
                Exit Do
            Else
                If I < rsJournal_HD.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsJournal_HD.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsJournal_HD = Nothing
    DoEvents


    Screen.MousePointer = 11
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "select * from AMIS_ChartAccount order by AcctCode asc", gconDMIS
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        I = 0
        MsgSpeech "Checking for Duplicate Account Codes..."
        Me.Caption = "Checking for Duplicate Account Codes..."
        rsChartAccount.MoveFirst
        AcctCodekey = Null2String(rsChartAccount!ACCTCODE)
        Do While Not rsChartAccount.EOF
            varDuprecno1 = rsChartAccount!ID
            DoEvents
            If rsChartAccount.EOF = True Then
                Exit Do
            Else
                If I < rsChartAccount.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsChartAccount.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            AcctCodekey = Null2String(rsChartAccount!ACCTCODE)
            labProcessing.Caption = "Processing: Account Code -> " & Null2String(rsChartAccount!ACCTCODE)
            DoEvents
            rsChartAccount.MoveNext
            If rsChartAccount.EOF = True Then Exit Do
            Do While rsChartAccount!ACCTCODE = AcctCodekey
                varDupJTYPE = "'A'"
                varDupVoucherNo = N2Str2Null(rsChartAccount!ACCTCODE)
                varDupFileNeym = "'ACCT'"
                varDuprecno2 = rsChartAccount!ID
                DupSql = "insert into AMIS_Duplicat " & _
                         "(TYPE,TRANTYPE,TRANNO,fileneym,recno1,recno2)" & _
                         " values ('A'," & varDupJTYPE & ", " & varDupVoucherNo & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ")"
                gconDMIS.Execute DupSql
                AcctCodekey = Null2String(rsChartAccount!ACCTCODE)
                I = I + 1
                progCPB.Value = (I / rsChartAccount.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsChartAccount.MoveNext
                If rsChartAccount.EOF = True Then Exit Do
            Loop
            DoEvents
            If rsChartAccount.EOF = True Then
                Exit Do
            Else
                If I < rsChartAccount.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsChartAccount.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsChartAccount = Nothing
    DoEvents
    Screen.MousePointer = 11
    Screen.MousePointer = 0
    Set rsJournal_Det = Nothing

Errorcode:
    If Err.Number = 3021 Then Resume Next
End Sub

Sub CheckMatchRec()
    Dim I                                         As Long
    Me.Caption = "CHECK MATCH RECORDS"
    Screen.MousePointer = 11
    DoEvents
    Set rsJournal_Det = New ADODB.Recordset
    rsJournal_Det.Open "select * from AMIS_Journal_Det where status <> 'C' order by id asc", gconDMIS
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        MsgSpeech "Checking Matching Records from Journal Details File..."
        Me.Caption = "Checking Matching Records from Journal Details File..."
        DoEvents
        I = 0
        Do While Not rsJournal_Det.EOF
            labProcessing.Caption = "Processing: " & Null2String(rsJournal_Det!jtype) & " #" & Null2String(rsJournal_Det!VOUCHERNO)
            DoEvents
            Set rsJournal_HD = New ADODB.Recordset
            rsJournal_HD.Open "select VoucherNo from AMIS_Journal_HD where JNO = " & N2Str2Null(rsJournal_Det!JNo) & " AND JTYPE = " & N2Str2Null(rsJournal_Det!jtype) & " AND VoucherNo ='" & Format(rsJournal_Det!VOUCHERNO, "000000") & "'", gconDMIS
            If rsJournal_HD.EOF And rsJournal_HD.BOF Then
                gconDMIS.Execute ("Delete from AMIS_Journal_Det where id = " & rsJournal_Det!ID)
                ' gconDMIS.Execute "insert into AMIS_NoHeader" & _
                  "([TYPE],TRANTYPE,TRANNO,recno,stat_d)" & _
                  " values ('D','" & rsJournal_Det!jType & "', '" & rsJournal_Det!VOUCHERNO & "', " & rsJournal_Det!ID & ", " & N2Str2Null(rsJournal_Det!Status) & ")"
            End If
            I = I + 1
            progCPB.Value = (I / rsJournal_Det.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsJournal_Det.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsJournal_Det = Nothing
    Set rsJournal_HD = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select id,JTYPE,VoucherNo,status from AMIS_Journal_HD where JTYPE <> 'COB' AND JTYPE <> 'ADJ' AND JTYPE <> 'VPJ' AND status <> 'C' order by JTYPE,VoucherNo asc", gconDMIS
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst
        MsgSpeech "Checking Matching records from Journal Header File..."
        Me.Caption = "Checking Matching records from Journal Header File..."
        I = 0
        Do While Not rsJournal_HD.EOF
            labProcessing.Caption = "Processing: " & Null2String(rsJournal_HD!jtype) & " #" & Null2String(rsJournal_HD!VOUCHERNO)
            DoEvents
            Set rsJournal_Det = New ADODB.Recordset
            rsJournal_Det.Open "select JTYPE,VoucherNo from AMIS_Journal_Det where JTYPE = " & N2Str2Null(rsJournal_HD!jtype) & " and VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO), gconDMIS
            If rsJournal_Det.EOF And rsJournal_Det.BOF Then
                gconDMIS.Execute "insert into AMIS_NoDetail " & _
                                 "([TYPE],TRANTYPE,TRANNO,recno,stat_h)" & _
                                 " values ('H', " & N2Str2Null(rsJournal_HD!jtype) & ", " & N2Str2Null(rsJournal_HD!VOUCHERNO) & ", " & rsJournal_HD!ID & ", " & N2Str2Null(rsJournal_HD!Status) & ")"
            End If
            I = I + 1
            progCPB.Value = (I / rsJournal_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsJournal_HD.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsJournal_HD = Nothing
    Set rsJournal_Det = Nothing
End Sub

Private Sub cmdCheck_Click()
'If Function_Access(LOGID, "Acess_Process", "PROCESSING CHECK ERROR TRANSACTIONS") = False Then Exit Sub
    cmdCheck.Enabled = False
    cmdExit.Enabled = False
    CheckDupTrans
    CheckMatchRec

    MsgSpeechBox "Check Complete!"
    Me.Caption = "Check Complete!"
    cmdCheck.Enabled = True
    cmdExit.Enabled = True
    LogAudit "R", "CHECK ERROR FILES"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


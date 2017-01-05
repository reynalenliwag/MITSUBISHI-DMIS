VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmCSMSMatCheckDupTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Check Error Files"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   ControlBox      =   0   'False
   FillColor       =   &H00DEDFDE&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MatCheckDupTrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1590
   ScaleWidth      =   5805
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
      Left            =   5025
      MouseIcon       =   "MatCheckDupTrans.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "MatCheckDupTrans.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   750
      Width           =   705
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&OK"
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
      Left            =   4275
      MouseIcon       =   "MatCheckDupTrans.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "MatCheckDupTrans.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   750
      Width           =   705
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   30
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   90
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
            ToolTipText     =   "Process progress"
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
            MICON           =   "MatCheckDupTrans.frx":0BAF
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "MatCheckDupTrans.frx":0BCB
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "MatCheckDupTrans.frx":0BE7
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
Attribute VB_Name = "frmCSMSMatCheckDupTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMATISS, rsMATREC   As ADODB.Recordset
Attribute rsMATREC.VB_VarUserMemId = 1073938432
Dim rsMatMas, rsShip5, rsTDAYTRAN As ADODB.Recordset
Attribute rsMatMas.VB_VarUserMemId = 1073938434
Attribute rsShip5.VB_VarUserMemId = 1073938434
Attribute rsTDAYTRAN.VB_VarUserMemId = 1073938434

Private Sub cmdCheck_Click()
    cmdCheck.Enabled = False
    cmdExit.Enabled = False
    CheckDupTrans
    CheckMatchRec
    MsgSpeechBox "Check Complete!"
    Me.Caption = "Check Complete!"
    cmdCheck.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
End Sub

Sub CheckDupTrans()
    Dim MatIsskey, MatReckey, MatCdekey As String
    Dim Ship5key, Tdaytrankey As String

    Dim varDupTrantype, varDupTranno, varDupFileNeym As String
    Dim varDuprecno1, varDuprecno2 As Integer
    Dim varDupstatus, DupSql As String
    Dim i                As Integer

    gconDMIS.Execute "delete from PMIS_Duplicat"
    gconDMIS.Execute "delete from PMIS_No_Mstr"
    gconDMIS.Execute "delete from PMIS_NoHeader"
    gconDMIS.Execute "delete from PMIS_NoDetail"

    MsgSpeech "Checking Duplicate Records..."
    Me.Caption = "Checking Duplicate Records..."
    DoEvents
    Screen.MousePointer = 11
    Set rsMATISS = New ADODB.Recordset
    rsMATISS.Open "select id,trantype,tranno,status from CSMS_MatIss where status <> 'C' order by trantype,tranno asc", gconDMIS
    If Not rsMATISS.EOF And Not rsMATISS.BOF Then
        i = 0
        MsgSpeech "Checking for Duplicate Issuances..."
        Me.Caption = "Checking for Duplicate Issuances..."
        rsMATISS.MoveFirst
        MatIsskey = rsMATISS!TRANTYPE & rsMATISS!Tranno
        Do While Not rsMATISS.EOF
            varDuprecno1 = rsMATISS!ID
            DoEvents
            If rsMATISS.EOF = True Then
                Exit Do
            Else
                If i < rsMATISS.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsMATISS.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            labProcessing.Caption = "Processing: " & Null2String(rsMATISS!TRANTYPE) & " #" & Null2String(rsMATISS!Tranno)
            DoEvents
            MatIsskey = rsMATISS!TRANTYPE & rsMATISS!Tranno
            rsMATISS.MoveNext
            If rsMATISS.EOF = True Then Exit Do
            Do While rsMATISS!TRANTYPE & rsMATISS!Tranno = MatIsskey
                varDupTrantype = N2Str2Null(rsMATISS!TRANTYPE)
                varDupTranno = N2Str2Null(rsMATISS!Tranno)
                varDupFileNeym = "'MatIss'"
                varDuprecno2 = rsMATISS!ID
                varDupstatus = N2Str2Null(rsMATISS!Status)
                DupSql = "insert into PMIS_Duplicat " & _
                         "(trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                gconDMIS.Execute DupSql
                MatIsskey = rsMATISS!TRANTYPE & rsMATISS!Tranno
                i = i + 1
                progCPB.Value = (i / rsMATISS.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsMATISS.MoveNext
                If rsMATISS.EOF = True Then Exit Do
            Loop
            DoEvents
            If rsMATISS.EOF = True Then
                Exit Do
            Else
                If i < rsMATISS.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsMATISS.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsMATISS = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsMATREC = New ADODB.Recordset
    rsMATREC.Open "select id,RRno,status from CSMS_MatRec where status <> 'C' order by RRno asc", gconDMIS
    If Not rsMATREC.EOF And Not rsMATREC.BOF Then
        i = 0
        MsgSpeech "Checking for Duplicate Receipts..."
        Me.Caption = "Checking for Duplicate Receipts..."
        rsMATREC.MoveFirst
        MatReckey = rsMATREC!rrno
        Do While Not rsMATREC.EOF
            varDuprecno1 = rsMATREC!ID
            DoEvents
            If rsMATREC.EOF = True Then
                Exit Do
            Else
                If i < rsMATREC.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsMATREC.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            MatReckey = rsMATREC!rrno
            labProcessing.Caption = "Processing: RR #" & Null2String(rsMATREC!rrno)
            DoEvents
            rsMATREC.MoveNext
            If rsMATREC.EOF = True Then Exit Do
            Do While rsMATREC!rrno = MatReckey
                varDupTrantype = "'RR'"
                varDupTranno = N2Str2Null(rsMATREC!rrno)
                varDupFileNeym = "'MatRec'"
                varDuprecno2 = rsMATREC!ID
                varDupstatus = N2Str2Null(rsMATREC!Status)
                DupSql = "insert into PMIS_Duplicat " & _
                         "(trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                gconDMIS.Execute DupSql
                MatReckey = rsMATREC!rrno
                i = i + 1
                progCPB.Value = (i / rsMATREC.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsMATREC.MoveNext
                If rsMATREC.EOF = True Then Exit Do
            Loop
            DoEvents
            If rsMATREC.EOF = True Then
                Exit Do
            Else
                If i < rsMATREC.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsMATREC.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsMATREC = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "select id,MatCde from CSMS_MatMas order by MatCde asc", gconDMIS
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        i = 0
        MsgSpeech "Checking for Duplicate Part Number..."
        Me.Caption = "Checking for Duplicate Part Number..."
        rsMatMas.MoveFirst
        MatCdekey = Null2String(rsMatMas!MATCDE)
        Do While Not rsMatMas.EOF
            varDuprecno1 = rsMatMas!ID
            DoEvents
            If rsMatMas.EOF = True Then
                Exit Do
            Else
                If i < rsMatMas.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsMatMas.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            MatCdekey = Null2String(rsMatMas!MATCDE)
            labProcessing.Caption = "Processing: Part Number " & Null2String(rsMatMas!MATCDE)
            DoEvents
            rsMatMas.MoveNext
            If rsMatMas.EOF = True Then Exit Do
            Do While rsMatMas!MATCDE = MatCdekey
                varDupTrantype = "'MST'"
                varDupTranno = N2Str2Null(rsMatMas!MATCDE)
                varDupFileNeym = "'MST'"
                varDuprecno2 = rsMatMas!ID
                DupSql = "insert into PMIS_Duplicat " & _
                         "(trantype,tranno,fileneym,recno1,recno2)" & _
                       " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ")"
                gconDMIS.Execute DupSql
                MatCdekey = Null2String(rsMatMas!MATCDE)
                i = i + 1
                progCPB.Value = (i / rsMatMas.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsMatMas.MoveNext
                If rsMatMas.EOF = True Then Exit Do
            Loop
            DoEvents
            If rsMatMas.EOF = True Then
                Exit Do
            Else
                If i < rsMatMas.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsMatMas.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsMatMas = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select id,trantype,tranno,itemno,status from PMIS_TdayTran where trantype <> 'ADJ' and status <> 'C' order by trantype,tranno asc", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        i = 0
        MsgSpeech "Checking for Duplicate Entry in Daily transactions File..."
        Me.Caption = "Checking for Duplicate Entry in Tdaytran File..."
        rsTDAYTRAN.MoveFirst
        Tdaytrankey = rsTDAYTRAN!TRANTYPE & rsTDAYTRAN!Tranno & rsTDAYTRAN!itemno
        Do While Not rsTDAYTRAN.EOF
            varDuprecno1 = rsTDAYTRAN!ID
            DoEvents
            If rsTDAYTRAN.EOF = True Then
                Exit Do
            Else
                If i < rsTDAYTRAN.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsTDAYTRAN.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            Tdaytrankey = rsTDAYTRAN!TRANTYPE & rsTDAYTRAN!Tranno & rsTDAYTRAN!itemno
            labProcessing.Caption = "Processing: " & Null2String(rsTDAYTRAN!TRANTYPE) & " #" & Null2String(rsTDAYTRAN!Tranno)
            DoEvents
            rsTDAYTRAN.MoveNext
            If rsTDAYTRAN.EOF = True Then Exit Do
            Do While rsTDAYTRAN!TRANTYPE & rsTDAYTRAN!Tranno & rsTDAYTRAN!itemno = Tdaytrankey
                varDupTrantype = N2Str2Null(rsTDAYTRAN!TRANTYPE)
                varDupTranno = N2Str2Null(rsTDAYTRAN!Tranno)
                varDupFileNeym = "'TDAYTRAN'"
                varDuprecno2 = rsTDAYTRAN!ID
                varDupstatus = N2Str2Null(rsTDAYTRAN!Status)
                DupSql = "insert into PMIS_Duplicat " & _
                         "(trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                gconDMIS.Execute DupSql
                Tdaytrankey = rsTDAYTRAN!TRANTYPE & rsTDAYTRAN!Tranno & rsTDAYTRAN!itemno
                i = i + 1
                progCPB.Value = (i / rsTDAYTRAN.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsTDAYTRAN.MoveNext
            Loop
            DoEvents
            If rsTDAYTRAN.EOF = True Then
                Exit Do
            Else
                If i < rsTDAYTRAN.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsTDAYTRAN.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsTDAYTRAN = Nothing
End Sub

Sub CheckMatchRec()
    Dim MatIsskey, MatReckey, MatCdekey As String
    Dim Ship5key, Tdaytrankey As String

    Dim varDupTrantype, varDupTranno, varDupFileNeym As String
    Dim varDuprecno1, varDuprecno2 As Integer
    Dim varDupstatus     As String

    Dim i                As Integer
    Me.Caption = "CHECK MATCH RECORDS"
    Screen.MousePointer = 11
    DoEvents
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select id,trandate,tranno,trantype,status,MatCde from CSMS_TdayTran where status <> 'C' order by id asc", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        rsTDAYTRAN.MoveFirst
        MsgSpeech "Checking Matching Records from Daily transactions File..."
        Me.Caption = "Checking Matching Records from PMIS_TdayTran File..."
        DoEvents
        i = 0
        Do While Not rsTDAYTRAN.EOF
            labProcessing.Caption = "Processing: " & Null2String(rsTDAYTRAN!TRANTYPE) & " #" & Null2String(rsTDAYTRAN!Tranno)
            DoEvents
            If rsTDAYTRAN!TRANTYPE <> "PO" Or rsTDAYTRAN!TRANTYPE <> "MID" Then
                If rsTDAYTRAN!TRANTYPE = "RR" Then
                    Set rsMATREC = New ADODB.Recordset
                    rsMATREC.Open "select rrno from CSMS_MatRec where rrno ='" & Format(rsTDAYTRAN!Tranno, "000000") & "'", gconDMIS
                    If rsMATREC.EOF And rsMATREC.BOF Then
                        gconDMIS.Execute "insert into PMIS_NoHeader" & _
                                         "(trantype,tranno,recno,stat_d)" & _
                                       " values ('" & rsTDAYTRAN!TRANTYPE & "', '" & rsTDAYTRAN!Tranno & "', " & rsTDAYTRAN!ID & ", " & N2Str2Null(rsTDAYTRAN!Status) & ")"
                    End If
                End If
                If rsTDAYTRAN!TRANTYPE = "CSH" Or rsTDAYTRAN!TRANTYPE = "CHG" Or rsTDAYTRAN!TRANTYPE = "RIV" Or rsTDAYTRAN!TRANTYPE = "DR" Then
                    Set rsMATISS = New ADODB.Recordset
                    rsMATISS.Open "select trantype,tranno from CSMS_MatIss where trantype = " & N2Str2Null(rsTDAYTRAN!TRANTYPE) & " and tranno =" & N2Str2Null(rsTDAYTRAN!Tranno), gconDMIS
                    If rsMATISS.EOF And rsMATISS.BOF Then
                        gconDMIS.Execute "insert into PMIS_NoHeader" & _
                                         "(trantype,tranno,recno,stat_d)" & _
                                       " values (" & N2Str2Null(rsTDAYTRAN!TRANTYPE) & ", " & N2Str2Null(rsTDAYTRAN!Tranno) & ", " & rsTDAYTRAN!ID & ", " & N2Str2Null(rsTDAYTRAN!Status) & ")"
                    End If
                End If
            End If
            If rsTDAYTRAN!TRANTYPE <> "ADB" Then
                Set rsMatMas = New ADODB.Recordset
                rsMatMas.Open "select MatCde from CSMS_MatMas where MatCde = " & N2Str2Null(rsTDAYTRAN!MATCDE) & " order by MatCde asc", gconDMIS
                If rsMatMas.EOF And rsMatMas.BOF Then
                    gconDMIS.Execute "insert into CSMS_No_Mstr" & _
                                     "(trantype,tranno,recno,stat_d)" & _
                                   " values (" & N2Str2Null(rsTDAYTRAN!TRANTYPE) & ", " & N2Str2Null(rsTDAYTRAN!Tranno) & ", " & rsTDAYTRAN!ID & ", " & N2Str2Null(rsTDAYTRAN!Status) & ")"
                End If
            End If
            i = i + 1
            progCPB.Value = (i / rsTDAYTRAN.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTDAYTRAN.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsTDAYTRAN = Nothing
    Set rsMATREC = Nothing
    Set rsMATISS = Nothing
    Set rsMatMas = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsMATISS = New ADODB.Recordset
    rsMATISS.Open "select id,trantype,tranno,status from CSMS_MatIss  where status <> 'C' order by trantype,tranno asc", gconDMIS
    If Not rsMATISS.EOF And Not rsMATISS.BOF Then
        rsMATISS.MoveFirst
        MsgSpeech "Checking Matching records from Issuances Header File..."
        Me.Caption = "Checking Matching records from Order Header File..."
        i = 0
        Do While Not rsMATISS.EOF
            labProcessing.Caption = "Processing: " & Null2String(rsMATISS!TRANTYPE) & " #" & Null2String(rsMATISS!Tranno)
            DoEvents
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select trantype,tranno from PMIS_TdayTran where trantype = " & N2Str2Null(rsMATISS!TRANTYPE) & " and tranno = " & N2Str2Null(rsMATISS!Tranno), gconDMIS
            If rsTDAYTRAN.EOF And rsTDAYTRAN.BOF Then
                gconDMIS.Execute "insert into CSMS_NoDetail " & _
                                 "(trantype,tranno,recno,stat_h)" & _
                               " values (" & N2Str2Null(rsMATISS!TRANTYPE) & ", " & N2Str2Null(rsMATISS!Tranno) & ", " & rsMATISS!ID & ", " & N2Str2Null(rsMATISS!Status) & ")"

            End If
            i = i + 1
            progCPB.Value = (i / rsMATISS.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsMATISS.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsMATISS = Nothing
    Set rsTDAYTRAN = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsMATREC = New ADODB.Recordset
    rsMATREC.Open "select id,rrno,status from CSMS_MatRec  where status <> 'C' order by rrno asc", gconDMIS
    If Not rsMATREC.EOF And Not rsMATREC.BOF Then
        rsMATREC.MoveFirst
        MsgSpeech "Checking Matching records from Receipts Header File..."
        Me.Caption = "Checking Matching records from Receipts Header File..."
        i = 0
        Do While Not rsMATREC.EOF
            labProcessing.Caption = "Processing: RR #" & Null2String(rsMATREC!rrno)
            DoEvents
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select tranno from PMIS_TdayTran where tranno = " & N2Str2Null(rsMATREC!rrno), gconDMIS
            If rsTDAYTRAN.EOF And rsTDAYTRAN.BOF Then
                gconDMIS.Execute "insert into PMIS_NoDetail" & _
                                 "(trantype,tranno,recno,stat_d)" & _
                               " values ('RR', " & N2Str2Null(rsMATREC!rrno) & ", " & rsMATREC!ID & ", " & N2Str2Null(rsMATREC!Status) & ")"
            End If
            i = i + 1
            progCPB.Value = (i / rsMATREC.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsMATREC.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsMATREC = Nothing
    Set rsTDAYTRAN = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISCheckDupTrans = Nothing
    UnloadForm Me
End Sub

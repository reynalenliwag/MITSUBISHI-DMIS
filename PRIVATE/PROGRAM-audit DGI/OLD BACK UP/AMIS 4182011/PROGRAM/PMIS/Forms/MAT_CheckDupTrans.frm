VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmPMISMAT_CheckDupTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Error Files"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   ControlBox      =   0   'False
   FillColor       =   &H0049B049&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_CheckDupTrans.frx":0000
   LinkTopic       =   "Form1"
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
      Left            =   4965
      MouseIcon       =   "MAT_CheckDupTrans.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "MAT_CheckDupTrans.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   750
      Width           =   705
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
      Left            =   4230
      MouseIcon       =   "MAT_CheckDupTrans.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "MAT_CheckDupTrans.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   750
      Width           =   705
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
            MICON           =   "MAT_CheckDupTrans.frx":0BAF
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
         Picture         =   "MAT_CheckDupTrans.frx":0BCB
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "MAT_CheckDupTrans.frx":0BE7
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
Attribute VB_Name = "frmPMISMAT_CheckDupTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOrd_Hd, rsRR_HD, rsPO_HD         As ADODB.Recordset
Attribute rsRR_HD.VB_VarUserMemId = 1073938432
Attribute rsPO_HD.VB_VarUserMemId = 1073938432
Dim rsPartMas, rsShip5, rsTDAYTRAN     As ADODB.Recordset
Attribute rsPartMas.VB_VarUserMemId = 1073938435
Attribute rsShip5.VB_VarUserMemId = 1073938435
Attribute rsTDAYTRAN.VB_VarUserMemId = 1073938435

Private Sub cmdCheck_Click()
    If Function_Access(LOGID, "Acess_Process") = False Then Exit Sub
    cmdCheck.Enabled = False
    cmdExit.Enabled = False
    CheckDupTrans
    CheckMatchRec
    LogAudit "V", "REPORT BIR YEAR END"

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
    Dim ORD_HDkey, RR_HDkey, STOCKNOkey As String
    Dim Ship5key, Tdaytrankey          As String

    Dim varDupTrantype, varDupTranno, varDupFileNeym As String
    Dim varDuprecno1, varDuprecno2     As Long
    Dim varDupstatus, DupSql           As String
    Dim I                              As Integer

    gconDMIS.Execute "delete from PMIS_Duplicat"
    gconDMIS.Execute "delete from PMIS_No_Mstr"
    gconDMIS.Execute "delete from PMIS_NoHeader"
    gconDMIS.Execute "delete from PMIS_NoDetail"

    MsgSpeech "Checking Duplicate Records..."
    Me.Caption = "Checking Duplicate Records..."
    DoEvents
    Screen.MousePointer = 11
    Set rsOrd_Hd = New ADODB.Recordset
    rsOrd_Hd.Open "select id,trantype,tranno,status from PMIS_Ord_Hd where status <> 'C' order by trantype,tranno asc", gconDMIS
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        I = 0
        MsgSpeech "Checking for Duplicate Issuances..."
        Me.Caption = "Checking for Duplicate Issuances..."
        rsOrd_Hd.MoveFirst
        ORD_HDkey = rsOrd_Hd!TRANTYPE & rsOrd_Hd!Tranno
        Do While Not rsOrd_Hd.EOF
            varDuprecno1 = rsOrd_Hd!ID
            DoEvents
            If rsOrd_Hd.EOF = True Then
                Exit Do
            Else
                If I < rsOrd_Hd.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsOrd_Hd.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            labProcessing.Caption = "Processing: " & Null2String(rsOrd_Hd!TRANTYPE) & " #" & Null2String(rsOrd_Hd!Tranno)
            DoEvents
            ORD_HDkey = rsOrd_Hd!TRANTYPE & rsOrd_Hd!Tranno
            rsOrd_Hd.MoveNext
            If rsOrd_Hd.EOF = True Then Exit Do
            Do While rsOrd_Hd!TRANTYPE & rsOrd_Hd!Tranno = ORD_HDkey
                varDupTrantype = N2Str2Null(rsOrd_Hd!TRANTYPE)
                varDupTranno = N2Str2Null(rsOrd_Hd!Tranno)
                varDupFileNeym = "'ORD_HD'"
                varDuprecno2 = rsOrd_Hd!ID
                varDupstatus = N2Str2Null(rsOrd_Hd!Status)
                DupSql = "insert into PMIS_Duplicat " & _
                         "(trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                gconDMIS.Execute DupSql
                ORD_HDkey = rsOrd_Hd!TRANTYPE & rsOrd_Hd!Tranno
                I = I + 1
                progCPB.Value = (I / rsOrd_Hd.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsOrd_Hd.MoveNext
                If rsOrd_Hd.EOF = True Then Exit Do
            Loop
            DoEvents
            If rsOrd_Hd.EOF = True Then
                Exit Do
            Else
                If I < rsOrd_Hd.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsOrd_Hd.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsOrd_Hd = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select id,POno,status from PMIS_PO_Hd where status <> 'C' order by POno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        I = 0
        MsgSpeech "Checking for Duplicate Receipts..."
        Me.Caption = "Checking for Duplicate Receipts..."
        rsRR_HD.MoveFirst
        RR_HDkey = rsRR_HD!PONO
        Do While Not rsRR_HD.EOF
            varDuprecno1 = rsRR_HD!ID
            DoEvents
            If rsRR_HD.EOF = True Then
                Exit Do
            Else
                If I < rsRR_HD.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsRR_HD.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            RR_HDkey = rsRR_HD!PONO
            labProcessing.Caption = "Processing: PO #" & Null2String(rsRR_HD!PONO)
            DoEvents
            rsRR_HD.MoveNext
            If rsRR_HD.EOF = True Then Exit Do
            Do While rsRR_HD!PONO = RR_HDkey
                varDupTrantype = "'PO'"
                varDupTranno = N2Str2Null(rsRR_HD!PONO)
                varDupFileNeym = "'PO_HD'"
                varDuprecno2 = rsRR_HD!ID
                varDupstatus = N2Str2Null(rsRR_HD!Status)
                DupSql = "insert into PMIS_Duplicat " & _
                         "(trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                gconDMIS.Execute DupSql
                RR_HDkey = rsRR_HD!PONO
                I = I + 1
                progCPB.Value = (I / rsRR_HD.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsRR_HD.MoveNext
                If rsRR_HD.EOF = True Then Exit Do
            Loop
            DoEvents
            If rsRR_HD.EOF = True Then
                Exit Do
            Else
                If I < rsRR_HD.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsRR_HD.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    DoEvents
    Screen.MousePointer = 11
    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select id,RRno,status from PMIS_RR_Hd where status <> 'C' order by RRno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        I = 0
        MsgSpeech "Checking for Duplicate Receipts..."
        Me.Caption = "Checking for Duplicate Receipts..."
        rsRR_HD.MoveFirst
        RR_HDkey = rsRR_HD!rrno
        Do While Not rsRR_HD.EOF
            varDuprecno1 = rsRR_HD!ID
            DoEvents
            If rsRR_HD.EOF = True Then
                Exit Do
            Else
                If I < rsRR_HD.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsRR_HD.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            RR_HDkey = rsRR_HD!rrno
            labProcessing.Caption = "Processing: RR #" & Null2String(rsRR_HD!rrno)
            DoEvents
            rsRR_HD.MoveNext
            If rsRR_HD.EOF = True Then Exit Do
            Do While rsRR_HD!rrno = RR_HDkey
                varDupTrantype = "'RR'"
                varDupTranno = N2Str2Null(rsRR_HD!rrno)
                varDupFileNeym = "'RR_HD'"
                varDuprecno2 = rsRR_HD!ID
                varDupstatus = N2Str2Null(rsRR_HD!Status)
                DupSql = "insert into PMIS_Duplicat " & _
                         "(trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                gconDMIS.Execute DupSql
                RR_HDkey = rsRR_HD!rrno
                I = I + 1
                progCPB.Value = (I / rsRR_HD.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsRR_HD.MoveNext
                If rsRR_HD.EOF = True Then Exit Do
            Loop
            DoEvents
            If rsRR_HD.EOF = True Then
                Exit Do
            Else
                If I < rsRR_HD.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsRR_HD.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsRR_HD = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "select id,STOCKNO from PMIS_STOCKMAS where ACTIVE = 'Y' order by STOCKNO asc", gconDMIS
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        I = 0
        MsgSpeech "Checking for Duplicate Part Number..."
        Me.Caption = "Checking for Duplicate Part Number..."
        rsPartMas.MoveFirst
        STOCKNOkey = Null2String(rsPartMas!STOCKNO)
        Do While Not rsPartMas.EOF
            varDuprecno1 = rsPartMas!ID
            DoEvents
            If rsPartMas.EOF = True Then
                Exit Do
            Else
                If I < rsPartMas.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsPartMas.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            STOCKNOkey = Null2String(rsPartMas!STOCKNO)
            labProcessing.Caption = "Processing: Part Number " & Null2String(rsPartMas!STOCKNO)
            DoEvents
            rsPartMas.MoveNext
            If rsPartMas.EOF = True Then Exit Do
            Do While rsPartMas!STOCKNO = STOCKNOkey
                varDupTrantype = "'MST'"
                varDupTranno = N2Str2Null(rsPartMas!STOCKNO)
                varDupFileNeym = "'MST'"
                varDuprecno2 = rsPartMas!ID
                DupSql = "insert into PMIS_Duplicat " & _
                         "(trantype,tranno,fileneym,recno1,recno2)" & _
                       " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ")"
                gconDMIS.Execute DupSql
                STOCKNOkey = Null2String(rsPartMas!STOCKNO)
                I = I + 1
                progCPB.Value = (I / rsPartMas.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsPartMas.MoveNext
                If rsPartMas.EOF = True Then Exit Do
            Loop
            DoEvents
            If rsPartMas.EOF = True Then
                Exit Do
            Else
                If I < rsPartMas.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsPartMas.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsPartMas = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select id,trantype,tranno,itemno,status from PMIS_TdayTran where trantype <> 'ADJ' and status <> 'C' order by trantype,tranno asc", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        I = 0
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
                If I < rsTDAYTRAN.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsTDAYTRAN.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            Tdaytrankey = rsTDAYTRAN!TRANTYPE & rsTDAYTRAN!Tranno & rsTDAYTRAN!itemno
            labProcessing.Caption = "Processing: " & Null2String(rsTDAYTRAN!TRANTYPE) & " #" & Null2String(rsTDAYTRAN!Tranno)
            DoEvents
            rsTDAYTRAN.MoveNext
            If rsTDAYTRAN.EOF = True Then Exit Do
            If Not rsTDAYTRAN.EOF Then
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
                    I = I + 1
                    progCPB.Value = (I / rsTDAYTRAN.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                    rsTDAYTRAN.MoveNext
                Loop
            End If
            DoEvents
            If rsTDAYTRAN.EOF = True Then
                Exit Do
            Else
                If I < rsTDAYTRAN.RecordCount Then
                    I = I + 1
                    progCPB.Value = (I / rsTDAYTRAN.RecordCount) * 100
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
    Dim ORD_HDkey, RR_HDkey, STOCKNOkey As String
    Dim Ship5key, Tdaytrankey          As String

    Dim varDupTrantype, varDupTranno, varDupFileNeym As String
    Dim varDuprecno1, varDuprecno2     As Integer
    Dim varDupstatus                   As String

    Dim I                              As Integer
    Me.Caption = "CHECK MATCH RECORDS"
    Screen.MousePointer = 11
    DoEvents
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select id,trandate,tranno,trantype,status,STOCK_SUP from PMIS_TdayTran where status <> 'C' order by id asc", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        rsTDAYTRAN.MoveFirst
        MsgSpeech "Checking Matching Records from Daily transactions File..."
        Me.Caption = "Checking Matching Records from PMIS_TdayTran File..."
        DoEvents
        I = 0
        Do While Not rsTDAYTRAN.EOF
            labProcessing.Caption = "Processing: " & Null2String(rsTDAYTRAN!TRANTYPE) & " #" & Null2String(rsTDAYTRAN!Tranno)
            DoEvents
            'MODIFIED 3/19/2007
            'If rsTDAYTRAN!TRANTYPE <> "PO" Or rsTDAYTRAN!TRANTYPE <> "MID" Then
            If rsTDAYTRAN!TRANTYPE <> "MID" Then
                If rsTDAYTRAN!TRANTYPE = "PO" Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "select POno from PMIS_PO_Hd where POno ='" & Format(rsTDAYTRAN!Tranno, "000000") & "'", gconDMIS
                    If rsRR_HD.EOF And rsRR_HD.BOF Then
                        gconDMIS.Execute "insert into PMIS_NoHeader" & _
                                         "(trantype,tranno,recno,stat_d)" & _
                                       " values ('" & rsTDAYTRAN!TRANTYPE & "', '" & rsTDAYTRAN!Tranno & "', " & rsTDAYTRAN!ID & ", " & N2Str2Null(rsTDAYTRAN!Status) & ")"
                    End If
                End If
                If rsTDAYTRAN!TRANTYPE = "RR" Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "select rrno from PMIS_RR_Hd where rrno ='" & Format(rsTDAYTRAN!Tranno, "000000") & "'", gconDMIS
                    If rsRR_HD.EOF And rsRR_HD.BOF Then
                        gconDMIS.Execute "insert into PMIS_NoHeader" & _
                                         "(trantype,tranno,recno,stat_d)" & _
                                       " values ('" & rsTDAYTRAN!TRANTYPE & "', '" & rsTDAYTRAN!Tranno & "', " & rsTDAYTRAN!ID & ", " & N2Str2Null(rsTDAYTRAN!Status) & ")"
                    End If
                End If
                If rsTDAYTRAN!TRANTYPE = "CSH" Or rsTDAYTRAN!TRANTYPE = "CHG" Or rsTDAYTRAN!TRANTYPE = "RIV" Or rsTDAYTRAN!TRANTYPE = "DR" Then
                    Set rsOrd_Hd = New ADODB.Recordset
                    rsOrd_Hd.Open "select trantype,tranno from PMIS_Ord_Hd where trantype = " & N2Str2Null(rsTDAYTRAN!TRANTYPE) & " and tranno =" & N2Str2Null(rsTDAYTRAN!Tranno), gconDMIS
                    If rsOrd_Hd.EOF And rsOrd_Hd.BOF Then
                        gconDMIS.Execute "insert into PMIS_NoHeader" & _
                                         "(trantype,tranno,recno,stat_d)" & _
                                       " values (" & N2Str2Null(rsTDAYTRAN!TRANTYPE) & ", " & N2Str2Null(rsTDAYTRAN!Tranno) & ", " & rsTDAYTRAN!ID & ", " & N2Str2Null(rsTDAYTRAN!Status) & ")"
                    End If
                End If
            End If
            If rsTDAYTRAN!TRANTYPE <> "ADB" Then
                Set rsPartMas = New ADODB.Recordset
                rsPartMas.Open "select STOCKNO from PMIS_STOCKMAS where STOCKNO = " & N2Str2Null(rsTDAYTRAN!STOCK_SUP) & " order by STOCKNO asc", gconDMIS
                If rsPartMas.EOF And rsPartMas.BOF Then
                    gconDMIS.Execute "insert into PMIS_No_Mstr" & _
                                     "(trantype,tranno,recno,stat_d)" & _
                                   " values (" & N2Str2Null(rsTDAYTRAN!TRANTYPE) & ", " & N2Str2Null(rsTDAYTRAN!Tranno) & ", " & rsTDAYTRAN!ID & ", " & N2Str2Null(rsTDAYTRAN!Status) & ")"
                End If
            End If
            I = I + 1
            progCPB.Value = (I / rsTDAYTRAN.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTDAYTRAN.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsTDAYTRAN = Nothing
    Set rsRR_HD = Nothing
    Set rsOrd_Hd = Nothing
    Set rsPartMas = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsOrd_Hd = New ADODB.Recordset
    rsOrd_Hd.Open "select id,trantype,tranno,status from PMIS_Ord_Hd where status <> 'C' order by trantype,tranno asc", gconDMIS
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        rsOrd_Hd.MoveFirst
        MsgSpeech "Checking Matching records from Issuances Header File..."
        Me.Caption = "Checking Matching records from Order Header File..."
        I = 0
        Do While Not rsOrd_Hd.EOF
            labProcessing.Caption = "Processing: " & Null2String(rsOrd_Hd!TRANTYPE) & " #" & Null2String(rsOrd_Hd!Tranno)
            DoEvents
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select trantype,tranno from PMIS_TdayTran where trantype = " & N2Str2Null(rsOrd_Hd!TRANTYPE) & " and tranno = " & N2Str2Null(rsOrd_Hd!Tranno), gconDMIS
            If rsTDAYTRAN.EOF And rsTDAYTRAN.BOF Then
                gconDMIS.Execute "insert into PMIS_NoDetail " & _
                                 "(trantype,tranno,recno,stat_h)" & _
                               " values (" & N2Str2Null(rsOrd_Hd!TRANTYPE) & ", " & N2Str2Null(rsOrd_Hd!Tranno) & ", " & rsOrd_Hd!ID & ", " & N2Str2Null(rsOrd_Hd!Status) & ")"
            End If
            I = I + 1
            progCPB.Value = (I / rsOrd_Hd.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsOrd_Hd.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsOrd_Hd = Nothing
    Set rsTDAYTRAN = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select id,rrno,status from PMIS_RR_Hd where status <> 'C' order by rrno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        rsRR_HD.MoveFirst
        MsgSpeech "Checking Matching records from Receipts Header File..."
        Me.Caption = "Checking Matching records from Receipts Header File..."
        I = 0
        Do While Not rsRR_HD.EOF
            labProcessing.Caption = "Processing: RR #" & Null2String(rsRR_HD!rrno)
            DoEvents
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select tranno from PMIS_TdayTran where tranno = " & N2Str2Null(rsRR_HD!rrno), gconDMIS
            If rsTDAYTRAN.EOF And rsTDAYTRAN.BOF Then
                gconDMIS.Execute "insert into PMIS_NoDetail" & _
                                 "(trantype,tranno,recno,stat_d)" & _
                               " values ('RR', " & N2Str2Null(rsRR_HD!rrno) & ", " & rsRR_HD!ID & ", " & N2Str2Null(rsRR_HD!Status) & ")"
            End If
            I = I + 1
            progCPB.Value = (I / rsRR_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsRR_HD.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsRR_HD = Nothing
    Set rsTDAYTRAN = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISCheckDupTrans = Nothing
    UnloadForm Me
End Sub

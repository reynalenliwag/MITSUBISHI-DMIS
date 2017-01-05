VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISProcess_CheckDupTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Error Files"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   ControlBox      =   0   'False
   FillColor       =   &H0049B049&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "CheckDupTrans.frx":0000
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
      MouseIcon       =   "CheckDupTrans.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "CheckDupTrans.frx":045C
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
      MouseIcon       =   "CheckDupTrans.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "CheckDupTrans.frx":0914
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
            MICON           =   "CheckDupTrans.frx":0BAF
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
         Picture         =   "CheckDupTrans.frx":0BCB
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "CheckDupTrans.frx":0BE7
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
Attribute VB_Name = "frmPMISProcess_CheckDupTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOrd_Hd, rsRR_HD, RSPO_HD                                        As ADODB.Recordset
Attribute rsRR_HD.VB_VarUserMemId = 1073938432
Attribute RSPO_HD.VB_VarUserMemId = 1073938432
Dim rsPartMas, rsShip5, rsTdayTran                                    As ADODB.Recordset
Attribute rsPartMas.VB_VarUserMemId = 1073938435
Attribute rsShip5.VB_VarUserMemId = 1073938435
Attribute rsTdayTran.VB_VarUserMemId = 1073938435

Sub CheckDupTrans()
    Dim ORD_HDkey, RR_HDkey, STOCKNOkey                               As String
    Dim varDupTrantype, varDupTranno, varDupFileNeym                  As String
    Dim varDuprecno1, varDuprecno2                                    As Long
    Dim varDupstatus, DupSql                                          As String
    Dim I                                                             As Integer

    gconDMIS.Execute "delete from PMIS_Duplicat"
    gconDMIS.Execute "delete from PMIS_No_Mstr"
    gconDMIS.Execute "delete from PMIS_NoHeader"
    gconDMIS.Execute "delete from PMIS_NoDetail"

    MsgSpeech "Checking Duplicate Records..."
    Me.Caption = "Checking Duplicate Records..."
    DoEvents
    Screen.MousePointer = 11
    Set rsOrd_Hd = New ADODB.Recordset
    rsOrd_Hd.Open "select id,[TYPE],trantype,tranno,status from PMIS_Ord_Hd where (TRANTYPE <> 'ARS' AND TRANTYPE <> 'MRS' AND TRANTYPE <> 'PRS') AND  status <> 'C' order by trantype,tranno asc", gconDMIS
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        I = 0
        MsgSpeech "Checking for Duplicate Issuances..."
        Me.Caption = "Checking for Duplicate Issuances..."
        rsOrd_Hd.MoveFirst
        ORD_HDkey = rsOrd_Hd![Type] & rsOrd_Hd!TranType & rsOrd_Hd!TRANNO
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
            labProcessing.Caption = "Processing: " & Null2String(rsOrd_Hd!TranType) & " #" & Null2String(rsOrd_Hd!TRANNO)
            DoEvents
            ORD_HDkey = rsOrd_Hd!Type & rsOrd_Hd!TranType & rsOrd_Hd!TRANNO
            rsOrd_Hd.MoveNext
            If rsOrd_Hd.EOF = True Then Exit Do
            Do While rsOrd_Hd!Type & rsOrd_Hd!TranType & rsOrd_Hd!TRANNO = ORD_HDkey
                varDupTrantype = N2Str2Null(rsOrd_Hd!TranType)
                varDupTranno = N2Str2Null(rsOrd_Hd!TRANNO)
                varDupFileNeym = "'ORD_HD'"
                varDuprecno2 = rsOrd_Hd!ID
                varDupstatus = N2Str2Null(rsOrd_Hd!STATUS)
                DupSql = "insert into PMIS_Duplicat " & _
                         "([TYPE],trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & N2Str2Null(rsOrd_Hd![Type]) & ", " & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                gconDMIS.Execute DupSql
                ORD_HDkey = rsOrd_Hd!Type & rsOrd_Hd!TranType & rsOrd_Hd!TRANNO
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
    rsRR_HD.Open "select id,[TYPE],POno,status from PMIS_PO_Hd where [TYPE] = 'P' and status <> 'C' order by POno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        I = 0
        MsgSpeech "Checking for Duplicate Receipts..."
        Me.Caption = "Checking for Duplicate Receipts..."
        rsRR_HD.MoveFirst
        RR_HDkey = rsRR_HD!Type & rsRR_HD!PONO
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
            RR_HDkey = rsRR_HD!Type & rsRR_HD!PONO
            labProcessing.Caption = "Processing: PO #" & Null2String(rsRR_HD!PONO)
            DoEvents
            rsRR_HD.MoveNext
            If rsRR_HD.EOF = True Then Exit Do
            Do While rsRR_HD!Type & rsRR_HD!PONO = RR_HDkey
                varDupTrantype = "'PO'"
                varDupTranno = N2Str2Null(rsRR_HD!PONO)
                varDupFileNeym = "'PO_HD'"
                varDuprecno2 = rsRR_HD!ID
                varDupstatus = N2Str2Null(rsRR_HD!STATUS)
                DupSql = "insert into PMIS_Duplicat " & _
                         "([TYPE],trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & N2Str2Null(rsRR_HD!Type) & "," & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
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
    rsRR_HD.Open "select id,[TYPE],RRno,status from PMIS_RR_Hd where [TYPE] = 'P' and status <> 'C' order by RRno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        I = 0
        MsgSpeech "Checking for Duplicate Receipts..."
        Me.Caption = "Checking for Duplicate Receipts..."
        rsRR_HD.MoveFirst
        RR_HDkey = rsRR_HD!Type & rsRR_HD!RRNO
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
            RR_HDkey = rsRR_HD!Type & rsRR_HD!RRNO
            labProcessing.Caption = "Processing: RR #" & Null2String(rsRR_HD!RRNO)
            DoEvents
            rsRR_HD.MoveNext
            If rsRR_HD.EOF = True Then Exit Do
            Do While rsRR_HD!Type & rsRR_HD!RRNO = RR_HDkey
                varDupTrantype = "'RR'"
                varDupTranno = N2Str2Null(rsRR_HD!RRNO)
                varDupFileNeym = "'RR_HD'"
                varDuprecno2 = rsRR_HD!ID
                varDupstatus = N2Str2Null(rsRR_HD!STATUS)
                DupSql = "insert into PMIS_Duplicat " & _
                         "(TYPE,trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & N2Str2Null(rsRR_HD!Type) & "," & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                gconDMIS.Execute DupSql
                RR_HDkey = rsRR_HD!RRNO
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
    rsPartMas.Open "select id,[TYPE],STOCKNO from PMIS_STOCKMAS where ACTIVE = 'Y' order by [TYPE] Desc, STOCKNO asc", gconDMIS
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        I = 0
        MsgSpeech "Checking for Duplicate Stocks..."
        Me.Caption = "Checking for Duplicate Stocks..."
        rsPartMas.MoveFirst
        STOCKNOkey = Null2String(rsPartMas!Type) & Null2String(rsPartMas!STOCKNO)
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
            STOCKNOkey = Null2String(rsPartMas!Type) & Null2String(rsPartMas!STOCKNO)
            labProcessing.Caption = "Processing: Product Code -> " & Null2String(rsPartMas!STOCKNO)
            DoEvents
            rsPartMas.MoveNext
            If rsPartMas.EOF = True Then Exit Do
            Do While rsPartMas!Type & rsPartMas!STOCKNO = STOCKNOkey
                varDupTrantype = "'MST'"
                varDupTranno = N2Str2Null(rsPartMas!STOCKNO)
                varDupFileNeym = "'MST'"
                varDuprecno2 = rsPartMas!ID
                DupSql = "insert into PMIS_Duplicat " & _
                         "(TYPE,trantype,tranno,fileneym,recno1,recno2)" & _
                       " values (" & N2Str2Null(Null2String(rsPartMas!Type)) & "," & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
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
    Screen.MousePointer = 0
    Set rsTdayTran = Nothing

ERRORCODE:
    If err.Number = 3021 Then Resume Next
End Sub

Sub CheckMatchRec()
    Dim I                                                             As Integer
    Me.Caption = "CHECK MATCH RECORDS"
    Screen.MousePointer = 11
    DoEvents
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,[TYPE],trandate,tranno,trantype,status,STOCK_SUP from PMIS_TdayTran where (TRANTYPE <> 'ARS' AND TRANTYPE <> 'MRS' AND TRANTYPE <> 'PRS') AND TRANTYPE <> 'ADJ' AND TRANTYPE <> 'BEG' AND status <> 'C' order by id asc", gconDMIS
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        rsTdayTran.MoveFirst
        MsgSpeech "Checking Matching Records from Daily transactions File..."
        Me.Caption = "Checking Matching Records from PMIS_TdayTran File..."
        DoEvents
        I = 0
        Do While Not rsTdayTran.EOF
            labProcessing.Caption = "Processing: " & Null2String(rsTdayTran!Type) & Null2String(rsTdayTran!TranType) & " #" & Null2String(rsTdayTran!TRANNO)
            DoEvents
            If rsTdayTran!TranType <> "MID" Then
                If rsTdayTran!TranType = "PO" Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "select POno from PMIS_PO_Hd where TYPE = " & N2Str2Null(rsTdayTran!Type) & " AND POno ='" & Format(rsTdayTran!TRANNO, "000000") & "'", gconDMIS
                    If rsRR_HD.EOF And rsRR_HD.BOF Then
                        gconDMIS.Execute "insert into PMIS_NoHeader" & _
                                         "([TYPE],trantype,tranno,recno,stat_d)" & _
                                       " values ('" & rsTdayTran!Type & "','" & rsTdayTran!TranType & "', '" & rsTdayTran!TRANNO & "', " & rsTdayTran!ID & ", " & N2Str2Null(rsTdayTran!STATUS) & ")"
                    End If
                End If
                If rsTdayTran!TranType = "RR" Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "select rrno from PMIS_RR_Hd where TYPE = " & N2Str2Null(rsTdayTran!Type) & " AND rrno ='" & Format(rsTdayTran!TRANNO, "000000") & "'", gconDMIS
                    If rsRR_HD.EOF And rsRR_HD.BOF Then
                        gconDMIS.Execute "insert into PMIS_NoHeader" & _
                                         "([TYPE],trantype,tranno,recno,stat_d)" & _
                                       " values ('" & rsTdayTran![Type] & "','" & rsTdayTran!TranType & "', '" & rsTdayTran!TRANNO & "', " & rsTdayTran!ID & ", " & N2Str2Null(rsTdayTran!STATUS) & ")"
                    End If
                End If
                If rsTdayTran!TranType = "CSH" Or rsTdayTran!TranType = "CHG" Or rsTdayTran!TranType = "RIV" Or rsTdayTran!TranType = "DR" Then
                    Set rsOrd_Hd = New ADODB.Recordset
                    rsOrd_Hd.Open "select trantype,tranno from PMIS_Ord_Hd where TYPE = " & N2Str2Null(rsTdayTran!Type) & " AND trantype = " & N2Str2Null(rsTdayTran!TranType) & " And tranno = " & N2Str2Null(rsTdayTran!TRANNO), gconDMIS
                    If rsOrd_Hd.EOF And rsOrd_Hd.BOF Then
                        gconDMIS.Execute "insert into PMIS_NoHeader" & _
                                         "([TYPE],trantype,tranno,recno,stat_d)" & _
                                       " values (" & N2Str2Null(rsTdayTran![Type]) & "," & N2Str2Null(rsTdayTran!TranType) & ", " & N2Str2Null(rsTdayTran!TRANNO) & ", " & rsTdayTran!ID & ", " & N2Str2Null(rsTdayTran!STATUS) & ")"
                    End If
                End If
            End If
            If rsTdayTran!TranType <> "ADB" Then
                Set rsPartMas = New ADODB.Recordset
                rsPartMas.Open "select STOCKNO from PMIS_STOCKMAS where [TYPE] = " & N2Str2Null(rsTdayTran!Type) & " AND STOCKNO = " & N2Str2Null(rsTdayTran!STOCK_SUP) & " order by STOCKNO asc", gconDMIS
                If rsPartMas.EOF And rsPartMas.BOF Then
                    gconDMIS.Execute "insert into PMIS_No_Mstr" & _
                                     "(TYPE,trantype,tranno,recno,stat_d)" & _
                                   " values (" & N2Str2Null(rsTdayTran![Type]) & "," & N2Str2Null(rsTdayTran!TranType) & ", " & N2Str2Null(rsTdayTran!TRANNO) & ", " & rsTdayTran!ID & ", " & N2Str2Null(rsTdayTran!STATUS) & ")"
                End If
            End If
            I = I + 1
            progCPB.Value = (I / rsTdayTran.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTdayTran.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsTdayTran = Nothing
    Set rsRR_HD = Nothing
    Set rsOrd_Hd = Nothing
    Set rsPartMas = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsOrd_Hd = New ADODB.Recordset
    rsOrd_Hd.Open "select id,[TYPE],trantype,tranno,status from PMIS_Ord_Hd where (TRANTYPE <> 'ARS' AND TRANTYPE <> 'MRS' AND TRANTYPE <> 'PRS') AND status <> 'C' order by trantype,tranno asc", gconDMIS
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        rsOrd_Hd.MoveFirst
        MsgSpeech "Checking Matching records from Issuances Header File..."
        Me.Caption = "Checking Matching records from Order Header File..."
        I = 0
        Do While Not rsOrd_Hd.EOF
            labProcessing.Caption = "Processing: " & Null2String(rsOrd_Hd!Type) & Null2String(rsOrd_Hd!TranType) & " #" & Null2String(rsOrd_Hd!TRANNO)
            DoEvents
            Set rsTdayTran = New ADODB.Recordset
            rsTdayTran.Open "select trantype,tranno from PMIS_TdayTran where TYPE = " & N2Str2Null(rsOrd_Hd!Type) & " AND trantype = " & N2Str2Null(rsOrd_Hd!TranType) & " and tranno = " & N2Str2Null(rsOrd_Hd!TRANNO), gconDMIS
            If rsTdayTran.EOF And rsTdayTran.BOF Then
                gconDMIS.Execute "insert into PMIS_NoDetail " & _
                                 "([TYPE],trantype,tranno,recno,stat_h)" & _
                               " values (" & N2Str2Null(rsOrd_Hd!Type) & ", " & N2Str2Null(rsOrd_Hd!TranType) & ", " & N2Str2Null(rsOrd_Hd!TRANNO) & ", " & rsOrd_Hd!ID & ", " & N2Str2Null(rsOrd_Hd!STATUS) & ")"
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
    Set rsTdayTran = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select id,[TYPE],rrno,status from PMIS_RR_Hd where status <> 'C' order by rrno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        rsRR_HD.MoveFirst
        MsgSpeech "Checking Matching records from Receipts Header File..."
        Me.Caption = "Checking Matching records from Receipts Header File..."
        I = 0
        Do While Not rsRR_HD.EOF
            labProcessing.Caption = "Processing: RR #" & Null2String(rsRR_HD!RRNO)
            DoEvents
            Set rsTdayTran = New ADODB.Recordset
            rsTdayTran.Open "select tranno from PMIS_TdayTran where [TYPE] = " & N2Str2Null(rsRR_HD!Type) & " AND tranno = " & N2Str2Null(rsRR_HD!RRNO), gconDMIS
            If rsTdayTran.EOF And rsTdayTran.BOF Then
                gconDMIS.Execute "insert into PMIS_NoDetail" & _
                                 "([TYPE],trantype,tranno,recno,stat_d)" & _
                               " values (" & N2Str2Null(rsRR_HD!Type) & ",'RR', " & N2Str2Null(rsRR_HD!RRNO) & ", " & rsRR_HD!ID & ", " & N2Str2Null(rsRR_HD!STATUS) & ")"
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
    Set rsTdayTran = Nothing
End Sub

Private Sub cmdCheck_Click()
    If Function_Access(LOGID, "Acess_Process", "PROCESSING CHECK ERROR TRANSACTIONS") = False Then Exit Sub
    cmdCheck.Enabled = False
    cmdExit.Enabled = False
    CheckDupTrans
    CheckMatchRec
    LogAudit "V", "REPORT BIR YEAR END"

    MsgSpeechBox "Check Complete!"
    Me.Caption = "Check Complete!"
    cmdCheck.Enabled = True
    cmdExit.Enabled = True
    FLAG = 1 ' make the commandbars control  set to true
    frmPMISInquiry_ErrorQuery.Show
    NEW_LogAudit "R", "PROCESSING CHECK ERROR TRANSACTIONS", "", "", "", LOGDATE, "", ""
    
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


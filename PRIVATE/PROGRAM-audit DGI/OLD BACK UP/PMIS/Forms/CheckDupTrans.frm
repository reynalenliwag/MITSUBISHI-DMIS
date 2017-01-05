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
Dim RSORD_HD, rsRR_HD, RSPO_HD                         As ADODB.Recordset
Attribute rsRR_HD.VB_VarUserMemId = 1073938432
Attribute RSPO_HD.VB_VarUserMemId = 1073938432
Dim RSPARTMAS, rsShip5, RSTDAYTRAN                     As ADODB.Recordset
Attribute RSPARTMAS.VB_VarUserMemId = 1073938435
Attribute rsShip5.VB_VarUserMemId = 1073938435
Attribute RSTDAYTRAN.VB_VarUserMemId = 1073938435

Sub CheckDupTrans()
    Dim ORD_HDkey, RR_HDkey, STOCKNOkey                As String
    Dim varDupTrantype, varDupTranno, varDupFileNeym   As String
    Dim varDuprecno1, varDuprecno2                     As Long
    Dim varDupstatus, DupSql                           As String
    Dim i                                              As Long

    gconDMIS.Execute "delete from PMIS_Duplicat"
    gconDMIS.Execute "delete from PMIS_No_Mstr"
    gconDMIS.Execute "delete from PMIS_NoHeader"
    gconDMIS.Execute "delete from PMIS_NoDetail"

    MsgSpeech "Checking Duplicate Records..."
    Me.Caption = "Checking Duplicate Records..."
    DoEvents
    Screen.MousePointer = 11
    Set RSORD_HD = New ADODB.Recordset
    RSORD_HD.Open "select id,[TYPE],trantype,tranno,status from PMIS_Ord_Hd where (TRANTYPE <> 'ARS' AND TRANTYPE <> 'MRS' AND TRANTYPE <> 'PRS') AND  status <> 'C' order by trantype,tranno asc", gconDMIS
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        i = 0
        MsgSpeech "Checking for Duplicate Issuances..."
        Me.Caption = "Checking for Duplicate Issuances..."
        RSORD_HD.MoveFirst
        ORD_HDkey = RSORD_HD![Type] & RSORD_HD!TranType & RSORD_HD!TRANNO
        Do While Not RSORD_HD.EOF
            varDuprecno1 = RSORD_HD!ID
            DoEvents
            If RSORD_HD.EOF = True Then
                Exit Do
            Else
                If i < RSORD_HD.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / RSORD_HD.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            labProcessing.Caption = "Processing: " & Null2String(RSORD_HD!TranType) & " #" & Null2String(RSORD_HD!TRANNO)
            DoEvents
            ORD_HDkey = RSORD_HD!Type & RSORD_HD!TranType & RSORD_HD!TRANNO
            RSORD_HD.MoveNext
            If RSORD_HD.EOF = True Then Exit Do
            Do While RSORD_HD!Type & RSORD_HD!TranType & RSORD_HD!TRANNO = ORD_HDkey
                varDupTrantype = N2Str2Null(RSORD_HD!TranType)
                varDupTranno = N2Str2Null(RSORD_HD!TRANNO)
                varDupFileNeym = "'ORD_HD'"
                varDuprecno2 = RSORD_HD!ID
                varDupstatus = N2Str2Null(RSORD_HD!Status)
                DupSql = "insert into PMIS_Duplicat " & _
                         "([TYPE],trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & N2Str2Null(RSORD_HD![Type]) & ", " & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                gconDMIS.Execute DupSql
                ORD_HDkey = RSORD_HD!Type & RSORD_HD!TranType & RSORD_HD!TRANNO
                i = i + 1
                progCPB.Value = (i / RSORD_HD.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                RSORD_HD.MoveNext
                If RSORD_HD.EOF = True Then Exit Do
            Loop
            DoEvents
            If RSORD_HD.EOF = True Then
                Exit Do
            Else
                If i < RSORD_HD.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / RSORD_HD.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set RSORD_HD = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select id,[TYPE],POno,status from PMIS_PO_Hd where [TYPE] = 'P' and status <> 'C' order by POno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        i = 0
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
                If i < rsRR_HD.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsRR_HD.RecordCount) * 100
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
                varDupstatus = N2Str2Null(rsRR_HD!Status)
                DupSql = "insert into PMIS_Duplicat " & _
                         "([TYPE],trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & N2Str2Null(rsRR_HD!Type) & "," & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                gconDMIS.Execute DupSql
                RR_HDkey = rsRR_HD!PONO
                i = i + 1
                progCPB.Value = (i / rsRR_HD.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsRR_HD.MoveNext
                If rsRR_HD.EOF = True Then Exit Do
            Loop
            DoEvents
            If rsRR_HD.EOF = True Then
                Exit Do
            Else
                If i < rsRR_HD.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsRR_HD.RecordCount) * 100
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
        i = 0
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
                If i < rsRR_HD.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsRR_HD.RecordCount) * 100
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
                varDupstatus = N2Str2Null(rsRR_HD!Status)
                DupSql = "insert into PMIS_Duplicat " & _
                         "(TYPE,trantype,tranno,fileneym,recno1,recno2,status)" & _
                       " values (" & N2Str2Null(rsRR_HD!Type) & "," & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
                gconDMIS.Execute DupSql
                RR_HDkey = rsRR_HD!RRNO
                i = i + 1
                progCPB.Value = (i / rsRR_HD.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsRR_HD.MoveNext
                If rsRR_HD.EOF = True Then Exit Do
            Loop
            DoEvents
            If rsRR_HD.EOF = True Then
                Exit Do
            Else
                If i < rsRR_HD.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / rsRR_HD.RecordCount) * 100
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
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "select id,[TYPE],STOCKNO from PMIS_STOCKMAS where ACTIVE = 'Y' order by [TYPE] Desc, STOCKNO asc", gconDMIS
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        i = 0
        MsgSpeech "Checking for Duplicate Stocks..."
        Me.Caption = "Checking for Duplicate Stocks..."
        RSPARTMAS.MoveFirst
        STOCKNOkey = Null2String(RSPARTMAS!Type) & Null2String(RSPARTMAS!STOCKNO)
        Do While Not RSPARTMAS.EOF
            varDuprecno1 = RSPARTMAS!ID
            DoEvents
            If RSPARTMAS.EOF = True Then
                Exit Do
            Else
                If i < RSPARTMAS.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / RSPARTMAS.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
            STOCKNOkey = Null2String(RSPARTMAS!Type) & Null2String(RSPARTMAS!STOCKNO)
            labProcessing.Caption = "Processing: Product Code -> " & Null2String(RSPARTMAS!STOCKNO)
            DoEvents
            RSPARTMAS.MoveNext
            If RSPARTMAS.EOF = True Then Exit Do
            Do While RSPARTMAS!Type & RSPARTMAS!STOCKNO = STOCKNOkey
                varDupTrantype = "'MST'"
                varDupTranno = N2Str2Null(RSPARTMAS!STOCKNO)
                varDupFileNeym = "'MST'"
                varDuprecno2 = RSPARTMAS!ID
                DupSql = "insert into PMIS_Duplicat " & _
                         "(TYPE,trantype,tranno,fileneym,recno1,recno2)" & _
                       " values (" & N2Str2Null(Null2String(RSPARTMAS!Type)) & "," & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                         ", " & varDuprecno1 & ", " & varDuprecno2 & ")"
                gconDMIS.Execute DupSql
                STOCKNOkey = Null2String(RSPARTMAS!STOCKNO)
                i = i + 1
                progCPB.Value = (i / RSPARTMAS.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                RSPARTMAS.MoveNext
                If RSPARTMAS.EOF = True Then Exit Do
            Loop
            DoEvents
            If RSPARTMAS.EOF = True Then
                Exit Do
            Else
                If i < RSPARTMAS.RecordCount Then
                    i = i + 1
                    progCPB.Value = (i / RSPARTMAS.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed"
                    DoEvents
                End If
            End If
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set RSPARTMAS = Nothing
    DoEvents
    Screen.MousePointer = 11
    Screen.MousePointer = 0
    Set RSTDAYTRAN = Nothing

Errorcode:
    If err.Number = 3021 Then Resume Next
End Sub

Sub CheckMatchRec()
    Dim i                                              As Integer
    Me.Caption = "CHECK MATCH RECORDS"
    Screen.MousePointer = 11
    DoEvents
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select id,[TYPE],trandate,tranno,trantype,status,STOCK_SUP from PMIS_TdayTran where (TRANTYPE <> 'ARS' AND TRANTYPE <> 'MRS' AND TRANTYPE <> 'PRS') AND TRANTYPE <> 'ADJ' AND TRANTYPE <> 'BEG' AND status <> 'C' order by id asc", gconDMIS
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst
        MsgSpeech "Checking Matching Records from Daily transactions File..."
        Me.Caption = "Checking Matching Records from PMIS_TdayTran File..."
        DoEvents
        i = 0
        Do While Not RSTDAYTRAN.EOF
            labProcessing.Caption = "Processing: " & Null2String(RSTDAYTRAN!Type) & Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(RSTDAYTRAN!TRANNO)
            DoEvents
            If RSTDAYTRAN!TranType <> "MID" Then
                If RSTDAYTRAN!TranType = "PO" Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "select POno from PMIS_PO_Hd where TYPE = " & N2Str2Null(RSTDAYTRAN!Type) & " AND POno ='" & Format(RSTDAYTRAN!TRANNO, "000000") & "'", gconDMIS
                    If rsRR_HD.EOF And rsRR_HD.BOF Then
                        gconDMIS.Execute "insert into PMIS_NoHeader" & _
                                         "([TYPE],trantype,tranno,recno,stat_d)" & _
                                       " values ('" & RSTDAYTRAN!Type & "','" & RSTDAYTRAN!TranType & "', '" & RSTDAYTRAN!TRANNO & "', " & RSTDAYTRAN!ID & ", " & N2Str2Null(RSTDAYTRAN!Status) & ")"
                    End If
                End If
                If RSTDAYTRAN!TranType = "RR" Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "select rrno from PMIS_RR_Hd where TYPE = " & N2Str2Null(RSTDAYTRAN!Type) & " AND rrno ='" & Format(RSTDAYTRAN!TRANNO, "000000") & "'", gconDMIS
                    If rsRR_HD.EOF And rsRR_HD.BOF Then
                        gconDMIS.Execute "insert into PMIS_NoHeader" & _
                                         "([TYPE],trantype,tranno,recno,stat_d)" & _
                                       " values ('" & RSTDAYTRAN![Type] & "','" & RSTDAYTRAN!TranType & "', '" & RSTDAYTRAN!TRANNO & "', " & RSTDAYTRAN!ID & ", " & N2Str2Null(RSTDAYTRAN!Status) & ")"
                    End If
                End If
                If RSTDAYTRAN!TranType = "CSH" Or RSTDAYTRAN!TranType = "CHG" Or RSTDAYTRAN!TranType = "RIV" Or RSTDAYTRAN!TranType = "DR" Then
                    Set RSORD_HD = New ADODB.Recordset
                    RSORD_HD.Open "select trantype,tranno from PMIS_Ord_Hd where TYPE = " & N2Str2Null(RSTDAYTRAN!Type) & " AND trantype = " & N2Str2Null(RSTDAYTRAN!TranType) & " And tranno = " & N2Str2Null(RSTDAYTRAN!TRANNO), gconDMIS
                    If RSORD_HD.EOF And RSORD_HD.BOF Then
                        gconDMIS.Execute "insert into PMIS_NoHeader" & _
                                         "([TYPE],trantype,tranno,recno,stat_d)" & _
                                       " values (" & N2Str2Null(RSTDAYTRAN![Type]) & "," & N2Str2Null(RSTDAYTRAN!TranType) & ", " & N2Str2Null(RSTDAYTRAN!TRANNO) & ", " & RSTDAYTRAN!ID & ", " & N2Str2Null(RSTDAYTRAN!Status) & ")"
                    End If
                End If
            End If
            If RSTDAYTRAN!TranType <> "ADB" Then
                Set RSPARTMAS = New ADODB.Recordset
                RSPARTMAS.Open "select STOCKNO from PMIS_STOCKMAS where [TYPE] = " & N2Str2Null(RSTDAYTRAN!Type) & " AND STOCKNO = " & N2Str2Null(RSTDAYTRAN!STOCK_SUP) & " order by STOCKNO asc", gconDMIS
                If RSPARTMAS.EOF And RSPARTMAS.BOF Then
                    gconDMIS.Execute "insert into PMIS_No_Mstr" & _
                                     "(TYPE,trantype,tranno,recno,stat_d)" & _
                                   " values (" & N2Str2Null(RSTDAYTRAN![Type]) & "," & N2Str2Null(RSTDAYTRAN!TranType) & ", " & N2Str2Null(RSTDAYTRAN!TRANNO) & ", " & RSTDAYTRAN!ID & ", " & N2Str2Null(RSTDAYTRAN!Status) & ")"
                End If
            End If
            i = i + 1
            progCPB.Value = (i / RSTDAYTRAN.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            RSTDAYTRAN.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set RSTDAYTRAN = Nothing
    Set rsRR_HD = Nothing
    Set RSORD_HD = Nothing
    Set RSPARTMAS = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set RSORD_HD = New ADODB.Recordset
    RSORD_HD.Open "select id,[TYPE],trantype,tranno,status from PMIS_Ord_Hd where (TRANTYPE <> 'ARS' AND TRANTYPE <> 'MRS' AND TRANTYPE <> 'PRS') AND status <> 'C' order by trantype,tranno asc", gconDMIS
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        RSORD_HD.MoveFirst
        MsgSpeech "Checking Matching records from Issuances Header File..."
        Me.Caption = "Checking Matching records from Order Header File..."
        i = 0
        Do While Not RSORD_HD.EOF
            labProcessing.Caption = "Processing: " & Null2String(RSORD_HD!Type) & Null2String(RSORD_HD!TranType) & " #" & Null2String(RSORD_HD!TRANNO)
            DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select trantype,tranno from PMIS_TdayTran where TYPE = " & N2Str2Null(RSORD_HD!Type) & " AND trantype = " & N2Str2Null(RSORD_HD!TranType) & " and tranno = " & N2Str2Null(RSORD_HD!TRANNO), gconDMIS
            If RSTDAYTRAN.EOF And RSTDAYTRAN.BOF Then
                gconDMIS.Execute "insert into PMIS_NoDetail " & _
                                 "([TYPE],trantype,tranno,recno,stat_h)" & _
                               " values (" & N2Str2Null(RSORD_HD!Type) & ", " & N2Str2Null(RSORD_HD!TranType) & ", " & N2Str2Null(RSORD_HD!TRANNO) & ", " & RSORD_HD!ID & ", " & N2Str2Null(RSORD_HD!Status) & ")"
            End If
            i = i + 1
            progCPB.Value = (i / RSORD_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            RSORD_HD.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set RSORD_HD = Nothing
    Set RSTDAYTRAN = Nothing
    DoEvents
    Screen.MousePointer = 11
    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select id,[TYPE],rrno,status from PMIS_RR_Hd where status <> 'C' order by rrno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        rsRR_HD.MoveFirst
        MsgSpeech "Checking Matching records from Receipts Header File..."
        Me.Caption = "Checking Matching records from Receipts Header File..."
        i = 0
        Do While Not rsRR_HD.EOF
            labProcessing.Caption = "Processing: RR #" & Null2String(rsRR_HD!RRNO)
            DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select tranno from PMIS_TdayTran where [TYPE] = " & N2Str2Null(rsRR_HD!Type) & " AND tranno = " & N2Str2Null(rsRR_HD!RRNO), gconDMIS
            If RSTDAYTRAN.EOF And RSTDAYTRAN.BOF Then
                gconDMIS.Execute "insert into PMIS_NoDetail" & _
                                 "([TYPE],trantype,tranno,recno,stat_d)" & _
                               " values (" & N2Str2Null(rsRR_HD!Type) & ",'RR', " & N2Str2Null(rsRR_HD!RRNO) & ", " & rsRR_HD!ID & ", " & N2Str2Null(rsRR_HD!Status) & ")"
            End If
            i = i + 1
            progCPB.Value = (i / rsRR_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsRR_HD.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    End If
    Screen.MousePointer = 0
    Set rsRR_HD = Nothing
    Set RSTDAYTRAN = Nothing
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
    FLAG = 1                                          ' make the commandbars control  set to true
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


VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmPMIOSRECONCheckDupTrans 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Error Files"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   ControlBox      =   0   'False
   FillColor       =   &H0049B049&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "RECON_CheckDupTrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "RECON_CheckDupTrans.frx":030A
   ScaleHeight     =   1590
   ScaleWidth      =   5775
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   765
      Left            =   4770
      MouseIcon       =   "RECON_CheckDupTrans.frx":3046
      MousePointer    =   99  'Custom
      Picture         =   "RECON_CheckDupTrans.frx":3350
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Close Window"
      Top             =   750
      Width           =   945
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Check"
      Height          =   765
      Left            =   3840
      MouseIcon       =   "RECON_CheckDupTrans.frx":365A
      MousePointer    =   99  'Custom
      Picture         =   "RECON_CheckDupTrans.frx":3964
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   750
      Width           =   945
   End
   Begin VB.PictureBox picCPB 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   0
      Picture         =   "RECON_CheckDupTrans.frx":422E
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
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         Picture         =   "RECON_CheckDupTrans.frx":6F6A
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
            MICON           =   "RECON_CheckDupTrans.frx":9CA6
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
         Picture         =   "RECON_CheckDupTrans.frx":9CC2
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "RECON_CheckDupTrans.frx":9CDE
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
Attribute VB_Name = "frmPMIOSRECONCheckDupTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRECON_ORD_HIST, rsRECON_REC_HIST, rsRECON_PO_HIST As ADODB.Recordset
Dim rsRECON_PARTMAS, rsShip5, rsRECON_DAYTRAN As ADODB.Recordset

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
Dim RECON_ORD_HISTkey, RECON_REC_HISTkey, PartNokey As String
Dim Ship5key, RECON_DAYTRANkey As String

Dim varDupTrantype, varDupTranno, varDupFileNeym As String
Dim varDuprecno1, varDuprecno2 As Long
Dim varDupstatus, DupSql As String
Dim i As Long

gconPMIOS.Execute "delete from duplicat"
gconPMIOS.Execute "delete from no_mstr"
gconPMIOS.Execute "delete from noheader"
gconPMIOS.Execute "delete from nodetail"

MsgSpeech "Checking Duplicate Records..."
Me.Caption = "Checking Duplicate Records..."
DoEvents
Screen.MousePointer = 11
Set rsRECON_ORD_HIST = New ADODB.Recordset
    rsRECON_ORD_HIST.Open "select id,trantype,tranno,status from RECON_ORD_HIST where status <> 'C' order by trantype,tranno asc", gconPMIOS
If Not rsRECON_ORD_HIST.EOF And Not rsRECON_ORD_HIST.BOF Then
   i = 0
   MsgSpeech "Checking for Duplicate Issuances..."
   Me.Caption = "Checking for Duplicate Issuances..."
   rsRECON_ORD_HIST.MoveFirst
   RECON_ORD_HISTkey = rsRECON_ORD_HIST!trantype & rsRECON_ORD_HIST!tranno
   Do While Not rsRECON_ORD_HIST.EOF
      varDuprecno1 = rsRECON_ORD_HIST!ID
      DoEvents
      If rsRECON_ORD_HIST.EOF = True Then
         Exit Do
      Else
         If i < rsRECON_ORD_HIST.RecordCount Then
            i = i + 1
            progCPB.Value = (i / rsRECON_ORD_HIST.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
         End If
      End If
      labProcessing.Caption = "Processing: " & Null2String(rsRECON_ORD_HIST!trantype) & " #" & Null2String(rsRECON_ORD_HIST!tranno)
      DoEvents
      RECON_ORD_HISTkey = rsRECON_ORD_HIST!trantype & rsRECON_ORD_HIST!tranno
      rsRECON_ORD_HIST.MoveNext
      If rsRECON_ORD_HIST.EOF = True Then Exit Do
      Do While rsRECON_ORD_HIST!trantype & rsRECON_ORD_HIST!tranno = RECON_ORD_HISTkey
         varDupTrantype = N2Str2Null(rsRECON_ORD_HIST!trantype)
         varDupTranno = N2Str2Null(rsRECON_ORD_HIST!tranno)
         varDupFileNeym = "'RECON_ORD_HIST'"
         varDuprecno2 = rsRECON_ORD_HIST!ID
         varDupstatus = N2Str2Null(rsRECON_ORD_HIST!Status)
         DupSql = "insert into duplicat " & _
                  "(trantype,tranno,fileneym,recno1,recno2,status)" & _
                  " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                  ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
         gconPMIOS.Execute DupSql
         RECON_ORD_HISTkey = rsRECON_ORD_HIST!trantype & rsRECON_ORD_HIST!tranno
         i = i + 1
         progCPB.Value = (i / rsRECON_ORD_HIST.RecordCount) * 100
         labCPB.Caption = Int(progCPB.Value) & "% Completed"
         DoEvents
         rsRECON_ORD_HIST.MoveNext
         If rsRECON_ORD_HIST.EOF = True Then Exit Do
      Loop
      DoEvents
      If rsRECON_ORD_HIST.EOF = True Then
         Exit Do
      Else
         If i < rsRECON_ORD_HIST.RecordCount Then
            i = i + 1
            progCPB.Value = (i / rsRECON_ORD_HIST.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
         End If
      End If
   Loop
   labProcessing.Caption = ""
   DoEvents
End If
Screen.MousePointer = 0
Set rsRECON_ORD_HIST = Nothing
DoEvents
Screen.MousePointer = 11
Set rsRECON_REC_HIST = New ADODB.Recordset
    rsRECON_REC_HIST.Open "select id,RRno,status from RECON_REC_HIST where status <> 'C' order by RRno asc", gconPMIOS
If Not rsRECON_REC_HIST.EOF And Not rsRECON_REC_HIST.BOF Then
   i = 0
   MsgSpeech "Checking for Duplicate Receipts..."
   Me.Caption = "Checking for Duplicate Receipts..."
   rsRECON_REC_HIST.MoveFirst
   RECON_REC_HISTkey = rsRECON_REC_HIST!rrno
   Do While Not rsRECON_REC_HIST.EOF
      varDuprecno1 = rsRECON_REC_HIST!ID
      DoEvents
      If rsRECON_REC_HIST.EOF = True Then
         Exit Do
      Else
         If i < rsRECON_REC_HIST.RecordCount Then
            i = i + 1
            progCPB.Value = (i / rsRECON_REC_HIST.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
         End If
      End If
      RECON_REC_HISTkey = rsRECON_REC_HIST!rrno
      labProcessing.Caption = "Processing: RR #" & Null2String(rsRECON_REC_HIST!rrno)
      DoEvents
      rsRECON_REC_HIST.MoveNext
      If rsRECON_REC_HIST.EOF = True Then Exit Do
      Do While rsRECON_REC_HIST!rrno = RECON_REC_HISTkey
         varDupTrantype = "'RR'"
         varDupTranno = N2Str2Null(rsRECON_REC_HIST!rrno)
         varDupFileNeym = "'RECON_REC_HIST'"
         varDuprecno2 = rsRECON_REC_HIST!ID
         varDupstatus = N2Str2Null(rsRECON_REC_HIST!Status)
         DupSql = "insert into duplicat " & _
                  "(trantype,tranno,fileneym,recno1,recno2,status)" & _
                  " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                  ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
         gconPMIOS.Execute DupSql
         RECON_REC_HISTkey = rsRECON_REC_HIST!rrno
         i = i + 1
         progCPB.Value = (i / rsRECON_REC_HIST.RecordCount) * 100
         labCPB.Caption = Int(progCPB.Value) & "% Completed"
         DoEvents
         rsRECON_REC_HIST.MoveNext
         If rsRECON_REC_HIST.EOF = True Then Exit Do
      Loop
      DoEvents
      If rsRECON_REC_HIST.EOF = True Then
         Exit Do
      Else
         If i < rsRECON_REC_HIST.RecordCount Then
            i = i + 1
            progCPB.Value = (i / rsRECON_REC_HIST.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
         End If
      End If
   Loop
   labProcessing.Caption = ""
   DoEvents
End If
Screen.MousePointer = 0
Set rsRECON_REC_HIST = Nothing
DoEvents
Screen.MousePointer = 11
Set rsRECON_PARTMAS = New ADODB.Recordset
    rsRECON_PARTMAS.Open "select id,partno from RECON_PARTMAS order by partno asc", gconPMIOS
If Not rsRECON_PARTMAS.EOF And Not rsRECON_PARTMAS.BOF Then
   i = 0
   MsgSpeech "Checking for Duplicate Part Number..."
   Me.Caption = "Checking for Duplicate Part Number..."
   rsRECON_PARTMAS.MoveFirst
   PartNokey = Null2String(rsRECON_PARTMAS!PartNo)
   Do While Not rsRECON_PARTMAS.EOF
      varDuprecno1 = rsRECON_PARTMAS!ID
      DoEvents
      If rsRECON_PARTMAS.EOF = True Then
         Exit Do
      Else
         If i < rsRECON_PARTMAS.RecordCount Then
            i = i + 1
            progCPB.Value = (i / rsRECON_PARTMAS.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
         End If
      End If
      PartNokey = Null2String(rsRECON_PARTMAS!PartNo)
      labProcessing.Caption = "Processing: Part Number " & Null2String(rsRECON_PARTMAS!PartNo)
      DoEvents
      rsRECON_PARTMAS.MoveNext
      If rsRECON_PARTMAS.EOF = True Then Exit Do
      Do While rsRECON_PARTMAS!PartNo = PartNokey
         varDupTrantype = "'MST'"
         varDupTranno = N2Str2Null(rsRECON_PARTMAS!PartNo)
         varDupFileNeym = "'MST'"
         varDuprecno2 = rsRECON_PARTMAS!ID
         DupSql = "insert into duplicat " & _
                  "(trantype,tranno,fileneym,recno1,recno2)" & _
                  " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                  ", " & varDuprecno1 & ", " & varDuprecno2 & ")"
         gconPMIOS.Execute DupSql
         PartNokey = Null2String(rsRECON_PARTMAS!PartNo)
         i = i + 1
         progCPB.Value = (i / rsRECON_PARTMAS.RecordCount) * 100
         labCPB.Caption = Int(progCPB.Value) & "% Completed"
         DoEvents
         rsRECON_PARTMAS.MoveNext
         If rsRECON_PARTMAS.EOF = True Then Exit Do
      Loop
      DoEvents
      If rsRECON_PARTMAS.EOF = True Then
         Exit Do
      Else
         If i < rsRECON_PARTMAS.RecordCount Then
            i = i + 1
            progCPB.Value = (i / rsRECON_PARTMAS.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
         End If
      End If
   Loop
   labProcessing.Caption = ""
   DoEvents
End If
Screen.MousePointer = 0
Set rsRECON_PARTMAS = Nothing
DoEvents
Screen.MousePointer = 11
Set rsRECON_DAYTRAN = New ADODB.Recordset
    rsRECON_DAYTRAN.Open "select id,trantype,tranno,itemno,status from RECON_DAYTRAN where trantype <> 'ADJ' and status <> 'C' order by trantype,tranno asc", gconPMIOS
If Not rsRECON_DAYTRAN.EOF And Not rsRECON_DAYTRAN.BOF Then
   i = 0
   MsgSpeech "Checking for Duplicate Entry in Daily transactions File..."
   Me.Caption = "Checking for Duplicate Entry in RECON_DAYTRAN File..."
   rsRECON_DAYTRAN.MoveFirst
   RECON_DAYTRANkey = rsRECON_DAYTRAN!trantype & rsRECON_DAYTRAN!tranno & rsRECON_DAYTRAN!itemno
   Do While Not rsRECON_DAYTRAN.EOF
      varDuprecno1 = rsRECON_DAYTRAN!ID
      DoEvents
      If rsRECON_DAYTRAN.EOF = True Then
         Exit Do
      Else
         If i < rsRECON_DAYTRAN.RecordCount Then
            i = i + 1
            progCPB.Value = (i / rsRECON_DAYTRAN.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
         End If
      End If
      RECON_DAYTRANkey = rsRECON_DAYTRAN!trantype & rsRECON_DAYTRAN!tranno & rsRECON_DAYTRAN!itemno
      labProcessing.Caption = "Processing: " & Null2String(rsRECON_DAYTRAN!trantype) & " #" & Null2String(rsRECON_DAYTRAN!tranno)
      DoEvents
      rsRECON_DAYTRAN.MoveNext
      If rsRECON_DAYTRAN.EOF = True Then Exit Do
      Do While rsRECON_DAYTRAN!trantype & rsRECON_DAYTRAN!tranno & rsRECON_DAYTRAN!itemno = RECON_DAYTRANkey
         varDupTrantype = N2Str2Null(rsRECON_DAYTRAN!trantype)
         varDupTranno = N2Str2Null(rsRECON_DAYTRAN!tranno)
         varDupFileNeym = "'RECON_DAYTRAN'"
         varDuprecno2 = rsRECON_DAYTRAN!ID
         varDupstatus = N2Str2Null(rsRECON_DAYTRAN!Status)
         DupSql = "insert into duplicat " & _
                  "(trantype,tranno,fileneym,recno1,recno2,status)" & _
                  " values (" & varDupTrantype & ", " & varDupTranno & ", " & varDupFileNeym & _
                  ", " & varDuprecno1 & ", " & varDuprecno2 & ", " & varDupstatus & ")"
         gconPMIOS.Execute DupSql
         RECON_DAYTRANkey = rsRECON_DAYTRAN!trantype & rsRECON_DAYTRAN!tranno & rsRECON_DAYTRAN!itemno
         i = i + 1
         progCPB.Value = (i / rsRECON_DAYTRAN.RecordCount) * 100
         labCPB.Caption = Int(progCPB.Value) & "% Completed"
         DoEvents
         rsRECON_DAYTRAN.MoveNext
      Loop
      DoEvents
      If rsRECON_DAYTRAN.EOF = True Then
         Exit Do
      Else
         If i < rsRECON_DAYTRAN.RecordCount Then
            i = i + 1
            progCPB.Value = (i / rsRECON_DAYTRAN.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
         End If
      End If
   Loop
   labProcessing.Caption = ""
   DoEvents
End If
Screen.MousePointer = 0
Set rsRECON_DAYTRAN = Nothing
End Sub

Sub CheckMatchRec()
Dim RECON_ORD_HISTkey, RECON_REC_HISTkey, PartNokey As String
Dim Ship5key, RECON_DAYTRANkey As String

Dim varDupTrantype, varDupTranno, varDupFileNeym As String
Dim varDuprecno1, varDuprecno2 As Long
Dim varDupstatus As String

Dim i As Long
Me.Caption = "CHECK MATCH RECORDS"
Screen.MousePointer = 11
DoEvents
Set rsRECON_DAYTRAN = New ADODB.Recordset
    rsRECON_DAYTRAN.Open "select id,trandate,tranno,trantype,status,part_sup from RECON_DAYTRAN where status <> 'C' order by id asc", gconPMIOS
If Not rsRECON_DAYTRAN.EOF And Not rsRECON_DAYTRAN.BOF Then
   rsRECON_DAYTRAN.MoveFirst
   MsgSpeech "Checking Matching Records from Daily transactions File..."
   Me.Caption = "Checking Matching Records from RECON_DAYTRAN File..."
   DoEvents
   i = 0
   Do While Not rsRECON_DAYTRAN.EOF
      labProcessing.Caption = "Processing: " & Null2String(rsRECON_DAYTRAN!trantype) & " #" & Null2String(rsRECON_DAYTRAN!tranno)
      DoEvents
      If rsRECON_DAYTRAN!trantype <> "PO" Or rsRECON_DAYTRAN!trantype <> "MID" Then
         If rsRECON_DAYTRAN!trantype = "RR" Then
            Set rsRECON_REC_HIST = New ADODB.Recordset
                rsRECON_REC_HIST.Open "select rrno from RECON_REC_HIST where rrno ='" & Format(rsRECON_DAYTRAN!tranno, "000000") & "'", gconPMIOS
            If rsRECON_REC_HIST.EOF And rsRECON_REC_HIST.BOF Then
               gconPMIOS.Execute "insert into noheader" & _
                                "(trantype,tranno,recno,stat_d)" & _
                                " values ('" & rsRECON_DAYTRAN!trantype & "', '" & rsRECON_DAYTRAN!tranno & "', " & rsRECON_DAYTRAN!ID & ", " & N2Str2Null(rsRECON_DAYTRAN!Status) & ")"
            End If
         End If
         If rsRECON_DAYTRAN!trantype = "CSH" Or rsRECON_DAYTRAN!trantype = "CHG" Or rsRECON_DAYTRAN!trantype = "RIV" Or rsRECON_DAYTRAN!trantype = "DR" Then
            Set rsRECON_ORD_HIST = New ADODB.Recordset
                rsRECON_ORD_HIST.Open "select trantype,tranno from RECON_ORD_HIST where trantype = " & N2Str2Null(rsRECON_DAYTRAN!trantype) & " and tranno =" & N2Str2Null(rsRECON_DAYTRAN!tranno), gconPMIOS
            If rsRECON_ORD_HIST.EOF And rsRECON_ORD_HIST.BOF Then
               gconPMIOS.Execute "insert into noheader" & _
                                "(trantype,tranno,recno,stat_d)" & _
                                " values (" & N2Str2Null(rsRECON_DAYTRAN!trantype) & ", " & N2Str2Null(rsRECON_DAYTRAN!tranno) & ", " & rsRECON_DAYTRAN!ID & ", " & N2Str2Null(rsRECON_DAYTRAN!Status) & ")"
            End If
         End If
      End If
      If rsRECON_DAYTRAN!trantype <> "ADB" Then
         Set rsRECON_PARTMAS = New ADODB.Recordset
             rsRECON_PARTMAS.Open "select partno from RECON_PARTMAS where partno = " & N2Str2Null(rsRECON_DAYTRAN!part_sup) & " order by partno asc", gconPMIOS
         If rsRECON_PARTMAS.EOF And rsRECON_PARTMAS.BOF Then
            gconPMIOS.Execute "insert into no_mstr" & _
                             "(trantype,tranno,recno,stat_d)" & _
                             " values (" & N2Str2Null(rsRECON_DAYTRAN!trantype) & ", " & N2Str2Null(rsRECON_DAYTRAN!tranno) & ", " & rsRECON_DAYTRAN!ID & ", " & N2Str2Null(rsRECON_DAYTRAN!Status) & ")"
         End If
      End If
      i = i + 1
      progCPB.Value = (i / rsRECON_DAYTRAN.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsRECON_DAYTRAN.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
End If
Screen.MousePointer = 0
Set rsRECON_DAYTRAN = Nothing
Set rsRECON_REC_HIST = Nothing
Set rsRECON_ORD_HIST = Nothing
Set rsRECON_PARTMAS = Nothing
DoEvents
Screen.MousePointer = 11
Set rsRECON_ORD_HIST = New ADODB.Recordset
    rsRECON_ORD_HIST.Open "select id,trantype,tranno,status from RECON_ORD_HIST  where status <> 'C' order by trantype,tranno asc", gconPMIOS
If Not rsRECON_ORD_HIST.EOF And Not rsRECON_ORD_HIST.BOF Then
   rsRECON_ORD_HIST.MoveFirst
   MsgSpeech "Checking Matching records from Issuances Header File..."
   Me.Caption = "Checking Matching records from Order Header File..."
   i = 0
   Do While Not rsRECON_ORD_HIST.EOF
      labProcessing.Caption = "Processing: " & Null2String(rsRECON_ORD_HIST!trantype) & " #" & Null2String(rsRECON_ORD_HIST!tranno)
      DoEvents
      Set rsRECON_DAYTRAN = New ADODB.Recordset
          rsRECON_DAYTRAN.Open "select trantype,tranno from RECON_DAYTRAN where trantype = " & N2Str2Null(rsRECON_ORD_HIST!trantype) & " and tranno = " & N2Str2Null(rsRECON_ORD_HIST!tranno), gconPMIOS
      If rsRECON_DAYTRAN.EOF And rsRECON_DAYTRAN.BOF Then
         gconPMIOS.Execute "insert into nodetail " & _
                          "(trantype,tranno,recno,stat_h)" & _
                          " values (" & N2Str2Null(rsRECON_ORD_HIST!trantype) & ", " & N2Str2Null(rsRECON_ORD_HIST!tranno) & ", " & rsRECON_ORD_HIST!ID & ", " & N2Str2Null(rsRECON_ORD_HIST!Status) & ")"
                                                           
      End If
      i = i + 1
      progCPB.Value = (i / rsRECON_ORD_HIST.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsRECON_ORD_HIST.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
End If
Screen.MousePointer = 0
Set rsRECON_ORD_HIST = Nothing
Set rsRECON_DAYTRAN = Nothing
DoEvents
Screen.MousePointer = 11
Set rsRECON_REC_HIST = New ADODB.Recordset
    rsRECON_REC_HIST.Open "select id,rrno,status from RECON_REC_HIST  where status <> 'C' order by rrno asc", gconPMIOS
If Not rsRECON_REC_HIST.EOF And Not rsRECON_REC_HIST.BOF Then
   rsRECON_REC_HIST.MoveFirst
   MsgSpeech "Checking Matching records from Receipts Header File..."
   Me.Caption = "Checking Matching records from Receipts Header File..."
   i = 0
   Do While Not rsRECON_REC_HIST.EOF
      labProcessing.Caption = "Processing: RR #" & Null2String(rsRECON_REC_HIST!rrno)
      DoEvents
      Set rsRECON_DAYTRAN = New ADODB.Recordset
          rsRECON_DAYTRAN.Open "select tranno from RECON_DAYTRAN where tranno = " & N2Str2Null(rsRECON_REC_HIST!rrno), gconPMIOS
      If rsRECON_DAYTRAN.EOF And rsRECON_DAYTRAN.BOF Then
         gconPMIOS.Execute "insert into nodetail" & _
                          "(trantype,tranno,recno,stat_d)" & _
                          " values ('RR', " & N2Str2Null(rsRECON_REC_HIST!rrno) & ", " & rsRECON_REC_HIST!ID & ", " & N2Str2Null(rsRECON_REC_HIST!Status) & ")"
      End If
      i = i + 1
      progCPB.Value = (i / rsRECON_REC_HIST.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsRECON_REC_HIST.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
End If
Screen.MousePointer = 0
Set rsRECON_REC_HIST = Nothing
Set rsRECON_DAYTRAN = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPMIOSCheckDupTrans = Nothing
UnloadForm Me
End Sub

VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISMAT_CreateConsPhyCNT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consolidate Physical Count"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_ConsolidatePHYCNT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5790
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
      Left            =   4770
      MouseIcon       =   "MAT_ConsolidatePHYCNT.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "MAT_ConsolidatePHYCNT.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Exit Window"
      Top             =   690
      Width           =   975
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Consolidate"
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
      Left            =   3810
      MouseIcon       =   "MAT_ConsolidatePHYCNT.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "MAT_ConsolidatePHYCNT.frx":0ACC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Consolidate"
      Top             =   690
      Width           =   975
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   30
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
         TabIndex        =   2
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
            TabIndex        =   3
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
            MICON           =   "MAT_ConsolidatePHYCNT.frx":0DEC
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "MAT_ConsolidatePHYCNT.frx":0E08
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "MAT_ConsolidatePHYCNT.frx":0E24
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
         TabIndex        =   7
         Top             =   30
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frmPMISMAT_CreateConsPhyCNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ConsolidatePhysicalCount()
    Dim i, i2                                                         As Integer

    Dim varPmasSTOCKNO                                                As String
    Dim varPmasSTOCKDESC                                              As String
    Dim varPmasLOCATION                                               As String
    Dim varPmasOnhand                                                 As Integer
    Dim varPmasQCount                                                 As Integer
    Dim varPmasVARIANCE                                               As Integer
    Dim varPmasAmark                                                  As String
    Dim varPmasADate                                                  As String
    Dim varPmasTagNo                                                  As String
    Dim varPmasDate_ISS                                               As String
    Dim varPmasMAC                                                    As Double
    Dim varPmasStatus                                                 As String
    Dim varPmasLASTUPDATE                                             As String
    Dim varPmasTime                                                   As String
    Dim varPmasGroup_No                                               As String
    Dim varPmasPrint_Stat                                             As String
    Dim varPmasUSERCODE                                               As String
    Dim varPmasTotalMAC                                               As Double
    Dim varPmasNewSTOCKNO                                             As String

    Dim rsCUTOFF                                                      As ADODB.Recordset
    Dim rsCONPHY                                                      As ADODB.Recordset
    Dim rsPHYCNT                                                      As ADODB.Recordset
    Set rsCUTOFF = New ADODB.Recordset
    rsCUTOFF.Open "select * from CUTOFF  order by STOCKNO asc", gconINVENTORY
    If Not rsCUTOFF.EOF And Not rsCUTOFF.BOF Then
        MsgSpeech "Consolidating Cut Off Master File "
        Me.Caption = "Consolidating Cut Off Master File "
        Screen.MousePointer = 11
        DoEvents
        i2 = 0
        gconINVENTORY.Execute "delete * from CONPHY"
        Do While Not rsCUTOFF.EOF
            varPmasSTOCKNO = N2Str2Null(rsCUTOFF!STOCKNO)
            Set rsCONPHY = New ADODB.Recordset
            rsCONPHY.Open "select STOCKNO from CONPHY where STOCKNO =" & varPmasSTOCKNO, gconINVENTORY
            If rsCONPHY.EOF And rsCONPHY.BOF Then
                labProcessing.Caption = "Processing Part Number: " & Null2String(rsCUTOFF!STOCKNO)
                DoEvents
                varPmasID = i + 1
                varPmasSTOCKDESC = N2Str2Null(rsCUTOFF!STOCKDESC)
                varPmasLOCATION = N2Str2Null(rsCUTOFF!Location)
                varPmasOnhand = N2Str2Zero(rsCUTOFF!ONHAND)
                varPmasQCount = 0
                varPmasVARIANCE = 0
                varPmasMAC = N2Str2Zero(rsCUTOFF!Mac)
                varPmasTotalMAC = varPmasOnhand * N2Str2Zero(rsCUTOFF!Mac)
                gconINVENTORY.Execute "insert into CONPHY " & _
                                      "(id,STOCKNO,STOCKDESC,qcount,location,onhand,mac,variance,totalmac)" & _
                                    " values (" & varPmasID & ", " & varPmasSTOCKNO & ", " & varPmasSTOCKDESC & ", " & varPmasQCount & ", " & varPmasLOCATION & ", " & varPmasOnhand & ", " & varPmasMAC & _
                                      ", " & varPmasVARIANCE & ", " & varPmasTotalMAC & ")"
                i = i + 1
            End If
            i2 = i2 + 1
            progCPB.Value = (i2 / rsCUTOFF.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsCUTOFF.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        MsgSpeech "Cut Off Master File Consolidated."
        Me.Caption = "Cut Off Master File Consolidated."
        Screen.MousePointer = 0
        DoEvents
    End If
    Set rsPHYCNT = New ADODB.Recordset
    rsPHYCNT.Open "select * from PHYCNT  order by tagno asc", gconINVENTORY
    If Not rsPHYCNT.EOF And Not rsPHYCNT.BOF Then
        MsgSpeech "Consolidating Physical Count "
        Me.Caption = "Consolidating Physical Count "
        Screen.MousePointer = 11
        DoEvents
        i = 0
        Do While Not rsPHYCNT.EOF
            varPmasID = i + 1
            labProcessing.Caption = "Processing Part Number: " & Null2String(rsPHYCNT!STOCKNO)
            DoEvents
            varPmasSTOCKNO = N2Str2Null(rsPHYCNT!STOCKNO)
            varPmasSTOCKDESC = N2Str2Null(rsPHYCNT!STOCKDESC)
            varPmasLOCATION = N2Str2Null(rsPHYCNT!Location)
            varPmasOnhand = N2Str2Zero(rsPHYCNT!ONHAND)
            varPmasQCount = N2Str2Zero(rsPHYCNT!Qcount)
            varPmasVARIANCE = N2Str2Zero(rsPHYCNT!variance)
            varPmasAmark = N2Str2Null(rsPHYCNT!Amark)
            varPmasADate = N2Date2Null(rsPHYCNT!ADate)
            varPmasTagNo = N2Str2Null(rsPHYCNT!TagNo)
            varPmasDate_ISS = N2Date2Null(rsPHYCNT!Date_ISS)
            varPmasMAC = N2Str2Zero(rsPHYCNT!Mac)
            varPmasStatus = N2Str2Null(rsPHYCNT!Status)
            varPmasLASTUPDATE = N2Date2Null(LOGDATE)
            varPmasTime = N2Str2Null(LOGTIME)
            varPmasGroup_No = N2Str2Null(rsPHYCNT!Group_No)
            varPmasPrint_Stat = N2Str2Null(rsPHYCNT!Print_Stat)
            varPmasUSERCODE = N2Str2Null(rsPHYCNT!USERCODE)
            varPmasTotalMAC = N2Str2Zero(rsPHYCNT!totalmac)
            varPmasNewSTOCKNO = N2Str2Null(rsPHYCNT!STOCKNO)
            Set rsCONPHY = New ADODB.Recordset
            Set rsCONPHY = gconINVENTORY.Execute("Select * from CONPHY where STOCKNO = " & N2Str2Null(rsPHYCNT!STOCKNO))
            If Not rsCONPHY.EOF And Not rsCONPHY.BOF Then
                If N2Str2Zero(rsCONPHY!ONHAND) - (N2Str2Zero(rsCONPHY!Qcount) + N2Str2Zero(rsPHYCNT!Qcount)) > 0 Then
                    varPmasVARIANCE = N2Str2Zero(rsCONPHY!ONHAND) - (N2Str2Zero(rsCONPHY!Qcount) + N2Str2Zero(rsPHYCNT!Qcount))
                Else
                    varPmasVARIANCE = (N2Str2Zero(rsCONPHY!Qcount) + N2Str2Zero(rsPHYCNT!Qcount)) - N2Str2Zero(rsCONPHY!ONHAND)
                End If
                gconINVENTORY.Execute "Update CONPHY Set" & _
                                    " tagno = " & varPmasTagNo & "," & _
                                    " Qcount = Qcount + " & N2Str2Zero(rsPHYCNT!Qcount) & "," & _
                                    " variance = " & varPmasVARIANCE & "," & _
                                    " location = " & varPmasLOCATION & "," & _
                                    " amark = " & varPmasAmark & "," & _
                                    " adate = " & varPmasADate & "," & _
                                    " group_no = " & varPmasGroup_No & "," & _
                                    " status = " & varPmasStatus & _
                                    " Where STOCKNO = " & N2Str2Null(rsPHYCNT!STOCKNO)
            Else
                gconINVENTORY.Execute "insert into CONPHY " & _
                                      "(id,tagno,STOCKNO,STOCKDESC,qcount,adate,location,amark,group_no,status,onhand,mac" & _
                                      ",variance,totalmac,print_stat,lastupdate,[time],usercode,newSTOCKNO)" & _
                                    " values (" & varPmasID & ", " & varPmasTagNo & ", " & varPmasSTOCKNO & ", " & varPmasSTOCKDESC & ", " & varPmasQCount & ", " & varPmasADate & ", " & varPmasLOCATION & ", " & varPmasAmark & ", " & varPmasGroup_No & ", " & varPmasStatus & ", " & varPmasOnhand & ", " & varPmasMAC & _
                                      ", " & varPmasVARIANCE & ", " & varPmasTotalMAC & ", " & varPmasPrint_Stat & ", " & varPmasLASTUPDATE & ", " & varPmasTime & ", " & varPmasUSERCODE & ", " & varPmasNewSTOCKNO & ")"
            End If
            gconINVENTORY.Execute "Update CUTOFF Set " & _
                                  "TagNo = " & N2Str2Null(rsPHYCNT!TagNo) & _
                                " Where STOCKNO = " & N2Str2Null(rsPHYCNT!STOCKNO)
            i = i + 1
            progCPB.Value = (i / rsPHYCNT.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsPHYCNT.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        MsgSpeech "Physical Count Consolidated."
        Me.Caption = "Physical Count Consolidated."
        Screen.MousePointer = 0
        DoEvents
    End If
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdCreate_Click()


    cmdCreate.Enabled = False
    cmdExit.Enabled = False
    DoEvents
    ConsolidatePhysicalCount
    LogAudit "R", "CONSOLIDATE PHY COUNT"
    cmdExit.Enabled = True
    DoEvents
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISCreateCutOffMaster = Nothing
    UnloadForm Me
End Sub


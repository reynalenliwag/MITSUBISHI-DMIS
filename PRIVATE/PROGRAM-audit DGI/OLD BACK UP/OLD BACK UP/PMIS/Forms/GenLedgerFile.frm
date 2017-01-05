VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISProcess_GenLedgerFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Ledger File"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "GenLedgerFile.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5940
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
      Left            =   4920
      MouseIcon       =   "GenLedgerFile.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "GenLedgerFile.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Exit Window"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Generate"
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
      Left            =   4080
      MouseIcon       =   "GenLedgerFile.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "GenLedgerFile.frx":0ACC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Generate Ledger File"
      Top             =   720
      Width           =   855
   End
   Begin wizProgBar.Prg prgDayTran 
      Height          =   315
      Left            =   90
      TabIndex        =   6
      Top             =   1080
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   556
      Picture         =   "GenLedgerFile.frx":0E3A
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "GenLedgerFile.frx":0E56
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
   Begin VB.PictureBox picOVERUNDER 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3420
      ScaleHeight     =   285
      ScaleWidth      =   2205
      TabIndex        =   5
      Top             =   2880
      Width           =   2265
      Begin VB.OptionButton optOver 
         Caption         =   "Over"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1290
         TabIndex        =   2
         ToolTipText     =   "More than the entered cost value"
         Top             =   60
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optUnder 
         Caption         =   "Under"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   90
         TabIndex        =   1
         ToolTipText     =   "Less than the entered cost value"
         Top             =   60
         Width           =   915
      End
   End
   Begin VB.TextBox txtTTLCostValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3450
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "Enter numerical value of the total cost. Do not use comma as separator (e.g. 500000,365420)"
      Top             =   2520
      Width           =   2235
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   30
      ScaleHeight     =   1155
      ScaleWidth      =   5865
      TabIndex        =   7
      Top             =   30
      Width           =   5865
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   8
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
            TabIndex        =   9
            ToolTipText     =   "Process progress"
            Top             =   -30
            Width           =   3525
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   556
         Picture         =   "GenLedgerFile.frx":0E72
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "GenLedgerFile.frx":0E8E
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
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   10
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   11
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
            MICON           =   "GenLedgerFile.frx":0EAA
         End
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
         Left            =   75
         TabIndex        =   13
         Top             =   30
         Width           =   5595
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Please Enter Total Cost Value:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   3585
   End
End
Attribute VB_Name = "frmPMISProcess_GenLedgerFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vLedgID                                                           As Long

Sub GenerateLedgerFile(STOCKNUMBER As Variant, STOCKDESCription As String)
    Dim I                                                             As Integer
    Dim rsCUTOFF                                                      As ADODB.Recordset
    Dim rsDAYTRAN                                                     As ADODB.Recordset
    Dim rsREC_HIST                                                    As ADODB.Recordset
    Dim rsORD_HIST                                                    As ADODB.Recordset

    Dim vStart                                                        As String
    Dim vTrandate                                                     As String
    Dim vTRANNO                                                       As String
    Dim vWHO                                                          As String
    Dim vRONO                                                         As String
    Dim vReceived                                                     As Long
    Dim vIssued                                                       As Long
    Dim vBalance                                                      As Integer
    Dim vUcost                                                        As Double
    Dim vMAC                                                          As Double
    Dim vTotalCost                                                    As Double
    Dim VTotalReceived, VTotalIssued                                  As Double
    Set rsCUTOFF = New ADODB.Recordset

    rsCUTOFF.Open "select * from cutoff where PARTNO = " & STOCKNUMBER, gconINVENTORY, adOpenKeyset, adLockReadOnly


    If Not rsCUTOFF.EOF And Not rsCUTOFF.BOF Then
        vLedgID = vLedgID + 1
        gconINVENTORY.Execute "insert into LEDGER " & _
                              "(ID,PARTNO,PARTDESC,WHO,TTLCOST,TRANDATE,RECEIVED,ISSUED,BALANCE,MAC,UCOST)" & _
                            " values (" & vLedgID & ", " & STOCKNUMBER & ", " & STOCKDESCription & ", 'BEGINNING BALANCE'" & _
                              ",0,0,0, NULL," & N2Str2Null(rsCUTOFF!LASTY_OH) & _
                              ", " & N2Str2Null(rsCUTOFF!LASTY_MAC) & ", " & N2Str2Null(rsCUTOFF!LASTY_MAC) & ")"
        If N2Str2Zero(rsCUTOFF!LASTY_OH) > 0 Then vBalance = N2Str2Zero(rsCUTOFF!LASTY_OH) Else vBalance = 0
        Set rsDAYTRAN = New ADODB.Recordset

        rsDAYTRAN.Open "select * from PMIS_DayTran where STOCK_ORD = " & STOCKNUMBER & " and (IN_OUT = 'I' OR IN_OUT = 'O') order by trandate asc,trantype desc,tranno asc", gconDMIS, adOpenKeyset, adLockReadOnly

        If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
            rsDAYTRAN.MoveFirst
            I = 0: VTotalReceived = 0: VTotalIssued = 0
            Do While Not rsDAYTRAN.EOF
                vTRANNO = "'" & Null2String(rsDAYTRAN!TranType) & "# " & Null2String(rsDAYTRAN!TRANNO) & "'"
                labProcessing.Caption = "Processing " & Null2String(rsDAYTRAN!TranType) & "# " & Null2String(rsDAYTRAN!TRANNO) & " Transaction"
                DoEvents
                vStart = N2Date2Null(rsDAYTRAN!trandate)
                vTotalCost = N2Str2Zero(rsCUTOFF!LASTY_OH) * N2Str2Zero(rsCUTOFF!LASTY_MAC)
                vMAC = N2Str2Zero(rsCUTOFF!LASTY_MAC)

                vTrandate = N2Date2Null(rsDAYTRAN!trandate)
                vUcost = N2Str2Zero(rsDAYTRAN!netcost) / N2Str2Zero(rsDAYTRAN!tranqty)
                vReceived = 0: vIssued = 0
                If Null2String(rsDAYTRAN!STATUS) = "C" Then
                    vWHO = "'*** CANCELLED ***'"
                Else
                    If Null2String(rsDAYTRAN!IN_OUT) = "I" Then
                        Set rsREC_HIST = New ADODB.Recordset
                        rsREC_HIST.Open "select rrno,recvd_from from PMIS_Rec_Hist where rrno = " & N2Str2Null(rsDAYTRAN!TRANNO), gconDMIS
                        If Not rsREC_HIST.EOF And Not rsREC_HIST.BOF Then
                            vWHO = N2Str2Null(Left(rsREC_HIST!recvd_from, 40))
                        Else
                            vWHO = "NULL"
                        End If
                        vReceived = N2Str2Zero(rsDAYTRAN!tranqty)
                        vBalance = vBalance + vReceived
                        If vTotalCost > 0 Then
                            vTotalCost = vTotalCost + N2Str2Zero(rsDAYTRAN!netcost)
                        Else
                            vTotalCost = vBalance * vUcost
                        End If
                    Else
                        Set rsORD_HIST = New ADODB.Recordset
                        rsORD_HIST.Open "select trantype,tranno,custname,RONO from PMIS_Ord_Hist where trantype = " & N2Str2Null(rsDAYTRAN!TranType) & " and tranno = " & N2Str2Null(rsDAYTRAN!TRANNO), gconDMIS
                        If Not rsORD_HIST.EOF And Not rsORD_HIST.BOF Then
                            vWHO = N2Str2Null(Left(rsORD_HIST!custname, 40))
                            vRONO = N2Str2Null(rsORD_HIST!rono)
                        Else
                            vWHO = "NULL": vRONO = "NULL"
                        End If
                        vIssued = N2Str2Zero(rsDAYTRAN!tranqty)
                        vBalance = vBalance - vIssued
                        vTotalCost = vTotalCost - N2Str2Zero(rsDAYTRAN!netcost)
                    End If
                End If
                If vTotalCost < 0 Then vTotalCost = 0
                If vBalance > 0 Then vMAC = vTotalCost / vBalance Else vMAC = 0
                I = I + 1
                vLedgID = vLedgID + 1
                gconINVENTORY.Execute "insert into LEDGER " & _
                                      "(ID,PARTNO,PARTDESC,TRANDATE,TRANNO,WHO,RECEIVED,ISSUED,BALANCE,UCOST,MAC,TTLCOST,STATUS)" & _
                                    " values (" & vLedgID & ", " & STOCKNUMBER & ", " & STOCKDESCription & ", " & vTrandate & "," & vTRANNO & "," & vWHO & _
                                      ", " & vReceived & "," & vIssued & "," & vBalance & "," & vUcost & "," & vMAC & "," & vTotalCost & ", 'Posted')"

                DoEvents
                rsDAYTRAN.MoveNext

            Loop
        End If
    End If
    labProcessing.Caption = ""
    DoEvents
    Me.Caption = "Finish Generating Ledger File!"
    Screen.MousePointer = 0
    DoEvents
    Exit Sub

ERRORCODE:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdCreate_Click()
    cmdCreate.Enabled = False
    cmdExit.Enabled = False
    picOVERUNDER.Enabled = False
    txtTTLCostValue.Enabled = False
    DoEvents
    Dim I                                                             As Integer
    Dim rsCONPHY                                                      As ADODB.Recordset
    Set rsCONPHY = New ADODB.Recordset
    Dim RCOUNT                                                        As Long
    rsCONPHY.Open "select * from CUTOFF  order by PARTNO asc", gconINVENTORY, adOpenKeyset, adLockReadOnly
    RCOUNT = rsCONPHY.RecordCount
    If Not rsCONPHY.EOF And Not rsCONPHY.BOF Then
        rsCONPHY.MoveFirst
        gconINVENTORY.Execute "DELETE * FROM LEDGER"
        vLedgID = 0
        I = 0
        Screen.MousePointer = 11
        Do While Not rsCONPHY.EOF
            Me.Caption = "Generating Ledger : " & Null2String(rsCONPHY!PARTNO)
            DoEvents
            GenerateLedgerFile N2Str2Null(rsCONPHY!PARTNO), N2Str2Null(rsCONPHY!PARTDESC)
            I = I + 1
            progCPB.Value = (I / RCOUNT) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"




            DoEvents
            rsCONPHY.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    LogAudit "G", "GEN LEDER FILE"
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
    Set frmPMIS_Physical_CreateCutOffMaster = Nothing
    UnloadForm Me
End Sub

Private Sub txtTTLCostValue_LostFocus()
    txtTTLCostValue.Text = Format(txtTTLCostValue.Text, MAXIMUM_DIGIT)
End Sub


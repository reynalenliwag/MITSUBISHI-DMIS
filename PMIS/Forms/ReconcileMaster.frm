VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmPMIOSReconcileMaster 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reconcile Master File"
   ClientHeight    =   4305
   ClientLeft      =   270
   ClientTop       =   360
   ClientWidth     =   5775
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "ReconcileMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5775
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generate Rank File"
      Height          =   765
      Left            =   2010
      MouseIcon       =   "ReconcileMaster.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "ReconcileMaster.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Update"
      Top             =   1860
      Width           =   1905
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create Stock Status"
      Height          =   765
      Left            =   2010
      MouseIcon       =   "ReconcileMaster.frx":0EDE
      MousePointer    =   99  'Custom
      Picture         =   "ReconcileMaster.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Update"
      Top             =   2670
      Width           =   1905
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Month-End Processing"
      Height          =   765
      Left            =   60
      MouseIcon       =   "ReconcileMaster.frx":1AB2
      MousePointer    =   99  'Custom
      Picture         =   "ReconcileMaster.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Update"
      Top             =   3480
      Width           =   1905
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Batch Posting"
      Height          =   765
      Left            =   60
      MouseIcon       =   "ReconcileMaster.frx":2686
      MousePointer    =   99  'Custom
      Picture         =   "ReconcileMaster.frx":2990
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Update"
      Top             =   2670
      Width           =   1905
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Update Master File"
      Height          =   765
      Left            =   90
      MouseIcon       =   "ReconcileMaster.frx":325A
      MousePointer    =   99  'Custom
      Picture         =   "ReconcileMaster.frx":3564
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Update"
      Top             =   1860
      Width           =   1905
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update Last Month Details"
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   5520
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update Last Month MAC"
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   5100
      Width           =   2445
   End
   Begin VB.CheckBox chkUpdateAdjustment 
      BackColor       =   &H00DEDFDE&
      Caption         =   "Update Master File Including Adjustment Transactions"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   1530
      Value           =   1  'Checked
      Width           =   5685
   End
   Begin MSMask.MaskEdBox txtFrom 
      Height          =   345
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "Input starting date"
      Top             =   1110
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTo 
      Height          =   345
      Left            =   2490
      TabIndex        =   1
      ToolTipText     =   "Input end date"
      Top             =   1110
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   765
      Left            =   4800
      MouseIcon       =   "ReconcileMaster.frx":3E2E
      MousePointer    =   99  'Custom
      Picture         =   "ReconcileMaster.frx":4138
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close window"
      Top             =   690
      Width           =   885
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Update"
      Height          =   765
      Left            =   4800
      MouseIcon       =   "ReconcileMaster.frx":4442
      MousePointer    =   99  'Custom
      Picture         =   "ReconcileMaster.frx":474C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Update"
      Top             =   690
      Width           =   885
   End
   Begin VB.PictureBox picCPB 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   30
      ScaleHeight     =   1095
      ScaleWidth      =   5715
      TabIndex        =   7
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
            Top             =   -30
            Width           =   3525
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   12
         ToolTipText     =   "Update progress"
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "ReconcileMaster.frx":5016
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "ReconcileMaster.frx":5032
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
         BackColor       =   &H00DEDFDE&
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
            MICON           =   "ReconcileMaster.frx":504E
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
         Left            =   60
         TabIndex        =   13
         Top             =   30
         Width           =   5595
      End
   End
   Begin wizProgBar.Prg Prg1 
      Height          =   315
      Left            =   90
      TabIndex        =   15
      ToolTipText     =   "Update progress"
      Top             =   5130
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      Picture         =   "ReconcileMaster.frx":506A
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "ReconcileMaster.frx":5086
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
   Begin wizProgBar.Prg Prg2 
      Height          =   315
      Left            =   90
      TabIndex        =   17
      ToolTipText     =   "Update progress"
      Top             =   5550
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      Picture         =   "ReconcileMaster.frx":50A2
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "ReconcileMaster.frx":50BE
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
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2100
      TabIndex        =   6
      Top             =   1110
      Width           =   765
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   1110
      Width           =   765
   End
End
Attribute VB_Name = "frmPMIOSReconcileMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTdaytran, rsPartmas, rsShipping As ADODB.Recordset
Dim rsRR_HD, rsOrd_Hd, rsORD_HIST As ADODB.Recordset
Dim rsREC_HIST, rsPO_HD, rsPO_HIST As ADODB.Recordset
Dim rsPO_Stat, rsDAYTRAN, rsNOHeader As ADODB.Recordset
Dim rsNODetail, rsNO_Mstr, rsSupplier As ADODB.Recordset

Dim vSupplier, vVatAmt, AddSql, upsql As String
Dim vTDTranno, vTDPartOrd, vTDTranType As String
Dim vTDInOut, vTDStatus As String
Dim vTotTranCost, vMAC As Double
Dim vTDRecNo, vPMRecNo As Long
Dim vPMOnhand, vPMTrecqty, vPMTissqty As Integer
Dim vPMLast_Recd, vTDTranDate As String
Dim vPMReceipts, vPMIssuances, vTDTranQTY As Integer
Dim vTDNetPrice, vTDNetCost, vTDRRNetCost, vTDRRInvAmt, vTDTranucost, vTDTranInvAmt As Double
Dim vORDTotPrice, vTDTranuprice As Double
Dim vShCurrMonth As Integer
Dim vShRecNo As Long
Dim vNetPrice, vNetCost As Double
Dim vOrdHDRecNo, vRRHDRecNo, vPOHDRecNo As Long

Private Sub cmdCheck_Click()
If IsDate(txtFrom.Text) = False Then
   MsgBox "Invalid from Date"
   Exit Sub
End If
If IsDate(txtTo.Text) = False Then
   MsgBox "Invalid To Date"
   Exit Sub
End If
Set rsTdaytran = New ADODB.Recordset
    rsTdaytran.Open "select trandate from NEW_daytran where trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "'", gconPMIOS
If rsTdaytran.EOF And rsTdaytran.BOF Then
   MsgSpeechBox "No Transactions Made from " & txtFrom.Text & " to " & txtTo.Text
   If UCase(LOGLEVEL) = "ADMIN" Or UCase(LOGLEVEL) = "AUTHOR" Then
      If MsgBoxXP("Update Anyway?", "Administrator Account", XP_YesNo, msg_Question) = False Then
         Exit Sub
      End If
   Else
      Exit Sub
   End If
End If
txtFrom.Enabled = False
txtTo.Enabled = False
cmdCheck.Enabled = False
cmdExit.Enabled = False
DoEvents
UpdateMaster
BatchPosting
MonthEndUpdate
GenRankFile
CreateStockStatus
cmdExit.Enabled = True
DoEvents
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Command3_Click()
UpdateMaster
End Sub

Private Sub Command4_Click()
BatchPosting
End Sub

Private Sub Command5_Click()
MonthEndUpdate
End Sub

Private Sub Command6_Click()
CreateStockStatus
End Sub

Private Sub Command7_Click()
GenRankFile
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
'Me.Height = 2025
Dim rsSupplier As ADODB.Recordset
Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "select supcode,supname from supplier where supname = 'PAMCOR'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsSupplier.EOF And Not rsSupplier.BOF Then
   vPAMCOR = Null2String(rsSupplier!SupCode)
End If
Set rsTdaytran = New ADODB.Recordset
    rsTdaytran.Open "select trantype from NEW_daytran where trantype = 'ADJ'", gconPMIOS
If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
   'Me.Height = 2310
   chkUpdateAdjustment.Enabled = True
End If
txtFrom.Text = firstDay(LOGDATE)
txtTo.Text = LOGDATE
Screen.MousePointer = 0
End Sub

Sub UpdateMaster()
Dim rsPartmas, rsCURPartmas, rsTdaytran, rsRR_HD As ADODB.Recordset
Dim i As Integer
Dim vTotTranCost, vTotTranInvAmt, vTotTranQTY, vTDTranQTY As Double
Dim vTDTranType, vTDTranno, vSupCode As String
Dim vVatAmt, vMAC As Double
Dim vPMOnhand As Integer
Dim vPartDesc As String
'If chkUpdateAdjustment.Value = 1 Then
   Set rsTdaytran = New ADODB.Recordset
       rsTdaytran.Open "select id,ItemNo,trantype,tranno,part_ord,tranqty,status,in_out,tranucost,traninvamt,trandate from NEW_daytran where trantype <> 'ADB' and status <> 'C' AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by trandate asc,trantype desc,tranno asc,itemno asc", gconPMIOS
       'rsTdaytran.Open "select id,ItemNo,trantype,tranno,part_ord,tranqty,status,in_out,tranucost,traninvamt,trandate from NEW_daytran where trantype <> 'ADB' and status <> 'C' AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by id asc", gconPMIOS
'Else
'   Set rsTdaytran = New ADODB.Recordset
'       rsTdaytran.Open "select id,ItemNo,trantype,tranno,part_ord,tranqty,status,in_out,tranucost from tdaytran where trantype <> 'ADJ' and trantype <> 'ADB' and status <> 'C' AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by trantype desc,trandate asc,tranno asc,itemno asc", gconPMIOS
'End If
If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
   rsTdaytran.MoveFirst
   Screen.MousePointer = 11
   DoEvents
   Me.Caption = "Updating Part Master File"
   gconPMIOS.Execute "update NEW_partmas set" & _
                     " lastm_mac = NEW_partmas.mac " & _
                     " where lastm_mac = 0 and onhand > 0"
   If Month(txtFrom.Text) = 1 Then
      gconPMIOS.Execute "update NEW_partmas set" & _
                        " onhand = NEW_partmas.lastm_oh," & _
                        " mac = NEW_partmas.lastm_mac," & _
                        " onorder = NEW_partmas.lastm_oo," & _
                        " tissqty = 0," & _
                        " trecqty = 0," & _
                        " receipts = 0," & _
                        " issuances = 0"
   Else
      gconPMIOS.Execute "update NEW_partmas set" & _
                        " onhand = NEW_partmas.lastm_oh," & _
                        " mac = NEW_partmas.lastm_mac," & _
                        " onorder = NEW_partmas.lastm_oo," & _
                        " tissqty = 0," & _
                        " trecqty = 0"
   End If
   DoEvents
   Me.Caption = "Updating Transactions to Part Master File"
   DoEvents
   i = 0
   Do While Not rsTdaytran.EOF
      gconPMIOS.Execute "update NEW_daytran set ItemNo = '" & Format(Null2String(rsTdaytran!itemno), "0000") & "' where ID = " & rsTdaytran!ID
      vTDTranDate = N2Date2Null(rsTdaytran!trandate)
      vTDTranType = Null2String(rsTdaytran!trantype)
      vTDTranno = Null2String(rsTdaytran!tranno)
      vTDTranQTY = N2Str2IntZero(rsTdaytran!tranqty)
      If N2Str2Zero(rsTdaytran!tranucost) > 0 Then
         vTotTranCost = rsTdaytran!tranucost * vTDTranQTY
      Else
         vTotTranCost = 0
      End If
      'vTotTranCost = N2Str2Zero(rsTdaytran!tranucost) * vTDTranQTY
      If N2Str2Zero(rsTdaytran!traninvamt) > 0 Then
         vTotTranInvAmt = rsTdaytran!traninvamt * vTDTranQTY
      Else
         vTotTranInvAmt = 0
      End If
      'vTotTranInvAmt = N2Str2Zero(rsTdaytran!traninvamt * vTDTranQTY)
      labProcessing.Caption = "Processing: " & Null2String(rsTdaytran!trantype) & " #" & Null2String(rsTdaytran!tranno)
      DoEvents
      Set rsPartmas = New ADODB.Recordset
      Set rsPartmas = gconPMIOS.Execute("select partno from NEW_partmas where partno = " & N2Str2Null(rsTdaytran!part_ord))
      If rsPartmas.EOF And rsPartmas.BOF Then
         Set rsCURPartmas = New ADODB.Recordset
         Set rsCURPartmas = gconPMIOS.Execute("Select partno,partdesc from partmas where partno = " & N2Str2Null(rsTdaytran!part_ord))
         If Not rsCURPartmas.EOF And Not rsCURPartmas.BOF Then
            vPartDesc = N2Str2Null(rsCURPartmas!PartDesc)
         Else
            vPartDesc = "'NO DESCRIPTION'"
         End If
         gconPMIOS.Execute ("Insert into NEW_partmas (partno,partdesc,date_entered) values (" & N2Str2Null(rsTdaytran!part_ord) & "," & vPartDesc & "," & N2Str2Null(rsTdaytran!trandate) & ")")
      End If
      Set rsPartmas = New ADODB.Recordset
          rsPartmas.Open "select id,partno,mac,tissqty,trecqty,poqty,tprqty,prqty,issuances,receipts,onhand from NEW_partmas where partno = " & N2Str2Null(rsTdaytran!part_ord), gconPMIOS
      If Not rsPartmas.EOF And Not rsPartmas.BOF Then
         If N2Str2Zero(rsPartmas!MAC) > 0 Then
            vMAC = rsPartmas!MAC
         Else
            vMAC = 0
         End If
         'vMAC = N2Str2Zero(rsPartmas!MAC)
         vPMOnhand = N2Str2IntZero(rsPartmas!Onhand)
         If Null2String(rsTdaytran!in_out) = "O" And vTDTranQTY <> 0 Then
            gconPMIOS.Execute "update NEW_partmas set " & _
                              "onhand = " & vPMOnhand - vTDTranQTY & ", " & _
                              "tissqty = " & N2Str2IntZero(rsPartmas!tissqty) + vTDTranQTY & ", " & _
                              "issuances = " & N2Str2IntZero(rsPartmas!issuances) + vTDTranQTY & _
                              " where id = " & rsPartmas!ID
            If vMAC = 0 Then
               MsgBox Null2String(rsPartmas!PartNo)
               Stop
               Set rsSupplier = New ADODB.Recordset
               Set rsSupplier = gconPMIOS.Execute("Select mac from daytran where month(trandate) = " & Month(txtFrom.Text) & " and year(trandate) = " & Year(txtFrom.Text) & " and trantype = 'RR' and part_ord = " & N2Str2Null(rsPartmas!PartNo) & " order by trandate asc")
               If Not rsSupplier.EOF And Not rsSupplier.BOF Then
                 If rsSupplier!MAC <> 0 Then
                     vMAC = rsSupplier!MAC
                     Stop
                  Else
                     Stop
                  End If
                 'Stop
               Else
                  Stop
               End If
               'Stop
               'vMAC = 2272.73
            End If
            gconPMIOS.Execute "update NEW_daytran set tranucost = " & vMAC & " where ID = " & rsTdaytran!ID
         End If
         
         If Null2String(rsTdaytran!in_out) = "I" And vTDTranQTY <> 0 Then
            Set rsRR_HD = New ADODB.Recordset
                rsRR_HD.Open "Select id,recvd_code,ds1,status,classcode,rrno from NEW_REC_HIST where rrno = '" & Format(vTDTranno, "000000") & "' and status <> 'C'", gconPMIOS
            If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
               vSupCode = Null2String(rsRR_HD!recvd_code)
               vVatAmt = N2Str2Zero(rsRR_HD!ds1)
               If Null2String(rsRR_HD!classcode) = "PCG" Or Null2String(rsRR_HD!classcode) = "PCS" Then
                  gconPMIOS.Execute "update NEW_REC_HIST set ds1 = " & vVatAmt & " where id = " & rsRR_HD!ID
                  If CheckIfNonVatSup(Null2String(rsRR_HD!recvd_code)) = False Then
                     If vSupCode <> vPAMCOR And vVatAmt <= 0 Then
                        vTotTranCost = vTotTranCost / 1.1
                     End If
                  Else
                     vTotTranCost = vTotTranInvAmt
                     If N2Str2Zero(rsTdaytran!traninvamt) > 0 Then
                        gconPMIOS.Execute ("Update NEW_daytran Set tranucost = " & rsTdaytran!traninvamt & " Where id = " & rsTdaytran!ID)
                     Else
                        gconPMIOS.Execute ("Update NEW_daytran Set tranucost = 0 Where id = " & rsTdaytran!ID)
                     End If
                     gconPMIOS.Execute ("Update NEW_REC_HIST Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                  End If
               End If
            Else
               If vTDTranType = "ADJ" Then
                  vTotTranCost = vMAC * vTDTranQTY
               End If
            End If
            If vPMOnhand <= 0 Then
               vMAC = vTotTranCost / vTDTranQTY
               gconPMIOS.Execute "update NEW_partmas set " & _
                                 "mac = " & vMAC & ", " & _
                                 "Onhand = " & vPMOnhand + vTDTranQTY & ", " & _
                                 "last_recd = " & vTDTranDate & ", " & _
                                 "trecqty = " & N2Str2IntZero(rsPartmas!trecqty) + vTDTranQTY & ", " & _
                                 "receipts = " & N2Str2IntZero(rsPartmas!receipts) + vTDTranQTY & _
                                 " where id =" & rsPartmas!ID
            Else
               vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
               gconPMIOS.Execute "update NEW_partmas set " & _
                                 "mac = " & vMAC & ", " & _
                                 "Onhand = " & vPMOnhand + vTDTranQTY & ", " & _
                                 "last_recd = " & vTDTranDate & ", " & _
                                 "trecqty = " & N2Str2IntZero(rsPartmas!trecqty) + vTDTranQTY & ", " & _
                                 "receipts = " & N2Str2IntZero(rsPartmas!receipts) + vTDTranQTY & _
                                 " where id =" & rsPartmas!ID
            End If
            gconPMIOS.Execute "update NEW_daytran set mac = " & vMAC & ", status = 'P' where id = " & rsTdaytran!ID
         End If
         
         If Null2String(rsTdaytran!trantype) = "PO" Then
            gconPMIOS.Execute "update NEW_partmas set " & _
                              "poqty = " & N2Str2IntZero(rsPartmas!poqty) + vTDTranQTY & _
                              " where id = " & rsPartmas!ID
         End If
      End If
      DoEvents
      i = i + 1
      progCPB.Value = (i / rsTdaytran.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsTdaytran.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
   Screen.MousePointer = 0
Else
   MsgSpeechBox "Error Opening Tdaytran File"
   Exit Sub
End If
Dim vTotalQty As Long
Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select id,rrno,status from NEW_REC_HIST where rrdate >= '" & txtFrom.Text & "' AND rrdate <= '" & txtTo.Text & "' order by rrno asc", gconPMIOS
If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
   rsRR_HD.MoveFirst
   i = 0
   MsgSpeech "Computing Total Quantity of Receipts..."
   Me.Caption = "Computing Total Quantity of Receipts..."
   Screen.MousePointer = 11
   DoEvents
   Do While Not rsRR_HD.EOF
      vOrdHDRecNo = rsRR_HD!ID
      labProcessing.Caption = "Processing: RR #" & Null2String(rsRR_HD!rrno)
      DoEvents
      Set rsTdaytran = New ADODB.Recordset
          rsTdaytran.Open "select id,trantype,tranno,tranqty,status,itemno from NEW_daytran where trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!rrno) & " order by itemno asc", gconPMIOS
      If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
         rsTdaytran.MoveFirst
         vTotalQty = 0
         Do While Not rsTdaytran.EOF
            vTotalQty = vTotalQty + N2Str2Zero(rsTdaytran!tranqty)
            rsTdaytran.MoveNext
         Loop
         If Null2String(rsRR_HD!Status) <> "C" Then
            gconPMIOS.Execute "update NEW_REC_HIST set TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
         End If
      End If
      i = i + 1
      progCPB.Value = (i / rsRR_HD.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsRR_HD.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
   Screen.MousePointer = 0
End If
Set rsOrd_Hd = New ADODB.Recordset
    rsOrd_Hd.Open "select id,trantype,tranno,status from NEW_ORD_HIST where trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by trantype,tranno asc", gconPMIOS
If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
   rsOrd_Hd.MoveFirst
   i = 0
   MsgSpeech "Computing Total Quantity of Issuances..."
   Me.Caption = "Computing Total Quantity of Issuances..."
   Screen.MousePointer = 11
   DoEvents
   Do While Not rsOrd_Hd.EOF
      vOrdHDRecNo = rsOrd_Hd!ID
      labProcessing.Caption = "Processing: " & Null2String(rsOrd_Hd!trantype) & " #" & Null2String(rsOrd_Hd!tranno)
      DoEvents
      Set rsTdaytran = New ADODB.Recordset
          rsTdaytran.Open "select id,trantype,tranno,tranqty,status,itemno from NEW_daytran where trantype = " & N2Str2Null(rsOrd_Hd!trantype) & " and tranno = " & N2Str2Null(rsOrd_Hd!tranno) & " order by itemno asc", gconPMIOS
      If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
         rsTdaytran.MoveFirst
         vTotalQty = 0
         Do While Not rsTdaytran.EOF
            vTotalQty = vTotalQty + N2Str2Zero(rsTdaytran!tranqty)
            rsTdaytran.MoveNext
         Loop
         If Null2String(rsOrd_Hd!Status) <> "C" Then
            gconPMIOS.Execute "update NEW_ORD_HIST set TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
         End If
      End If
      i = i + 1
      progCPB.Value = (i / rsOrd_Hd.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOrd_Hd.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
   Screen.MousePointer = 0
End If
MsgSpeechBox "Updating of Master File Completed..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
MoveKeyPress KeyCode
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
MoveKeyPress KeyCode
End Sub

Function CheckIfNonVatSup(SupplierCode As String) As Boolean
Dim rsSupplierMaster As ADODB.Recordset
Set rsSupplierMaster = New ADODB.Recordset
    rsSupplierMaster.Open "Select supcode,supname,NONVAT from supplier where supcode = '" & SupplierCode & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsSupplierMaster.EOF And Not rsSupplierMaster.BOF Then
   If Null2String(rsSupplierMaster!NONVAT) = "Y" Then
      CheckIfNonVatSup = True
   Else
      CheckIfNonVatSup = False
   End If
Else
   CheckIfNonVatSup = False
End If
End Function

Sub BatchPosting()
Dim i As Integer
Set rsTdaytran = New ADODB.Recordset
    rsTdaytran.Open "Select id,in_out,trantype,tranno,part_ord,status,tranqty,netcost,tranucost,trandate,tranuprice,traninvamt from NEW_daytran where status <> 'C'  AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by trandate asc,trantype desc,tranno asc,itemno asc", gconPMIOS
    'rsTdaytran.Open "Select id,in_out,trantype,tranno,part_ord,status,tranqty,netcost,tranucost,trandate,tranuprice,traninvamt from NEW_daytran where status <> 'C'  AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by id asc", gconPMIOS
If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
   rsTdaytran.MoveFirst
   i = 0
   Screen.MousePointer = 11
   MsgSpeech "Posting Transactions from Daily Transactions File..."
   Me.Caption = "Posting Transactions from Daily Transactions File..."
   DoEvents
   Do While Not rsTdaytran.EOF
      vTDRecNo = rsTdaytran!ID
      vTDInOut = Null2String(rsTdaytran!in_out)
      vTDTranType = Null2String(rsTdaytran!trantype)
      vTDTranno = Null2String(rsTdaytran!tranno)
      vTDPartOrd = Null2String(rsTdaytran!part_ord)
      vTDStatus = Null2String(rsTdaytran!Status)
      vTDTranQTY = N2Str2IntZero(rsTdaytran!tranqty)
      If N2Str2Zero(rsTdaytran!netcost) > 0 Then
         vTDNetCost = rsTdaytran!netcost
      Else
         vTDNetCost = N2Str2Zero(rsTdaytran!netcost)
      End If
      'vTDNetCost = N2Str2Zero(rsTdaytran!netcost)
      If N2Str2Zero(rsTdaytran!tranucost) > 0 Then
         vTDTranucost = rsTdaytran!tranucost
      Else
         vTDTranucost = N2Str2Zero(rsTdaytran!tranucost)
      End If
      'vTDTranucost = N2Str2Zero(rsTdaytran!tranucost)
      If N2Str2Zero(rsTdaytran!traninvamt) > 0 Then
         vTDTranInvAmt = rsTdaytran!traninvamt
      Else
         vTDTranInvAmt = N2Str2Zero(rsTdaytran!traninvamt)
      End If
      'vTDTranInvAmt = N2Str2Zero(rsTdaytran!traninvamt)
      vTDTranDate = Null2Date(rsTdaytran!trandate)
      vTotTranCost = vTDTranucost * vTDTranQTY
      vTDTranuprice = N2Str2Zero(rsTdaytran!tranuprice)
      labProcessing.Caption = "Processing: " & vTDTranType & " #" & vTDTranno
      DoEvents
      Set rsPartmas = New ADODB.Recordset
          rsPartmas.Open "Select id,onhand,trecqty,last_recd,receipts,tissqty,issuances,lastm_MAC,MAC from NEW_partmas where partno = '" & vTDPartOrd & "'", gconPMIOS
      If Not rsPartmas.EOF And Not rsPartmas.BOF Then
         If vTDTranType <> "ADJ" And vTDTranType <> "PO" And (vTDInOut = "I" Or vTDInOut = "O") And vTDTranQTY <> 0 And vTDStatus <> "C" Then
            If vTDTranType = "RR" Then
               Set rsRR_HD = New ADODB.Recordset
                   rsRR_HD.Open "Select recvd_code,ds1,status,classcode,rrno from NEW_REC_HIST where rrno = '" & Format(rsTdaytran!tranno, "000000") & "'", gconPMIOS
               If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                  vSupplier = Null2String(rsRR_HD!recvd_code)
                  vVatAmt = N2Str2IntZero(rsRR_HD!ds1)
                  If Null2String(rsRR_HD!classcode) = "PCG" Or Null2String(rsRR_HD!classcode) = "PCS" Then
                     If CheckIfNonVatSup(Null2String(rsRR_HD!recvd_code)) = False Then
                        If vSupplier <> vPAMCOR And vVatAmt <= 0 Then
                           vTotTranCost = vTotTranCost / 1.1
                        End If
                     Else
                        vTotTranCost = vTDTranInvAmt * vTDTranQTY
                     End If
                  End If
                  vPMRecNo = rsPartmas!ID
                  'vPMOnhand = N2Str2IntZero(rsPartmas!Onhand)
                  vPMTrecqty = N2Str2IntZero(rsPartmas!trecqty)
                  vPMLast_Recd = Null2Date(rsPartmas!Last_Recd)
                  'vPMReceipts = N2Str2IntZero(rsPartmas!receipts)
                  'vMAC = N2Str2Zero(rsPartmas!MAC)
                  gconPMIOS.Execute "update NEW_partmas set " & _
                                    "trecqty = " & vPMTrecqty - vTDTranQTY & ", " & _
                                    "last_recd = " & N2Str2Null(vTDTranDate) & _
                                    " where id =" & vPMRecNo
                  gconPMIOS.Execute "update NEW_daytran set status = 'P' where id = " & vTDRecNo
               Else
                  gconPMIOS.Execute "insert into noheader " & _
                                   "(trantype,tranno,recno,stat_h)" & _
                                   " values ('" & "RR" & "', '" & vTDTranno & "', " & vTDRecNo & ", '" & vTDStatus & "')"
                  MsgSpeechBox "Error in " & vTDTranType & "-" & vTDTranno & " does not have header File"
               End If
            End If
            If vTDInOut = "O" Then
               Set rsOrd_Hd = New ADODB.Recordset
                   rsOrd_Hd.Open "Select trantype,tranno from NEW_ORD_HIST where trantype = '" & vTDTranType & "' and tranno = '" & Format(rsTdaytran!tranno, "000000") & "'", gconPMIOS
               If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                  If vTDTranType = "CHG" Or vTDTranType = "CSH" Or vTDTranType = "RIV" Then
                     vORDTotPrice = (vTDTranuprice * vTDTranQTY) / 1.1
                  Else
                     vORDTotPrice = vTDTranuprice * vTDTranQTY
                  End If
                  vPMRecNo = rsPartmas!ID
                  vPMTissqty = N2Str2IntZero(rsPartmas!tissqty)
                  'vPMIssuances = N2Str2IntZero(rsPartmas!issuances)
                  'vMAC = N2Str2Zero(rsPartmas!MAC)
                  vTotTranCost = vTDTranucost * vTDTranQTY
                  gconPMIOS.Execute "update NEW_partmas set " & _
                                    "tissqty = " & vPMTissqty - vTDTranQTY & _
                                    " where id =" & vPMRecNo
                  gconPMIOS.Execute "update NEW_daytran set netcost = " & vTotTranCost & ", netprice = " & vORDTotPrice & ", status = 'P' where id = " & vTDRecNo
                  Set rsShipping = New ADODB.Recordset
                      rsShipping.Open "select * from NEW_shipping where partno = '" & vTDPartOrd & "'", gconPMIOS
                  If Not rsShipping.EOF And Not rsShipping.BOF Then
                     vShRecNo = rsShipping!ID
                     vShCurrMonth = N2Str2IntZero(rsShipping!curr_month)
                     gconPMIOS.Execute "update NEW_shipping set curr_month = " & vShCurrMonth + vTDTranQTY & ", " & _
                                       "freq_curr = 1 where id = " & vShRecNo
                  Else
                     gconPMIOS.Execute "insert into NEW_shipping (partno,curr_month,freq_curr)" & _
                                       " values ('" & vTDPartOrd & "', " & vTDTranQTY & ", 1)"
                  End If
               Else
                  gconPMIOS.Execute "insert into noheader " & _
                                   "(trantype,tranno,recno,stat_h)" & _
                                   " values ('" & vTDTranType & "', '" & vTDTranno & "', " & vTDRecNo & ", '" & vTDStatus & "')"
                  MsgSpeechBox "Error in " & vTDTranType & "-" & vTDTranno & " does not have header File"
               End If
            End If
         End If
         
         If vTDTranType = "ADJ" And vTDInOut = "I" And vTDTranQTY <> 0 And vTDStatus <> "C" Then
            vPMRecNo = rsPartmas!ID
            If N2Str2Zero(rsPartmas!MAC) > 0 Then
               vMAC = rsPartmas!MAC
            Else
               vMAC = N2Str2Zero(rsPartmas!MAC)
            End If
            'vMAC = N2Str2Zero(rsPartmas!MAC)
            vTotTranCost = vMAC * vTDTranQTY
            gconPMIOS.Execute "update NEW_daytran set " & _
                              "tranucost = " & vMAC & "," & _
                              "netcost = " & vTotTranCost & _
                              " where id = " & vTDRecNo
            'vPMOnhand = N2Str2IntZero(rsPartmas!Onhand)
            vPMTrecqty = N2Str2IntZero(rsPartmas!trecqty)
            
            gconPMIOS.Execute "update NEW_partmas set " & _
                              "trecqty = " & vPMTrecqty - vTDTranQTY & ", " & _
                              "last_recd = " & N2Str2Null(vTDTranDate) & _
                              " where id =" & vPMRecNo
            gconPMIOS.Execute "update NEW_daytran set mac = " & vMAC & ", status = 'P', netcost = " & vTotTranCost & " where id = " & vTDRecNo
         End If
         
         If vTDTranType = "ADJ" And vTDInOut = "O" And vTDTranQTY <> 0 And vTDStatus <> "C" Then
            vPMRecNo = rsPartmas!ID
            If N2Str2Zero(rsPartmas!MAC) > 0 Then
               vMAC = rsPartmas!MAC
            Else
               vMAC = N2Str2Zero(rsPartmas!MAC)
            End If
            'vMAC = N2Str2Zero(rsPartmas!MAC)
            vORDTotPrice = vMAC * vTDTranQTY
            vPMTissqty = N2Str2IntZero(rsPartmas!tissqty)
            vPMIssuances = N2Str2IntZero(rsPartmas!issuances)
            vTotTranCost = vMAC * vTDTranQTY
            gconPMIOS.Execute "update NEW_daytran set tranucost = " & vMAC & _
                             " where id = " & vTDRecNo
            gconPMIOS.Execute "update NEW_partmas set " & _
                              "tissqty = " & vPMTissqty - vTDTranQTY & ", " & _
                              "issuances = " & vPMIssuances - vTDTranQTY & _
                              " where id =" & vPMRecNo
            gconPMIOS.Execute "update NEW_daytran set tranucost = " & vMAC & ", netcost = " & vTotTranCost & ", netprice = " & vORDTotPrice & ", status = 'P' where id = " & vTDRecNo
         End If
         
      Else
         If vTDTranType <> "ADB" Then
            gconPMIOS.Execute "insert into no_mstr " & _
                              "(trantype,tranno,recno)" & _
                              " values ('" & vTDInOut & "', '" & vTDTranno & "', " & vTDRecNo & ")"
            MsgSpeechBox "Error in " & vTDTranType & "-" & vTDTranno & " Part Number: " & vTDPartOrd & " is not in Master File"
         End If
      End If
      i = i + 1
      progCPB.Value = (i / rsTdaytran.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsTdaytran.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
   Screen.MousePointer = 0
End If

Set rsOrd_Hd = New ADODB.Recordset
    rsOrd_Hd.Open "select id,trantype,tranno,status from NEW_ORD_HIST where trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by trantype,tranno asc", gconPMIOS
If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
   rsOrd_Hd.MoveFirst
   i = 0
   MsgSpeech "Computing Issuances Netcost and Netprice..."
   Me.Caption = "Computing Order Netcost and Netprice..."
   Screen.MousePointer = 11
   DoEvents
   Do While Not rsOrd_Hd.EOF
      vOrdHDRecNo = rsOrd_Hd!ID
      labProcessing.Caption = "Processing: " & Null2String(rsOrd_Hd!trantype) & " #" & Null2String(rsOrd_Hd!tranno)
      DoEvents
      Set rsTdaytran = New ADODB.Recordset
          rsTdaytran.Open "select id,trantype,tranno,netprice,netcost,status,itemno from NEW_daytran where trantype = " & N2Str2Null(rsOrd_Hd!trantype) & " and tranno = " & N2Str2Null(rsOrd_Hd!tranno) & " order by itemno asc", gconPMIOS
      If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
         rsTdaytran.MoveFirst
         vNetPrice = 0: vNetCost = 0
         Do While Not rsTdaytran.EOF
            If N2Str2Zero(rsTdaytran!NETprice) > 0 Then
               vTDNetPrice = rsTdaytran!NETprice
            Else
               vTDNetPrice = N2Str2Zero(rsTdaytran!NETprice)
            End If
            'vTDNetPrice = N2Str2Zero(rsTdaytran!NETprice)
            If N2Str2Zero(rsTdaytran!netcost) > 0 Then
               vTDNetCost = rsTdaytran!netcost
            Else
               vTDNetCost = N2Str2Zero(rsTdaytran!netcost)
            End If
            'vTDNetCost = N2Str2Zero(rsTdaytran!netcost)
            vTDStatus = Null2String(rsTdaytran!Status)
            If vTDStatus <> "C" Then
               vNetPrice = vNetPrice + vTDNetPrice
               vNetCost = vNetCost + vTDNetCost
            End If
            rsTdaytran.MoveNext
         Loop
         If Null2String(rsOrd_Hd!Status) <> "C" Then
            gconPMIOS.Execute "update NEW_ORD_HIST set netcost = " & vNetCost & ", netinvamt2 = " & vNetPrice & ", status = 'P' where id = " & vOrdHDRecNo
         End If
      End If
      i = i + 1
      progCPB.Value = (i / rsOrd_Hd.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOrd_Hd.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
   Screen.MousePointer = 0
End If

Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select id,rrno,recvd_code,status,classcode from NEW_REC_HIST where rrdate >= '" & txtFrom.Text & "' AND rrdate <= '" & txtTo.Text & "' order by rrno asc", gconPMIOS
If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
   rsRR_HD.MoveFirst
   i = 0
   Screen.MousePointer = 11
   MsgSpeech "Checking if details of receipts are already posted..."
   Me.Caption = "Checking if details of receipts are already posted..."
   DoEvents
   Do While Not rsRR_HD.EOF
      vRRHDRecNo = rsRR_HD!ID
      labProcessing.Caption = "Processing: RR #" & Null2String(rsRR_HD!rrno)
      DoEvents
      vTDRRNetCost = 0: vTDRRInvAmt = 0
      Set rsTdaytran = New ADODB.Recordset
          rsTdaytran.Open "select id,status,tranqty,trantype,tranno,itemno,tranucost,MAC,traninvamt from NEW_daytran where STATUS <> 'C' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!rrno) & " order by itemno asc", gconPMIOS
      If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
         rsTdaytran.MoveFirst
         Do While Not rsTdaytran.EOF
            vTDRecNo = rsTdaytran!ID
            vTDStatus = Null2String(rsTdaytran!Status)
            If N2Str2Zero(rsTdaytran!traninvamt) > 0 Then
               vTDRRInvAmt = vTDRRInvAmt + (rsTdaytran!traninvamt * N2Str2Zero(rsTdaytran!tranqty))
            Else
               vTDRRInvAmt = vTDRRInvAmt + (N2Str2Zero(rsTdaytran!traninvamt) * N2Str2Zero(rsTdaytran!tranqty))
            End If
            'vTDRRInvAmt = vTDRRInvAmt + (N2Str2Zero(rsTdaytran!traninvamt) * N2Str2Zero(rsTdaytran!tranqty))
            If N2Str2Zero(rsTdaytran!tranucost) > 0 Then
               vTDRRNetCost = vTDRRNetCost + (rsTdaytran!tranucost * N2Str2Zero(rsTdaytran!tranqty))
            Else
               vTDRRNetCost = vTDRRNetCost + (N2Str2Zero(rsTdaytran!tranucost) * N2Str2Zero(rsTdaytran!tranqty))
            End If
            'vTDRRNetCost = vTDRRNetCost + (N2Str2Zero(rsTdaytran!tranucost) * N2Str2Zero(rsTdaytran!tranqty))
            If vTDStatus <> "C" Then
               gconPMIOS.Execute "update NEW_daytran set status = 'P' where id =" & vTDRecNo
            End If
            rsTdaytran.MoveNext
         Loop
      End If
      If Null2String(rsRR_HD!Status) <> "C" Then
         If rsRR_HD!classcode = "PCG" Or rsRR_HD!classcode = "PCS" Then
            If CheckIfNonVatSup(Null2String(rsRR_HD!recvd_code)) = True Then
               gconPMIOS.Execute "update NEW_REC_HIST set ttlrramt = " & vTDRRInvAmt & ", netcost = " & vTDRRNetCost & ", status = 'P' where id = " & vRRHDRecNo
            Else
               gconPMIOS.Execute "update NEW_REC_HIST set ttlrramt = " & vTDRRInvAmt / 1.1 & ", ds_amt1 = " & vTDRRInvAmt - (vTDRRInvAmt / 1.1) & ", netrramt = " & vTDRRInvAmt & ", netcost = " & vTDRRNetCost & ", status = 'P' where id = " & vRRHDRecNo
            End If
         End If
      End If
      i = i + 1
      progCPB.Value = (i / rsRR_HD.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsRR_HD.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
   Screen.MousePointer = 0
End If
'Set rsPO_HD = New ADODB.Recordset
'    rsPO_HD.Open "select id,pono,status from NEW_PO_HIST where podate >= '" & txtFrom.Text & "' AND podate <= '" & txtTo.Text & "' order by pono asc", gconPMIOS
'If Not rsPO_HD.EOF And Not rsPO_HD.BOF Then
'   rsPO_HD.MoveFirst
'   i = 0
'   Screen.MousePointer = 11
'   MsgSpeech "Checking if details of purchases are already posted..."
'   Me.Caption = "Checking if details of purchases are already posted..."
'   DoEvents
'   Do While Not rsPO_HD.EOF
 '     vPOHDRecNo = rsPO_HD!ID
'      labProcessing.Caption = "Processing: PO #" & Null2String(rsPO_HD!pono)
'      DoEvents
'      Set rsTdaytran = New ADODB.Recordset
'          rsTdaytran.Open "select id,status,trantype,tranno,itemno from NEW_daytran where trantype = 'PO' and tranno = " & N2Str2Null(rsPO_HD!pono) & " order by itemno asc", gconPMIOS
'      If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
'         rsTdaytran.MoveFirst
'         Do While Not rsTdaytran.EOF
'            vTDRecNo = rsTdaytran!ID
'            vTDStatus = Null2String(rsTdaytran!Status)
'            If vTDStatus <> "C" Then
'               gconPMIOS.Execute "update NEW_daytran set status = 'P' where id =" & vTDRecNo
'            End If
'            If Null2String(rsPO_HD!Status) <> "C" Then
'               gconPMIOS.Execute "update NEW_PO_HIST set status = 'P' where id = " & vPOHDRecNo
'            End If
'            rsTdaytran.MoveNext
'         Loop
'      End If
'      i = i + 1
'      progCPB.Value = (i / rsPO_HD.RecordCount) * 100
'      labCPB.Caption = Int(progCPB.Value) & "% Completed"
'      DoEvents
'      rsPO_HD.MoveNext
'   Loop
'   labProcessing.Caption = ""
'   DoEvents
'   Screen.MousePointer = 0
'End If
Set rsTdaytran = New ADODB.Recordset
    rsTdaytran.Open "select id,status,trantype,tranno,itemno from NEW_daytran where trantype = 'ADJ' AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by id asc", gconPMIOS
If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
   rsTdaytran.MoveFirst
   labProcessing.Caption = "Processing: ADJ #" & Null2String(rsTdaytran!tranno)
   DoEvents
   Do While Not rsTdaytran.EOF
      vTDRecNo = rsTdaytran!ID
      vTDStatus = Null2String(rsTdaytran!Status)
      If vTDStatus = "N" Then
         gconPMIOS.Execute "update NEW_daytran set status = 'P' where id =" & vTDRecNo
      End If
      rsTdaytran.MoveNext
   Loop
End If

MsgSpeechBox "Posting of Transactions Completed..."
frmMain.mnuBatchPosting.Enabled = False
Set rsTdaytran = Nothing
Set rsPartmas = Nothing
Set rsShipping = Nothing
Set rsOrd_Hd = Nothing
Set rsRR_HD = Nothing
Set rsPO_HD = Nothing
End Sub

Sub MonthEndUpdate()
On Error Resume Next
Dim rsPartmas, rsShipping As ADODB.Recordset

Dim vPmasID As Long
Dim vPmasPartno, vPmasPartDesc As String
Dim vPmasOnHand As Long
Dim vPmasMac, vPmasMad As Double
Dim vPmasOnOrder As Long
Dim vPmasInvClass As String
Dim vPmasSStock As Long
Dim vPmasResService As Long

Dim i As Integer
Screen.MousePointer = 11
progCPB.Value = 0
DoEvents
MsgSpeech "Updating Part Master File"
Me.Caption = "Updating Part Master File"
labCPB.Caption = "Updating Part Master File... Please Wait..."
DoEvents
If Month(txtTo.Text) = 12 Then
   gconPMIOS.Execute "update NEW_partmas set" & _
                    " lastY_oh = onhand," & _
                    " lastY_mac = Mac," & _
                    " lastY_mad = Mad," & _
                    " lastY_oo = onorder"
End If
gconPMIOS.Execute "update NEW_partmas set" & _
                 " lastm_oh = onhand," & _
                 " lastm_mac = Mac," & _
                 " lastm_mad = Mad," & _
                 " lastm_oo = onorder," & _
                 " noship = noship + 1," & _
                 " mad = (Curr_Month + Prev_Month + Months_2 + Months_3 + Months_4 + Months_5) / 6 from NEW_shipping" & _
                 " where Curr_Month <= 0"
progCPB.Value = 100
DoEvents
progCPB.Value = 0
DoEvents
gconPMIOS.Execute "update NEW_partmas set" & _
                 " lastm_oh = onhand," & _
                 " lastm_mac = Mac," & _
                 " lastm_mad = Mad," & _
                 " lastm_oo = onorder," & _
                 " noship = 0," & _
                 " mad = (Curr_Month + Prev_Month + Months_2 + Months_3 + Months_4 + Months_5) / 6 from NEW_shipping" & _
                 " where Curr_Month > 0"
progCPB.Value = 100
DoEvents
Screen.MousePointer = 11
progCPB.Value = 0
DoEvents
MsgSpeech "Updating Shipping File"
Me.Caption = "Updating Shipping File"
labCPB.Caption = "Updating Shipping File... Please Wait..."
DoEvents
gconPMIOS.Execute "update NEW_shipping set" & _
                 " months_60 = Months_59, months_59 = Months_58, months_58 = Months_57, months_57 = Months_56," & _
                 " months_56 = Months_55, months_55 = Months_54, months_54 = Months_53, months_53 = Months_52," & _
                 " months_52 = Months_51, months_51 = Months_50, months_50 = Months_49, months_49 = Months_48," & _
                 " months_48 = Months_47, months_47 = Months_46, months_46 = Months_45, months_45 = Months_44," & _
                 " months_44 = Months_43, months_43 = Months_42, months_42 = Months_41, months_41 = Months_40," & _
                 " months_40 = Months_39, months_39 = Months_38, months_38 = Months_37, months_37 = Months_36," & _
                 " months_36 = Months_35, months_35 = Months_34, months_34 = Months_33, months_33 = Months_32," & _
                 " months_32 = Months_31, months_31 = Months_30, months_30 = Months_29, months_29 = Months_28," & _
                 " months_28 = Months_27, months_27 = Months_26, months_26 = Months_25, months_25 = Months_24," & _
                 " months_24 = Months_23, months_23 = Months_22, months_22 = Months_21, months_21 = Months_20," & _
                 " months_20 = Months_19, months_19 = Months_18, months_18 = Months_17, months_17 = Months_16," & _
                 " months_16 = Months_15, months_15 = Months_14, months_14 = Months_13, months_13 = Months_12," & _
                 " months_12 = Months_11, months_11 = Months_10, months_10 = Months_9, months_9 = Months_8," & _
                 " months_8 = Months_7, months_7 = Months_6, months_6 = Months_5, months_5 = Months_4," & _
                 " months_4 = Months_3, months_3 = Months_2, months_2 = Prev_Month, prev_month = Curr_Month," & _
                 " curr_month = 0 "
DoEvents
progCPB.Value = 100
Screen.MousePointer = 0
Me.Caption = "Updating Complete!"
labCPB.Caption = "Updating Complete!"
MsgSpeechBox "Month End Processing Completed!"
End Sub

Sub GenRankFile()
Dim rsPartmas As ADODB.Recordset
Dim rsPartMas2 As ADODB.Recordset
Dim rsShipping As ADODB.Recordset
Dim i, rst As Integer

Dim SMonths_12, SMonths_11, SMonths_10 As Integer
Dim SMonths_9, SMonths_8, SMonths_7 As Integer
Dim SMonths_6, SMonths_5, SMonths_4 As Integer
Dim SMonths_3, SMonths_2, SPrev_Month As Integer
Dim vTotSales, vMAD12 As Double
Dim vRankType, vSubClass As String
Dim vPrevClass, vPrevSClass As String
Dim PmasNoShip As Integer
Dim OldStock As Integer
Dim S_year1, S_year2, S_year3, S_year4, S_year5 As Integer
Dim P_Onhand As Integer
Dim P_MAC As Double
Dim P_Last_recd, P_PartDesc As String
Set rsPartmas = New ADODB.Recordset
    rsPartmas.Open "select partno,partdesc,onhand,mac,last_recd,invclass,subinvclas from NEW_partmas order by partno asc", gconPMIOS
If Not rsPartmas.EOF And Not rsPartmas.BOF Then
   rsPartmas.MoveFirst
   MsgSpeech "Generating Rank File... This may take a while... Please wait..."
   Me.Caption = "Generating Rank File"
   DoEvents
   i = 0
   Do While Not rsPartmas.EOF
      labProcessing.Caption = "Processing Part Number: " & Null2String(rsPartmas!PartNo)
      DoEvents
      SMonths_12 = 0: SMonths_11 = 0
      SMonths_10 = 0: SMonths_9 = 0
      SMonths_8 = 0:  SMonths_7 = 0
      SMonths_6 = 0:  SMonths_5 = 0
      SMonths_4 = 0:  SMonths_3 = 0
      SMonths_2 = 0:  SPrev_Month = 0
      vTotSales = 0:  vMAD12 = 0
      S_year1 = 0: S_year2 = 0: S_year3 = 0: S_year4 = 0: S_year5 = 0
      OldStock = 0
      P_Onhand = N2Str2Zero(rsPartmas!Onhand)
      If N2Str2Zero(rsPartmas!MAC) > 0 Then
         P_MAC = rsPartmas!MAC
      Else
         P_MAC = 0
      End If
      P_Last_recd = N2Date2Null(rsPartmas!Last_Recd)
      P_PartDesc = N2Str2Null(rsPartmas!PartDesc)
      vPrevClass = N2Str2Null(rsPartmas!InvClass)
      vPrevSClass = N2Str2Null(rsPartmas!SubInvClas)
      Set rsShipping = New ADODB.Recordset
          rsShipping.Open "Select * from NEW_shipping where partno = " & N2Str2Null(rsPartmas!PartNo), gconPMIOS
      If Not rsShipping.EOF And Not rsShipping.BOF Then
         SMonths_12 = N2Str2Zero(rsShipping!Months_12)
         SMonths_11 = N2Str2Zero(rsShipping!Months_11)
         SMonths_10 = N2Str2Zero(rsShipping!Months_10)
         SMonths_9 = N2Str2Zero(rsShipping!Months_9)
         SMonths_8 = N2Str2Zero(rsShipping!Months_8)
         SMonths_7 = N2Str2Zero(rsShipping!Months_7)
         SMonths_6 = N2Str2Zero(rsShipping!Months_6)
         SMonths_5 = N2Str2Zero(rsShipping!Months_5)
         SMonths_4 = N2Str2Zero(rsShipping!Months_4)
         SMonths_3 = N2Str2Zero(rsShipping!Months_3)
         SMonths_2 = N2Str2Zero(rsShipping!Months_2)
         SPrev_Month = N2Str2Zero(rsShipping!Prev_Month)
         S_year1 = N2Str2Zero(rsShipping!Months_12) + N2Str2Zero(rsShipping!Months_11) + N2Str2Zero(rsShipping!Months_10) + N2Str2Zero(rsShipping!Months_9) + N2Str2Zero(rsShipping!Months_8) + N2Str2Zero(rsShipping!Months_7) + N2Str2Zero(rsShipping!Months_6) + N2Str2Zero(rsShipping!Months_5) + N2Str2Zero(rsShipping!Months_4) + N2Str2Zero(rsShipping!Months_3) + N2Str2Zero(rsShipping!Months_2) + N2Str2Zero(rsShipping!Prev_Month)
         S_year2 = N2Str2Zero(rsShipping!months_24) + N2Str2Zero(rsShipping!months_23) + N2Str2Zero(rsShipping!months_22) + N2Str2Zero(rsShipping!months_21) + N2Str2Zero(rsShipping!months_20) + N2Str2Zero(rsShipping!months_19) + N2Str2Zero(rsShipping!months_18) + N2Str2Zero(rsShipping!months_17) + N2Str2Zero(rsShipping!months_16) + N2Str2Zero(rsShipping!months_15) + N2Str2Zero(rsShipping!months_14) + N2Str2Zero(rsShipping!months_13)
         S_year3 = N2Str2Zero(rsShipping!months_36) + N2Str2Zero(rsShipping!months_35) + N2Str2Zero(rsShipping!months_34) + N2Str2Zero(rsShipping!months_33) + N2Str2Zero(rsShipping!months_32) + N2Str2Zero(rsShipping!months_31) + N2Str2Zero(rsShipping!months_30) + N2Str2Zero(rsShipping!months_29) + N2Str2Zero(rsShipping!months_28) + N2Str2Zero(rsShipping!months_27) + N2Str2Zero(rsShipping!months_26) + N2Str2Zero(rsShipping!months_25)
         S_year4 = N2Str2Zero(rsShipping!months_48) + N2Str2Zero(rsShipping!months_47) + N2Str2Zero(rsShipping!months_46) + N2Str2Zero(rsShipping!months_45) + N2Str2Zero(rsShipping!months_44) + N2Str2Zero(rsShipping!months_43) + N2Str2Zero(rsShipping!months_42) + N2Str2Zero(rsShipping!months_41) + N2Str2Zero(rsShipping!months_40) + N2Str2Zero(rsShipping!months_39) + N2Str2Zero(rsShipping!months_38) + N2Str2Zero(rsShipping!months_37)
         S_year5 = N2Str2Zero(rsShipping!months_60) + N2Str2Zero(rsShipping!months_59) + N2Str2Zero(rsShipping!months_58) + N2Str2Zero(rsShipping!months_57) + N2Str2Zero(rsShipping!Months_56) + N2Str2Zero(rsShipping!months_55) + N2Str2Zero(rsShipping!months_54) + N2Str2Zero(rsShipping!months_53) + N2Str2Zero(rsShipping!months_52) + N2Str2Zero(rsShipping!months_51) + N2Str2Zero(rsShipping!months_50) + N2Str2Zero(rsShipping!months_49)
         vTotSales = Format(S_year1, MAXIMUM_DIGIT)
         vMAD12 = Format(vTotSales / 12, MAXIMUM_DIGIT)
      End If
         
      If vTotSales < 99999 And vTotSales > 359 Then
         vRankType = "A": vSubClass = "1"
      ElseIf vTotSales < 360 And vTotSales > 239 Then
          vRankType = "A": vSubClass = "2 "
      ElseIf vTotSales < 240 And vTotSales > 119 Then
          vRankType = "A": vSubClass = "3"
      ElseIf vTotSales < 120 And vTotSales > 47 Then
          vRankType = "B": vSubClass = ""
      ElseIf vTotSales < 48 And vTotSales > 23 Then
          vRankType = "C": vSubClass = ""
      ElseIf vTotSales < 24 And vTotSales > 0 Then
          vRankType = "D": vSubClass = ""
      Else
          If IsNull(rsPartmas!Last_Recd) = False Then
             OldStock = Int((CDate(txtTo.Text) - Null2Date(rsPartmas!Last_Recd)) / 365)
             If OldStock > 0 Then
                vRankType = "E"
                If OldStock >= 5 And S_year1 + S_year2 + S_year3 + S_year4 + S_year5 = 0 Then
                   vSubClass = "5"
                ElseIf OldStock = 4 And S_year1 + S_year2 + S_year3 + S_year4 = 0 Then vSubClass = "4"
                ElseIf OldStock = 3 And S_year1 + S_year2 + S_year3 = 0 Then vSubClass = "3"
                ElseIf OldStock = 2 And S_year1 + S_year2 = 0 Then vSubClass = "2"
                ElseIf OldStock = 1 Then vSubClass = "1"
                Else
                   If S_year1 <> 0 Then
                      vSubClass = "1"
                   ElseIf S_year1 + S_year2 <> 0 Then vSubClass = "2"
                   ElseIf S_year1 + S_year2 + S_year3 <> 0 Then vSubClass = "3"
                   ElseIf S_year1 + S_year2 + S_year3 + S_year4 <> 0 Then vSubClass = "4"
                   ElseIf S_year1 + S_year2 + S_year3 + S_year4 + S_year5 <> 0 Then vSubClass = "5"
                   End If
                End If
             Else
                vRankType = "F": vSubClass = ""
             End If
          Else
             vRankType = "E"
             If S_year1 <> 0 Then
                vSubClass = "1"
             ElseIf S_year1 + S_year2 <> 0 Then vSubClass = "2"
             ElseIf S_year1 + S_year2 + S_year3 <> 0 Then vSubClass = "3"
             ElseIf S_year1 + S_year2 + S_year3 + S_year4 <> 0 Then vSubClass = "4"
             ElseIf S_year1 + S_year2 + S_year3 + S_year4 + S_year5 <> 0 Then vSubClass = "5"
             Else
                If S_year1 + S_year2 + S_year3 + S_year4 + S_year5 = 0 Then vSubClass = "5"
             End If
          End If
      End If
      gconPMIOS.Execute "update NEW_partmas set " & _
                        "invclass = " & N2Str2Null(vRankType) & "," & _
                        "subinvclas = " & N2Str2Null(vSubClass) & "," & _
                        "mad = " & N2Str2Zero(vMAD12) & _
                        " where partno = " & N2Str2Null(rsPartmas!PartNo)
      gconPMIOS.Execute "insert into NEW_rankfle " & _
                        "(partno,partdesc,invclass,subinvclas,onhand,mad12,sales12,last_recd,mac,month_gen,prev_month,months_2,months_3,months_4,months_5,months_6,months_7,months_8,months_9,months_10,months_11,months_12,prevclass,prevsclas,date_gen)" & _
                        " values (" & N2Str2Null(rsPartmas!PartNo) & ", " & P_PartDesc & _
                        "," & N2Str2Null(vRankType) & ", " & N2Str2Null(vSubClass) & ", " & P_Onhand & _
                        "," & vMAD12 & ", " & NumericVal(vTotSales) & ", " & P_Last_recd & ", " & P_MAC & ", " & Month(txtTo.Text) & ", " & SPrev_Month & _
                        "," & SMonths_2 & ", " & SMonths_3 & ", " & SMonths_4 & _
                        "," & SMonths_5 & ", " & SMonths_6 & ", " & SMonths_7 & _
                        "," & SMonths_8 & ", " & SMonths_9 & ", " & SMonths_10 & _
                        "," & SMonths_11 & ", " & SMonths_12 & ", " & vPrevClass & ", " & vPrevSClass & ", " & N2Date2Null(txtTo.Text) & ")"
      DoEvents
      i = i + 1
      progCPB.Value = (i / rsPartmas.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsPartmas.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
   MsgSpeechBox "Generate Rank File Complete!"
Else
   MsgSpeechBox "Error opening Part Master File"
End If
End Sub

Sub CreateStockStatus()
Screen.MousePointer = 11
progCPB.Value = 0
Me.Caption = "Updating Part Master File"
labCPB.Caption = "Updating Part Master File for Stock Status... Please Wait..."
DoEvents
progCPB.Value = 100
gconPMIOS.Execute "update NEW_partmas set" & _
                 " sstock = mad * 2," & _
                 " resservice = mad" & _
                 " where invclass='A'"
gconPMIOS.Execute "update NEW_partmas set" & _
                 " sstock = mad," & _
                 " resservice = 0" & _
                 " where invclass<>'A'"
DoEvents
Screen.MousePointer = 11
progCPB.Value = 0
'gconPMIOS.Execute "delete from stkstat"
Me.Caption = "Creating Stock Status"
labCPB.Caption = "Create Stock Status Master File... Please Wait..."
DoEvents
progCPB.Value = 100
gconPMIOS.Execute "insert into NEW_stkstat " & _
                 "(partno,partdesc,onhand,mac,mad,sstock,resservice,onorder)" & _
                 " select Partno,PartDesc,OnHand,Mac,Mad,SStock,ResService,OnOrder from NEW_partmas order by partno asc"
gconPMIOS.Execute "update NEW_stkstat set date_gen = " & N2Date2Null(txtTo.Text) & " where date_gen IS NULL"
frmMain.mnuCreateStockStatus.Enabled = False
MsgSpeechBox "Create Stock Status Complete!"
Screen.MousePointer = 0
DoEvents
End Sub

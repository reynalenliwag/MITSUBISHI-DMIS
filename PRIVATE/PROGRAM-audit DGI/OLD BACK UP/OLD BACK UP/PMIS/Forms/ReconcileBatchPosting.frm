VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmPMIOSReconcileBatchPosting 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reconcile Batch Posting"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "ReconcileBatchPosting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "ReconcileBatchPosting.frx":0442
   ScaleHeight     =   1515
   ScaleWidth      =   5745
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   765
      Left            =   4740
      MouseIcon       =   "ReconcileBatchPosting.frx":317E
      MousePointer    =   99  'Custom
      Picture         =   "ReconcileBatchPosting.frx":3488
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   690
      Width           =   945
   End
   Begin VB.CommandButton cmdPost 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Post"
      Height          =   765
      Left            =   3810
      MouseIcon       =   "ReconcileBatchPosting.frx":3792
      MousePointer    =   99  'Custom
      Picture         =   "ReconcileBatchPosting.frx":3A9C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   690
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
      Picture         =   "ReconcileBatchPosting.frx":4366
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   2
      Top             =   30
      Width           =   5715
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   3
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
            TabIndex        =   4
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
         Picture         =   "ReconcileBatchPosting.frx":70A2
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   5
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   6
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
            MICON           =   "ReconcileBatchPosting.frx":9DDE
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "ReconcileBatchPosting.frx":9DFA
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "ReconcileBatchPosting.frx":9E16
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
         TabIndex        =   8
         Top             =   30
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frmPMIOSReconcileBatchPosting"
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
Dim vTDNetPrice, vTDNetCost, vTDTranucost As Double
Dim vORDTotPrice, vTDTranuprice As Double
Dim vShCurrMonth As Integer
Dim vShRecNo As Long
Dim vNetPrice, vNetCost As Double
Dim vOrdHDRecNo, vRRHDRecNo, vPOHDRecNo As Long

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPost_Click()
If MsgQuestionBox("Post All Transactions, Are You Sure?", "Batch Posting") = True Then
   cmdPost.Enabled = False
   cmdExit.Enabled = False
   BatchPosting
   cmdExit.Enabled = True
End If
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
Screen.MousePointer = 0
End Sub

Sub BatchPosting()
Dim i As Integer
Set rsTdaytran = New ADODB.Recordset
    rsTdaytran.Open "Select id,in_out,trantype,tranno,part_ord,status,tranqty,netcost,tranucost,trandate,tranuprice from RECON_daytran where status <> 'C' order by id asc", gconPMIOS
If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
   rsTdaytran.MoveFirst
   i = 0
   Screen.MousePointer = 11
   MsgSpeech "Posting Transactions from Daily Transactions File..."
   Me.Caption = "Posting Transactions from Tdaytran File..."
   DoEvents
   Do While Not rsTdaytran.EOF
      vTDRecNo = rsTdaytran!ID
      vTDInOut = Null2String(rsTdaytran!in_out)
      vTDTranType = Null2String(rsTdaytran!trantype)
      vTDTranno = Null2String(rsTdaytran!tranno)
      vTDPartOrd = Null2String(rsTdaytran!part_ord)
      vTDStatus = Null2String(rsTdaytran!Status)
      vTDTranQTY = N2Str2IntZero(rsTdaytran!tranqty)
      vTDNetCost = N2Str2Zero(rsTdaytran!netcost)
      vTDTranucost = N2Str2Zero(rsTdaytran!tranucost)
      vTDTranDate = Null2Date(rsTdaytran!trandate)
      vTotTranCost = vTDTranucost * vTDTranQTY
      vTDTranuprice = N2Str2Zero(rsTdaytran!tranuprice)
      labProcessing.Caption = "Processing: " & vTDTranType & " #" & vTDTranno
      DoEvents
      Set rsPartmas = New ADODB.Recordset
          rsPartmas.Open "Select id,onhand,trecqty,last_recd,receipts,tissqty,issuances,lastm_MAC,MAC from RECON_partmas where partno = '" & vTDPartOrd & "'", gconPMIOS
      If Not rsPartmas.EOF And Not rsPartmas.BOF Then
         If vTDTranType <> "ADJ" And vTDTranType <> "PO" And (vTDInOut = "I" Or vTDInOut = "O") And vTDTranQTY <> 0 And vTDStatus <> "C" Then
            If vTDTranType = "RR" Then
               Set rsRR_HD = New ADODB.Recordset
                   rsRR_HD.Open "Select recvd_code,ds1,status,classcode,rrno from RECON_REC_HIST where rrno = '" & Format(rsTdaytran!tranno, "000000") & "'", gconPMIOS
               If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                  vSupplier = Null2String(rsRR_HD!recvd_code)
                  vVatAmt = N2Str2IntZero(rsRR_HD!ds1)
                  If rsRR_HD!classcode = "PCG" Or rsRR_HD!classcode = "PCS" Then
                     If CheckIfNonVatSup(Null2String(rsRR_HD!recvd_code)) = False Then
                        If vSupplier <> vPAMCOR And vVatAmt <= 0 Then
                           vTotTranCost = vTotTranCost / 1.1
                        End If
                     End If
                  End If
                  vPMRecNo = rsPartmas!ID
                  vPMOnhand = N2Str2IntZero(rsPartmas!Onhand)
                  vPMTrecqty = N2Str2IntZero(rsPartmas!trecqty)
                  vPMLast_Recd = Null2Date(rsPartmas!Last_Recd)
                  vPMReceipts = N2Str2IntZero(rsPartmas!receipts)
                  vMAC = N2Str2Zero(rsPartmas!MAC)
                  gconPMIOS.Execute "update RECON_partmas set " & _
                                    "trecqty = " & vPMTrecqty - vTDTranQTY & ", " & _
                                    "last_recd = " & N2Str2Null(vTDTranDate) & _
                                    " where id =" & vPMRecNo
                  gconPMIOS.Execute "update RECON_daytran set mac = " & vMAC & ", status = 'P', netcost = " & vTotTranCost & " where id = " & vTDRecNo
               Else
                  gconPMIOS.Execute "insert into noheader " & _
                                   "(trantype,tranno,recno,stat_h)" & _
                                   " values ('" & "RR" & "', '" & vTDTranno & "', " & vTDRecNo & ", '" & vTDStatus & "')"
                  MsgSpeechBox "Error in " & vTDTranType & "-" & vTDTranno & " does not have header File"
               End If
            End If
            If vTDInOut = "O" Then
               Set rsOrd_Hd = New ADODB.Recordset
                   rsOrd_Hd.Open "Select trantype,tranno from RECON_ORD_HIST where trantype = '" & vTDTranType & "' and tranno = '" & Format(rsTdaytran!tranno, "000000") & "'", gconPMIOS
               If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                  If vTDTranType = "CHG" Or vTDTranType = "CSH" Or vTDTranType = "RIV" Then
                     vORDTotPrice = (vTDTranuprice * vTDTranQTY) / 1.1
                  Else
                     vORDTotPrice = (vTDTranuprice * vTDTranQTY)
                  End If
                  vPMRecNo = rsPartmas!ID
                  vPMTissqty = N2Str2IntZero(rsPartmas!tissqty)
                  vPMIssuances = N2Str2IntZero(rsPartmas!issuances)
                  vMAC = N2Str2Zero(rsPartmas!MAC)
                  vTotTranCost = vTDTranucost * vTDTranQTY
                  gconPMIOS.Execute "update RECON_partmas set " & _
                                    "tissqty = " & vPMTissqty - vTDTranQTY & _
                                    " where id =" & vPMRecNo
                  gconPMIOS.Execute "update RECON_daytran set tranucost = " & vMAC & ", netcost = " & vTotTranCost & ", netprice = " & vORDTotPrice & ", status = 'P' where id = " & vTDRecNo
                  Set rsShipping = New ADODB.Recordset
                      rsShipping.Open "select * from RECON_shipping where partno = '" & vTDPartOrd & "'", gconPMIOS
                  If Not rsShipping.EOF And Not rsShipping.BOF Then
                     vShRecNo = rsShipping!ID
                     vShCurrMonth = N2Str2IntZero(rsShipping!curr_month)
                     gconPMIOS.Execute "update RECON_shipping set curr_month = " & vShCurrMonth + vTDTranQTY & ", " & _
                                       "freq_curr = 1 where id = " & vShRecNo
                  Else
                     gconPMIOS.Execute "insert into RECON_shipping (partno,curr_month,freq_curr)" & _
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
            gconPMIOS.Execute "update RECON_daytran set " & _
                              "tranucost = " & N2Str2Zero(rsPartmas!MAC) & "," & _
                              "netcost = " & N2Str2Zero(rsPartmas!MAC) * vTDTranQTY & _
                              " where id = " & vTDRecNo
            vTotTranCost = N2Str2Zero(rsPartmas!MAC) * vTDTranQTY
            
            vMAC = N2Str2Zero(rsPartmas!MAC)
            vPMRecNo = rsPartmas!ID
            vPMOnhand = N2Str2IntZero(rsPartmas!Onhand)
            vPMTrecqty = N2Str2IntZero(rsPartmas!trecqty)
            
            gconPMIOS.Execute "update RECON_partmas set " & _
                              "trecqty = " & vPMTrecqty - vTDTranQTY & ", " & _
                              "last_recd = " & N2Str2Null(vTDTranDate) & _
                              " where id =" & vPMRecNo
            gconPMIOS.Execute "update tdaytran set mac = " & vMAC & ", status = 'P', netcost = " & vTotTranCost & " where id = " & vTDRecNo
         End If
         
         If vTDTranType = "ADJ" And vTDInOut = "O" And vTDTranQTY <> 0 And vTDStatus <> "C" Then
            gconPMIOS.Execute "update RECON_daytran set tranucost = " & N2Str2Zero(rsPartmas!MAC) & _
                             " where id = " & vTDRecNo
            vTotTranCost = N2Str2Zero(rsPartmas!MAC) * vTDTranQTY
         
            vPMRecNo = rsPartmas!ID
            vMAC = N2Str2Zero(rsPartmas!MAC)
            vORDTotPrice = (vMAC * vTDTranQTY)
            vPMTissqty = N2Str2IntZero(rsPartmas!tissqty)
            vPMIssuances = N2Str2IntZero(rsPartmas!issuances)
         
            gconPMIOS.Execute "update RECON_partmas set " & _
                              "tissqty = " & vPMTissqty - vTDTranQTY & ", " & _
                              "issuances = " & vPMIssuances - vTDTranQTY & _
                              " where id =" & vPMRecNo
            gconPMIOS.Execute "update RECON_daytran set tranucost = " & vMAC & ", netcost = " & vTotTranCost & ", netprice = " & vORDTotPrice & ", status = 'P' where id = " & vTDRecNo
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
    rsOrd_Hd.Open "select id,trantype,tranno,status from RECON_ORD_HIST order by trantype,tranno asc", gconPMIOS
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
          rsTdaytran.Open "select id,trantype,tranno,netprice,netcost,status,itemno from RECON_daytran where trantype = " & N2Str2Null(rsOrd_Hd!trantype) & " and tranno = " & N2Str2Null(rsOrd_Hd!tranno) & " order by itemno asc", gconPMIOS
      If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
         rsTdaytran.MoveFirst
         vNetPrice = 0: vNetCost = 0
         Do While Not rsTdaytran.EOF
            vTDNetPrice = N2Str2Zero(rsTdaytran!NETprice)
            vTDNetCost = N2Str2Zero(rsTdaytran!netcost)
            vTDStatus = Null2String(rsTdaytran!Status)
            If vTDStatus <> "C" Then
               vNetPrice = vNetPrice + vTDNetPrice
               vNetCost = vNetCost + vTDNetCost
            End If
            'MoveTdaytran (rsTdaytran!ID)
            rsTdaytran.MoveNext
         Loop
         If Null2String(rsOrd_Hd!Status) <> "C" Then
            gconPMIOS.Execute "update RECON_ORD_HIST set netcost = " & vNetCost & ", netinvamt2 = " & vNetPrice & ", status = 'P' where id = " & vOrdHDRecNo
         End If
         'MoveOrdHd (vOrdHDRecNo)
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
    rsRR_HD.Open "select id,rrno,status from RECON_REC_HIST order by rrno asc", gconPMIOS
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
      Set rsTdaytran = New ADODB.Recordset
          rsTdaytran.Open "select id,status,trantype,tranno,itemno from RECON_daytran where STATUS <> 'C' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!rrno) & " order by itemno asc", gconPMIOS
      If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
         rsTdaytran.MoveFirst
         Do While Not rsTdaytran.EOF
            vTDRecNo = rsTdaytran!ID
            vTDStatus = Null2String(rsTdaytran!Status)
            If vTDStatus <> "C" Then
               gconPMIOS.Execute "update RECON_daytran set status = 'P' where id =" & vTDRecNo
            End If
            If Null2String(rsRR_HD!Status) <> "C" Then
               gconPMIOS.Execute "update RECON_REC_HIST set status = 'P' where id = " & vRRHDRecNo
            End If
            'MoveTdaytran (rsTdaytran!ID)
            rsTdaytran.MoveNext
         Loop
         'MoveRRhd (vRRHDRecNo)
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

Set rsPO_HD = New ADODB.Recordset
    rsPO_HD.Open "select id,pono,status from RECON_PO_HIST order by pono asc", gconPMIOS
If Not rsPO_HD.EOF And Not rsPO_HD.BOF Then
   rsPO_HD.MoveFirst
   i = 0
   Screen.MousePointer = 11
   MsgSpeech "Checking if details of purchases are already posted..."
   Me.Caption = "Checking if details of purchases are already posted..."
   DoEvents
   Do While Not rsPO_HD.EOF
      vPOHDRecNo = rsPO_HD!ID
      labProcessing.Caption = "Processing: PO #" & Null2String(rsPO_HD!pono)
      DoEvents
      Set rsTdaytran = New ADODB.Recordset
          rsTdaytran.Open "select id,status,trantype,tranno,itemno from RECON_DAYTRAN where trantype = 'PO' and tranno = " & N2Str2Null(rsPO_HD!pono) & " order by itemno asc", gconPMIOS
      If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
         rsTdaytran.MoveFirst
         Do While Not rsTdaytran.EOF
            vTDRecNo = rsTdaytran!ID
            vTDStatus = Null2String(rsTdaytran!Status)
            If vTDStatus <> "C" Then
               gconPMIOS.Execute "update tdaytran set status = 'P' where id =" & vTDRecNo
            End If
            If Null2String(rsPO_HD!Status) <> "C" Then
               gconPMIOS.Execute "update po_hd set status = 'P' where id = " & vPOHDRecNo
            End If
            'MoveTdaytran (rsTdaytran!ID)
            rsTdaytran.MoveNext
         Loop
         'MovePOhd (vPOHDRecNo)
      End If
      i = i + 1
      progCPB.Value = (i / rsPO_HD.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsPO_HD.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
   Screen.MousePointer = 0
End If

Set rsTdaytran = New ADODB.Recordset
    rsTdaytran.Open "select id,status,trantype,tranno,itemno from RECON_daytran where trantype = 'ADJ' order by id asc", gconPMIOS
If Not rsTdaytran.EOF And Not rsTdaytran.BOF Then
   rsTdaytran.MoveFirst
   labProcessing.Caption = "Processing: ADJ #" & Null2String(rsTdaytran!tranno)
   DoEvents
   Do While Not rsTdaytran.EOF
      vTDRecNo = rsTdaytran!ID
      vTDStatus = Null2String(rsTdaytran!Status)
      If vTDStatus = "N" Then
         gconPMIOS.Execute "update RECON_daytran set status = 'P' where id =" & vTDRecNo
      End If
      'MoveTdaytran (rsTdaytran!ID)
      rsTdaytran.MoveNext
   Loop
End If

MsgSpeechBox "Posting of Transactions Completed..."
frmMain.mnuBatchPosting.Enabled = False
cmdPost.Enabled = False
Set rsTdaytran = Nothing
Set rsPartmas = Nothing
Set rsShipping = Nothing
Set rsOrd_Hd = Nothing
Set rsRR_HD = Nothing
Set rsPO_HD = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPMIOSBatchPosting = Nothing
UnloadForm Me
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



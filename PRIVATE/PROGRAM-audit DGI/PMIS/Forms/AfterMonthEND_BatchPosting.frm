VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizprogbar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "xpbutton.ocx"
Begin VB.Form frmPMISAfterMonthEND_BatchPosting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "After Month End Batch Posting"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AfterMonthEND_BatchPosting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5760
   Begin VB.CommandButton cmdPost 
      Caption         =   "Post"
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
      Left            =   4200
      MaskColor       =   &H0000FFFF&
      MouseIcon       =   "AfterMonthEND_BatchPosting.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "AfterMonthEND_BatchPosting.frx":0594
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Press F11 for Posting By Range"
      Top             =   750
      Width           =   705
   End
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
      Left            =   4950
      MouseIcon       =   "AfterMonthEND_BatchPosting.frx":08B9
      MousePointer    =   99  'Custom
      Picture         =   "AfterMonthEND_BatchPosting.frx":0A0B
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
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   3
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   4
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
            MICON           =   "AfterMonthEND_BatchPosting.frx":0D71
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "AfterMonthEND_BatchPosting.frx":0D8D
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "AfterMonthEND_BatchPosting.frx":0DA9
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
         Top             =   0
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frmPMISAfterMonthEND_BatchPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTDAYTRAN, rsPartMas, rsShipping  As ADODB.Recordset
Attribute rsPartMas.VB_VarUserMemId = 1073938432
Attribute rsShipping.VB_VarUserMemId = 1073938432
Dim rsRR_HD, rsOrd_Hd, rsORD_HIST      As ADODB.Recordset
Attribute rsRR_HD.VB_VarUserMemId = 1073938435
Attribute rsOrd_Hd.VB_VarUserMemId = 1073938435
Attribute rsORD_HIST.VB_VarUserMemId = 1073938435
Dim rsREC_HIST, rsPO_HD, rsPO_HIST     As ADODB.Recordset
Attribute rsREC_HIST.VB_VarUserMemId = 1073938438
Attribute rsPO_HD.VB_VarUserMemId = 1073938438
Attribute rsPO_HIST.VB_VarUserMemId = 1073938438
Dim rsPO_Stat, rsDAYTRAN, rsNOHeader   As ADODB.Recordset
Attribute rsPO_Stat.VB_VarUserMemId = 1073938441
Attribute rsDAYTRAN.VB_VarUserMemId = 1073938441
Attribute rsNOHeader.VB_VarUserMemId = 1073938441
Dim rsNODetail, rsNO_Mstr, rsSupplier  As ADODB.Recordset
Attribute rsNODetail.VB_VarUserMemId = 1073938444
Attribute rsNO_Mstr.VB_VarUserMemId = 1073938444
Attribute rsSupplier.VB_VarUserMemId = 1073938444

Dim vSupplier, vVatAmt, AddSql, upsql  As String
Attribute vSupplier.VB_VarUserMemId = 1073938447
Attribute vVatAmt.VB_VarUserMemId = 1073938447
Attribute AddSql.VB_VarUserMemId = 1073938447
Attribute upsql.VB_VarUserMemId = 1073938447
Dim vTDTranno, vTDPartOrd, vTDTranType As String
Attribute vTDTranno.VB_VarUserMemId = 1073938451
Attribute vTDPartOrd.VB_VarUserMemId = 1073938451
Attribute vTDTranType.VB_VarUserMemId = 1073938451
Dim vTDInOut, vTDStatus                As String
Attribute vTDInOut.VB_VarUserMemId = 1073938454
Attribute vTDStatus.VB_VarUserMemId = 1073938454
Dim vTotTranCost, vMAC                 As Double
Attribute vTotTranCost.VB_VarUserMemId = 1073938456
Attribute vMAC.VB_VarUserMemId = 1073938456
Dim vTDRecNo, vPMRecNo                 As Long
Attribute vTDRecNo.VB_VarUserMemId = 1073938458
Attribute vPMRecNo.VB_VarUserMemId = 1073938458
Dim vPMOnhand, vPMTrecqty, vPMTissqty  As Integer
Attribute vPMOnhand.VB_VarUserMemId = 1073938460
Attribute vPMTrecqty.VB_VarUserMemId = 1073938460
Attribute vPMTissqty.VB_VarUserMemId = 1073938460
Dim vPMLast_Recd, vTDTranDate          As String
Attribute vPMLast_Recd.VB_VarUserMemId = 1073938463
Attribute vTDTranDate.VB_VarUserMemId = 1073938463
Dim vPMReceipts, vPMIssuances, vTDTranQTY As Integer
Attribute vPMReceipts.VB_VarUserMemId = 1073938465
Attribute vPMIssuances.VB_VarUserMemId = 1073938465
Attribute vTDTranQTY.VB_VarUserMemId = 1073938465
Dim vTDNetPrice, vTDNetCost, vTDTranucost As Double
Attribute vTDNetPrice.VB_VarUserMemId = 1073938468
Attribute vTDNetCost.VB_VarUserMemId = 1073938468
Attribute vTDTranucost.VB_VarUserMemId = 1073938468
Dim vORDTotPrice, vTDTranuprice        As Double
Attribute vORDTotPrice.VB_VarUserMemId = 1073938471
Attribute vTDTranuprice.VB_VarUserMemId = 1073938471
Dim vShCurrMonth                       As Integer
Attribute vShCurrMonth.VB_VarUserMemId = 1073938473
Dim vShRecNo                           As Long
Attribute vShRecNo.VB_VarUserMemId = 1073938474
Dim vNetPrice, vNetCost                As Double
Attribute vNetPrice.VB_VarUserMemId = 1073938475
Attribute vNetCost.VB_VarUserMemId = 1073938475
Dim vOrdHDRecNo, vRRHDRecNo, vPOHDRecNo As Long
Attribute vOrdHDRecNo.VB_VarUserMemId = 1073938477
Attribute vRRHDRecNo.VB_VarUserMemId = 1073938477
Attribute vPOHDRecNo.VB_VarUserMemId = 1073938477

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post") = False Then Exit Sub

    If MsgQuestionBox("Post All Transactions, Are You Sure?", "Batch Posting") = True Then
        cmdPost.Enabled = False
        cmdExit.Enabled = False
        BatchPosting
        LogAudit "O", "MONTH END BATCH POSTING "
        cmdExit.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
End Sub

Sub BatchPosting()
    Dim I                              As Integer
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "Select id,in_out,trantype,tranno,STOCK_ORD,status,tranqty,netcost,tranucost,trandate,tranuprice from PMIS_TdayTran where status <> 'C' and month(trandate) = 4 and year(trandate) = 2004 order by id asc", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        rsTDAYTRAN.MoveFirst
        I = 0
        Screen.MousePointer = 11
        MsgSpeech "Posting Transactions from Daily Transactions File..."
        Me.Caption = "Posting Transactions from PMIS_DayTran File..."
        DoEvents
        Do While Not rsTDAYTRAN.EOF
            vTDRecNo = rsTDAYTRAN!ID
            vTDInOut = Null2String(rsTDAYTRAN!IN_OUT)
            vTDTranType = Null2String(rsTDAYTRAN!TRANTYPE)
            vTDTranno = Null2String(rsTDAYTRAN!Tranno)
            vTDPartOrd = Null2String(rsTDAYTRAN!STOCK_ORD)
            vTDStatus = Null2String(rsTDAYTRAN!Status)
            vTDTranQTY = N2Str2IntZero(rsTDAYTRAN!tranqty)
            vTDNetCost = N2Str2Zero(rsTDAYTRAN!netcost)
            vTDTranucost = N2Str2Zero(rsTDAYTRAN!TRANUCOST)
            vTDTranDate = Null2Date(rsTDAYTRAN!trandate)
            vTotTranCost = vTDTranucost * vTDTranQTY
            vTDTranuprice = N2Str2Zero(rsTDAYTRAN!TRANUPRICE)
            labProcessing.Caption = "Processing: " & vTDTranType & " #" & vTDTranno
            DoEvents
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "Select id,onhand,trecqty,last_recd,receipts,tissqty,issuances,lastm_MAC,MAC from PMIS_STOCKMAS where STOCKNO = '" & vTDPartOrd & "'", gconDMIS
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                If vTDTranType <> "ADJ" And vTDTranType <> "PO" And (vTDInOut = "I" Or vTDInOut = "O") And vTDTranQTY <> 0 And vTDStatus <> "C" Then
                    If vTDTranType = "RR" Then
                        Set rsRR_HD = New ADODB.Recordset
                        rsRR_HD.Open "Select recvd_code,ds1,status,classcode,rrno from PMIS_Rec_Hist where rrno = '" & Format(rsTDAYTRAN!Tranno, "000000") & "'", gconDMIS
                        If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                            vSupplier = Null2String(rsRR_HD!recvd_code)
                            vVatAmt = N2Str2IntZero(rsRR_HD!ds1)
                            If rsRR_HD!classcode = "PCG" Or rsRR_HD!classcode = "PCS" Then
                                If vSupplier <> vPAMCOR And vVatAmt <= 0 Then
                                    vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(vVatAmt)
                                End If
                            End If
                            vPMRecNo = rsPartMas!ID
                            vPMOnhand = N2Str2IntZero(rsPartMas!ONHAND)
                            vPMTrecqty = N2Str2IntZero(rsPartMas!trecqty)
                            vPMLast_Recd = Null2Date(rsPartMas!last_recd)
                            vPMReceipts = N2Str2IntZero(rsPartMas!receipts)
                            vMAC = N2Str2Zero(rsPartMas!lastm_mac)
                            'gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                             '                  "trecqty = " & vPMTrecqty - vTDTranQTY & ", " & _
                             '                  "last_recd = " & N2Str2Null(vTDTranDate) & _
                             '                  " where id =" & vPMRecNo
                            'gconDMIS.Execute "update PMIS_DayTran set mac = " & vMAC & ", status = 'P', netcost = " & vTotTranCost & " where id = " & vTDRecNo
                        Else
                            'gconDMIS.Execute "insert into PMIS_NoHeader " & _
                             '                 "(trantype,tranno,recno,stat_h)" & _
                             '                 " values ('" & "RR" & "', '" & vTDTranno & "', " & vTDRecNo & ", '" & vTDStatus & "')"
                            MsgSpeechBox "Error in " & vTDTranType & "-" & vTDTranno & " does not have header File"
                        End If
                    End If
                    If vTDInOut = "O" Then
                        Set rsOrd_Hd = New ADODB.Recordset
                        rsOrd_Hd.Open "Select trantype,tranno from PMIS_Ord_Hist where trantype = '" & vTDTranType & "' and tranno = '" & Format(rsTDAYTRAN!Tranno, "000000") & "'", gconDMIS
                        If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                            If vTDTranType = "CHG" Or vTDTranType = "CSH" Or vTDTranType = "RIV" Then
                                vORDTotPrice = (vTDTranuprice * vTDTranQTY) / ConvertToBIRDecimalFormat(VAT_RATE)
                            Else
                                vORDTotPrice = (vTDTranuprice * vTDTranQTY)
                            End If
                            vPMRecNo = rsPartMas!ID
                            vPMTissqty = N2Str2IntZero(rsPartMas!TISSQTY)
                            vPMIssuances = N2Str2IntZero(rsPartMas!issuances)
                            vMAC = N2Str2Zero(rsPartMas!lastm_mac)
                            vTotTranCost = vTDTranucost * vTDTranQTY
                            'gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                             '                  "tissqty = " & vPMTissqty - vTDTranQTY & _
                             '                  " where id =" & vPMRecNo
                            gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & vMAC & ", netcost = " & vTotTranCost & ", netprice = " & vORDTotPrice & ", status = 'P' where id = " & vTDRecNo
                            Set rsShipping = New ADODB.Recordset
                            rsShipping.Open "select * from PMIS_Shipping where STOCKNO = '" & vTDPartOrd & "'", gconDMIS
                            If Not rsShipping.EOF And Not rsShipping.BOF Then
                                vShRecNo = rsShipping!ID
                                vShCurrMonth = N2Str2IntZero(rsShipping!curr_month)
                                'gconDMIS.Execute "update PMIS_Shipping set curr_month = " & vShCurrMonth + vTDTranQTY & ", " & _
                                 '                  "freq_curr = 1 where id = " & vShRecNo
                            Else
                                'gconDMIS.Execute "insert into PMIS_Shipping (STOCKNO,curr_month,freq_curr)" & _
                                 '                  " values ('" & vTDPartOrd & "', " & vTDTranQTY & ", 1)"
                            End If
                        Else
                            'gconDMIS.Execute "insert into PMIS_NoHeader " & _
                             '                 "(trantype,tranno,recno,stat_h)" & _
                             '                 " values ('" & vTDTranType & "', '" & vTDTranno & "', " & vTDRecNo & ", '" & vTDStatus & "')"
                            MsgSpeechBox "Error in " & vTDTranType & "-" & vTDTranno & " does not have header File"
                        End If
                    End If
                End If

                If vTDTranType = "ADJ" And vTDInOut = "I" And vTDTranQTY <> 0 And vTDStatus <> "C" Then
                    'gconDMIS.Execute "update PMIS_TdayTran set " & _
                     '                  "tranucost = " & N2Str2Zero(rsPartMas!lastm_Mac) & "," & _
                     '                  "netcost = " & N2Str2Zero(rsPartMas!lastm_Mac) * vTDTranQTY & _
                     '                  " where id = " & vTDRecNo
                    vTotTranCost = N2Str2Zero(rsPartMas!Mac) * vTDTranQTY

                    vMAC = N2Str2Zero(rsPartMas!Mac)
                    vPMRecNo = rsPartMas!ID
                    vPMOnhand = N2Str2IntZero(rsPartMas!ONHAND)
                    vPMTrecqty = N2Str2IntZero(rsPartMas!trecqty)

                    'gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                     '                  "trecqty = " & vPMTrecqty - vTDTranQTY & ", " & _
                     '                  "last_recd = " & N2Str2Null(vTDTranDate) & _
                     '                  " where id =" & vPMRecNo
                    'gconDMIS.Execute "update PMIS_TdayTran set mac = " & vMAC & ", status = 'P', netcost = " & vTotTranCost & " where id = " & vTDRecNo
                End If

                If vTDTranType = "ADJ" And vTDInOut = "O" And vTDTranQTY <> 0 And vTDStatus <> "C" Then
                    'gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & N2Str2Zero(rsPartMas!Mac) & _
                     '                 " where id = " & vTDRecNo
                    vTotTranCost = N2Str2Zero(rsPartMas!Mac) * vTDTranQTY

                    vPMRecNo = rsPartMas!ID
                    vMAC = N2Str2Zero(rsPartMas!Mac)
                    vORDTotPrice = (vMAC * vTDTranQTY)
                    vPMTissqty = N2Str2IntZero(rsPartMas!TISSQTY)
                    vPMIssuances = N2Str2IntZero(rsPartMas!issuances)

                    'gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                     '                  "tissqty = " & vPMTissqty - vTDTranQTY & ", " & _
                     '                  "issuances = " & vPMIssuances - vTDTranQTY & _
                     '                  " where id =" & vPMRecNo
                    'gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & vMAC & ", netcost = " & vTotTranCost & ", netprice = " & vORDTotPrice & ", status = 'P' where id = " & vTDRecNo
                End If

            Else
                'gconDMIS.Execute "insert into PMIS_No_Mstr " & _
                 '                 "(trantype,tranno,recno)" & _
                 '                 " values ('" & vTDInOut & "', '" & vTDTranno & "', " & vTDRecNo & ")"
                MsgSpeechBox "Error in " & vTDTranType & "-" & vTDTranno & " Part Number: " & vTDPartOrd & " is not in Master File"
            End If
            I = I + 1
            progCPB.Value = (I / rsTDAYTRAN.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTDAYTRAN.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsOrd_Hd = New ADODB.Recordset
    rsOrd_Hd.Open "select id,trantype,tranno,status,trandate from PMIS_Ord_Hist where month(trandate) = 4 and year(trandate) = 2004 order by trantype,tranno asc", gconDMIS
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        rsOrd_Hd.MoveFirst
        I = 0
        MsgSpeech "Computing Issuances Netcost and Netprice..."
        Me.Caption = "Computing Order Netcost and Netprice..."
        Screen.MousePointer = 11
        DoEvents
        Do While Not rsOrd_Hd.EOF
            vOrdHDRecNo = rsOrd_Hd!ID
            labProcessing.Caption = "Processing: " & Null2String(rsOrd_Hd!TRANTYPE) & " #" & Null2String(rsOrd_Hd!Tranno)
            DoEvents
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select id,trantype,tranno,netprice,netcost,status,itemno from PMIS_DayTran where trantype = " & N2Str2Null(rsOrd_Hd!TRANTYPE) & " and tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " order by itemno asc", gconDMIS
            If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
                rsTDAYTRAN.MoveFirst
                vNetPrice = 0: vNetCost = 0
                Do While Not rsTDAYTRAN.EOF
                    vTDNetPrice = N2Str2Zero(rsTDAYTRAN!NETprice)
                    vTDNetCost = N2Str2Zero(rsTDAYTRAN!netcost)
                    vTDStatus = Null2String(rsTDAYTRAN!Status)
                    If vTDStatus <> "C" Then
                        vNetPrice = vNetPrice + vTDNetPrice
                        vNetCost = vNetCost + vTDNetCost
                    End If
                    'MoveTdaytran (rsTdaytran!ID)
                    rsTDAYTRAN.MoveNext
                Loop
                If Null2String(rsOrd_Hd!Status) <> "C" Then
                    gconDMIS.Execute "update PMIS_Ord_Hist set netcost = " & vNetCost & ", netinvamt2 = " & vNetPrice & ", status = 'P' where id = " & vOrdHDRecNo
                End If
                'MoveOrdHd (vOrdHDRecNo)
            End If
            I = I + 1
            progCPB.Value = (I / rsOrd_Hd.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsOrd_Hd.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsRR_HD = New ADODB.Recordset
    'rsRR_HD.Open "select id,rrno,status,[rrdate] from PMIS_Rec_Hist month([rrdate]) = 4 and year([rrdate]) = 2004 order by rrno asc", gconDMIS
    'If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
    rsRR_HD.Open "select id,rrno,status,[rrdate] from PMIS_Rec_Hist where month([rrdate]) = 4 and year([rrdate]) = 2004 order by rrno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        rsRR_HD.MoveFirst
        I = 0
        Screen.MousePointer = 11
        MsgSpeech "Checking if details of receipts are already posted..."
        Me.Caption = "Checking if details of receipts are already posted..."
        DoEvents
        Do While Not rsRR_HD.EOF
            vRRHDRecNo = rsRR_HD!ID
            labProcessing.Caption = "Processing: RR #" & Null2String(rsRR_HD!rrno)
            DoEvents
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select id,status,trantype,tranno,itemno from PMIS_DayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!rrno) & " order by itemno asc", gconDMIS
            If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
                rsTDAYTRAN.MoveFirst
                Do While Not rsTDAYTRAN.EOF
                    vTDRecNo = rsTDAYTRAN!ID
                    vTDStatus = Null2String(rsTDAYTRAN!Status)
                    If vTDStatus <> "C" Then
                        'gconDMIS.Execute "update PMIS_TdayTran set status = 'P' where id =" & vTDRecNo
                    End If
                    If Null2String(rsRR_HD!Status) <> "C" Then
                        'gconDMIS.Execute "update PMIS_RR_Hd set status = 'P' where id = " & vRRHDRecNo
                    End If
                    'MoveTdaytran (rsTdaytran!ID)
                    rsTDAYTRAN.MoveNext
                Loop
                'MoveRRhd (vRRHDRecNo)
            End If
            I = I + 1
            progCPB.Value = (I / rsRR_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsRR_HD.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsPO_HD = New ADODB.Recordset
    rsPO_HD.Open "select id,pono,status,podate from PMIS_PO_Hist where month(podate) = 4 and year(podate) = 2004 order by pono asc", gconDMIS
    If Not rsPO_HD.EOF And Not rsPO_HD.BOF Then
        rsPO_HD.MoveFirst
        I = 0
        Screen.MousePointer = 11
        MsgSpeech "Checking if details of purchases are already posted..."
        Me.Caption = "Checking if details of purchases are already posted..."
        DoEvents
        Do While Not rsPO_HD.EOF
            vPOHDRecNo = rsPO_HD!ID
            labProcessing.Caption = "Processing: PO #" & Null2String(rsPO_HD!PONO)
            DoEvents
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select id,status,trantype,tranno,itemno from PMIS_DayTran where trantype = 'PO' and tranno = " & N2Str2Null(rsPO_HD!PONO) & " order by itemno asc", gconDMIS
            If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
                rsTDAYTRAN.MoveFirst
                Do While Not rsTDAYTRAN.EOF
                    vTDRecNo = rsTDAYTRAN!ID
                    vTDStatus = Null2String(rsTDAYTRAN!Status)
                    If vTDStatus <> "C" Then
                        'gconDMIS.Execute "update PMIS_TdayTran set status = 'P' where id =" & vTDRecNo
                    End If
                    If Null2String(rsPO_HD!Status) <> "C" Then
                        'gconDMIS.Execute "update PMIS_PO_Hd set status = 'P' where id = " & vPOHDRecNo
                    End If
                    'MoveTdaytran (rsTdaytran!ID)
                    rsTDAYTRAN.MoveNext
                Loop
                'MovePOhd (vPOHDRecNo)
            End If
            I = I + 1
            progCPB.Value = (I / rsPO_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsPO_HD.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select id,status,trantype,tranno,itemno,trandate from PMIS_DayTran where month(trandate) = 4 and year(trandate) = 2004 and trantype = 'ADJ' order by id asc", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        rsTDAYTRAN.MoveFirst
        labProcessing.Caption = "Processing: ADJ #" & Null2String(rsTDAYTRAN!Tranno)
        DoEvents
        Do While Not rsTDAYTRAN.EOF
            vTDRecNo = rsTDAYTRAN!ID
            vTDStatus = Null2String(rsTDAYTRAN!Status)
            If vTDStatus <> "C" Then
                'gconDMIS.Execute "update PMIS_TdayTran set status = 'P' where id =" & vTDRecNo
            End If
            'MoveTdaytran (rsTdaytran!ID)
            rsTDAYTRAN.MoveNext
        Loop
    End If

    MsgSpeechBox "Posting of Transactions Completed..."
    'frmMain.mnuBatchPosting.Enabled = False
    cmdPost.Enabled = False
    Set rsTDAYTRAN = Nothing
    Set rsPartMas = Nothing
    Set rsShipping = Nothing
    Set rsOrd_Hd = Nothing
    Set rsRR_HD = Nothing
    Set rsPO_HD = Nothing
End Sub

Sub MoveTdaytran(aydi As Long)
    Dim MoveSql                        As String
    Dim I                              As Integer

    Dim varTRANID, varTRANDATE, varTRANTYPE, varTRANNO As String
    Dim varITEMNO, varSTOCK_ORD, varSTOCK_SUP As String
    Dim varTRANQTY                     As Integer
    Dim varUNIT                        As String
    Dim varTRANUCOST, varTRANUPRICE, varNETCOST, varNETPRICE As Double
    Dim varSTATUS, varIN_OUT, varMATCH, varLISTED As String
    Dim varMAC, varTRANINVAMT          As Double
    Dim varUSERCODE, varLASTUPDATE, varTREMARKS As String

    Dim rsNewTdaytran                  As ADODB.Recordset
    Set rsNewTdaytran = New ADODB.Recordset
    rsNewTdaytran.Open "select * from PMIS_TdayTran where id =" & aydi, gconDMIS
    If Not rsNewTdaytran.EOF And Not rsNewTdaytran.BOF Then
        varTRANID = rsNewTdaytran!ID
        varTRANDATE = N2Str2Null(rsNewTdaytran!trandate)
        varTRANTYPE = N2Str2Null(rsNewTdaytran!TRANTYPE)
        varTRANNO = N2Str2Null(rsNewTdaytran!Tranno)
        varITEMNO = N2Str2Null(rsNewTdaytran!itemno)
        varSTOCK_ORD = N2Str2Null(rsNewTdaytran!STOCK_ORD)
        varSTOCK_SUP = N2Str2Null(rsNewTdaytran!STOCK_SUP)
        varTRANQTY = N2Str2IntZero(rsNewTdaytran!tranqty)
        varUNIT = N2Str2Null(rsNewTdaytran!unit)
        varTRANUCOST = N2Str2Zero(rsNewTdaytran!TRANUCOST)
        varTRANUPRICE = N2Str2Zero(rsNewTdaytran!TRANUPRICE)
        varNETCOST = N2Str2Zero(rsNewTdaytran!netcost)
        varNETPRICE = N2Str2Zero(rsNewTdaytran!NETprice)
        varSTATUS = N2Str2Null(rsNewTdaytran!Status)
        varIN_OUT = N2Str2Null(rsNewTdaytran!IN_OUT)
        varMATCH = N2Str2Null(rsNewTdaytran!Match)
        varLISTED = N2Str2Null(rsNewTdaytran!listed)
        varMAC = N2Str2Zero(rsNewTdaytran!Mac)
        varTRANINVAMT = N2Str2Zero(rsNewTdaytran!TRANINVAMT)
        varUSERCODE = N2Str2Null(rsNewTdaytran!usercode)
        varLASTUPDATE = N2Str2Null(rsNewTdaytran!lastupdate)
        varTREMARKS = N2Str2Null(rsNewTdaytran!TREMARKS)

        MoveSql = "INSERT into PMIS_DayTran " & _
                  "(TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,UNIT,TRANUCOST,TRANUPRICE,NETCOST,NETPRICE,STATUS,IN_OUT,LISTED,MAC,TRANINVAMT,USERCODE,LASTUPDATE,TREMARKS)" & _
                " values (" & varTRANDATE & "," & varTRANTYPE & "," & varTRANNO & "," & varITEMNO & "," & varSTOCK_ORD & "," & varSTOCK_SUP & "," & varTRANQTY & "," & varUNIT & "," & varTRANUCOST & "," & varTRANUPRICE & "," & varNETCOST & "," & varNETPRICE & "," & varSTATUS & "," & varIN_OUT & "," & varLISTED & "," & varMAC & "," & varTRANINVAMT & "," & varUSERCODE & "," & varLASTUPDATE & "," & varTREMARKS & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete from PMIS_TdayTran where id = " & varTRANID
    End If
    Set rsNewTdaytran = Nothing
End Sub

Sub MoveOrdHd(aydi As Long)
    Dim MoveSql                        As String
    Dim I                              As Integer

    Dim varOHID                        As Long
    Dim varOHTRANTYPE, varOHTRANNO, varOHTRANDATE As String
    Dim varOHCANCDATE, varOHCUSTCODE, varOHCUSTNAME As String
    Dim varOHCHARGETO, varOHRONO, varOHSALESMAN As String
    Dim varOHSMNAME, varOHTERMS        As String
    Dim varOHTTLINVAMT, varOHDS1       As Double
    Dim varOHDS_DESC1                  As String
    Dim varOHDS_AMT1, varOHNETINVAMT, varOHNETCOST As Double
    Dim varOHSTATUS                    As String
    Dim varOHNETINVAMT2, varOHNETCOST2 As Double
    Dim varOHLISTED, varOHUSERCODE, varOHLASTUPDATE As String
    Dim varOHTOTINVAMT, varOHDISCOUNT, varOHVAT As Double
    Dim varOHNETINVOICE, varOHTOTALCOST As Double
    Dim varOHREMARKS                   As String

    Dim rsNewOrd_HD                    As ADODB.Recordset
    Set rsNewOrd_HD = New ADODB.Recordset
    rsNewOrd_HD.Open "select * from PMIS_Ord_Hd where id =" & aydi, gconDMIS
    If Not rsNewOrd_HD.EOF And Not rsNewOrd_HD.BOF Then
        DoEvents
        varOHID = rsOrd_Hd!ID
        varOHTRANTYPE = N2Str2Null(rsNewOrd_HD!TRANTYPE)
        varOHTRANNO = N2Str2Null(rsNewOrd_HD!Tranno)
        varOHTRANDATE = N2Str2Null(rsNewOrd_HD!trandate)
        varOHCANCDATE = N2Str2Null(rsNewOrd_HD!cancdate)
        varOHCUSTCODE = N2Str2Null(rsNewOrd_HD!custcode)
        varOHCUSTNAME = N2Str2Null(rsNewOrd_HD!custname)
        varOHCHARGETO = N2Str2Null(rsNewOrd_HD!chargeto)
        varOHRONO = N2Str2Null(rsNewOrd_HD!rono)
        varOHSALESMAN = N2Str2Null(rsNewOrd_HD!salesman)
        varOHSMNAME = N2Str2Null(rsNewOrd_HD!smname)
        varOHTERMS = N2Str2Null(rsNewOrd_HD!Terms)
        varOHTTLINVAMT = N2Str2Zero(rsNewOrd_HD!ttlinvamt)
        varOHDS1 = N2Str2IntZero(rsNewOrd_HD!ds1)
        varOHDS_DESC1 = N2Str2Null(rsNewOrd_HD!ds_desc1)
        varOHDS_AMT1 = N2Str2Zero(rsNewOrd_HD!ds_amt1)
        varOHNETINVAMT = N2Str2Zero(rsNewOrd_HD!netinvamt)
        varOHNETCOST = N2Str2Zero(rsNewOrd_HD!netcost)
        varOHSTATUS = N2Str2Null(rsNewOrd_HD!Status)
        varOHNETINVAMT2 = N2Str2Zero(rsNewOrd_HD!NETINVAMT2)
        varOHNETCOST2 = N2Str2Zero(rsNewOrd_HD!NETCOST2)
        varOHLISTED = N2Str2Null(rsNewOrd_HD!listed)
        varOHUSERCODE = N2Str2Null(rsNewOrd_HD!usercode)
        varOHLASTUPDATE = N2Str2Null(rsNewOrd_HD!lastupdate)
        varOHTOTINVAMT = N2Str2Zero(rsNewOrd_HD!TOTINVAMT)
        varOHDISCOUNT = N2Str2Zero(rsNewOrd_HD!DISCOUNT)
        varOHVAT = N2Str2Zero(rsNewOrd_HD!Vat)
        varOHNETINVOICE = N2Str2Zero(rsNewOrd_HD!NETINVOICE)
        varOHTOTALCOST = N2Str2Zero(rsNewOrd_HD!TotalCost)
        varOHREMARKS = N2Str2Null(rsNewOrd_HD!remarks)

        MoveSql = "INSERT into PMIS_Ord_Hist " & _
                  "(TRANTYPE,TRANNO,TRANDATE,CANCDATE,CUSTCODE,CUSTNAME,CHARGETO,RONO,SALESMAN,SMNAME,TERMS,TTLINVAMT,DS1,DS_DESC1,DS_AMT1,NETINVAMT,NETCOST,STATUS,NETINVAMT2,NETCOST2,LISTED,USERCODE,LASTUPDATE,TOTINVAMT,DISCOUNT,VAT,NETINVOICE,TOTALCOST,REMARKS)" & _
                " values (" & varOHTRANTYPE & ", " & varOHTRANNO & ", " & varOHTRANDATE & ", " & varOHCANCDATE & ", " & varOHCUSTCODE & ", " & varOHCUSTNAME & ", " & varOHCHARGETO & ", " & varOHRONO & ", " & varOHSALESMAN & ", " & varOHSMNAME & ", " & varOHTERMS & ", " & varOHTTLINVAMT & ", " & varOHDS1 & ", " & varOHDS_DESC1 & ", " & varOHDS_AMT1 & ", " & varOHNETINVAMT & ", " & varOHNETCOST & ", " & varOHSTATUS & ", " & varOHNETINVAMT2 & ", " & varOHNETCOST2 & ", " & varOHLISTED & ", " & varOHUSERCODE & ", " & varOHLASTUPDATE & ", " & varOHTOTINVAMT & ", " & varOHDISCOUNT & ", " & varOHVAT & ", " & varOHNETINVOICE & ", " & varOHTOTALCOST & ", " & varOHREMARKS & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete from PMIS_Ord_Hd where id = " & varOHID
    End If
    Set rsNewOrd_HD = Nothing
End Sub

Sub MoveRRhd(aydi As Long)
    Dim MoveSql                        As String
    Dim I                              As Integer

    Dim varRRID                        As Long
    Dim varRRRRNO, varRRRRDATE, varRRCANCDATE, varRRPONO As String
    Dim varRRPODATE, varRRRECVD_CODE, varRRRECVD_FROM As String
    Dim varRRADDRESS, varRRDRNO, varRRINVNO As String
    Dim varRRCLASSCODE, varRRTERMS     As String
    Dim varRRTTLRRAMT, varRRDS1        As Double
    Dim varRRDS_DESC1                  As String
    Dim varRRDS_AMT1, varRRNETRRAMT    As Double
    Dim varRRSTATUS, varRRLISTED, varRRUSERCODE As String
    Dim varRRLASTUPDATE, varRRREMARKS  As String

    Dim rsNewRR_HD                     As ADODB.Recordset
    Set rsNewRR_HD = New ADODB.Recordset
    rsNewRR_HD.Open "select * from PMIS_RR_Hd where id =" & aydi, gconDMIS
    If Not rsNewRR_HD.EOF And Not rsNewRR_HD.BOF Then
        DoEvents
        varRRID = rsRR_HD!ID
        varRRRRNO = N2Str2Null(rsNewRR_HD!rrno)
        varRRRRDATE = N2Str2Null(rsNewRR_HD!rrdate)
        varRRCANCDATE = N2Str2Null(rsNewRR_HD!cancdate)
        varRRPONO = N2Str2Null(rsNewRR_HD!PONO)
        varRRPODATE = N2Str2Null(rsNewRR_HD!podate)
        varRRRECVD_CODE = N2Str2Null(rsNewRR_HD!recvd_code)
        varRRRECVD_FROM = N2Str2Null(rsNewRR_HD!recvd_from)
        varRRADDRESS = N2Str2Null(rsNewRR_HD!Address)
        varRRDRNO = N2Str2Null(rsNewRR_HD!drno)
        varRRINVNO = N2Str2Null(rsNewRR_HD!invno)
        varRRCLASSCODE = N2Str2Null(rsNewRR_HD!classcode)
        varRRTERMS = N2Str2Null(rsNewRR_HD!Terms)
        varRRTTLRRAMT = N2Str2Zero(rsNewRR_HD!ttlrramt)
        varRRDS1 = N2Str2IntZero(rsNewRR_HD!ds1)
        varRRDS_DESC1 = N2Str2Null(rsNewRR_HD!ds_desc1)
        varRRDS_AMT1 = N2Str2Zero(rsNewRR_HD!ds_amt1)
        varRRNETRRAMT = N2Str2Zero(rsNewRR_HD!netrramt)
        varRRSTATUS = N2Str2Null(rsNewRR_HD!Status)
        varRRLISTED = N2Str2Null(rsNewRR_HD!listed)
        varRRUSERCODE = N2Str2Null(rsNewRR_HD!usercode)
        varRRLASTUPDATE = N2Str2Null(rsNewRR_HD!lastupdate)
        varRRREMARKS = N2Str2Null(rsNewRR_HD!remarks)

        MoveSql = "INSERT into PMIS_Rec_Hist " & _
                  "(RRNO,RRDATE,CANCDATE,PONO,PODATE,RECVD_CODE,RECVD_FROM,ADDRESS,DRNO,INVNO,CLASSCODE,TERMS,TTLRRAMT,DS1,DS_DESC1,DS_AMT1,NETRRAMT,STATUS,LISTED,USERCODE,LASTUPDATE,REMARKS)" & _
                " values (" & varRRRRNO & ", " & varRRRRDATE & ", " & varRRCANCDATE & ", " & varRRPONO & ", " & varRRPODATE & ", " & varRRRECVD_CODE & ", " & varRRRECVD_FROM & ", " & varRRADDRESS & ", " & varRRDRNO & ", " & varRRINVNO & ", " & varRRCLASSCODE & ", " & varRRTERMS & ", " & varRRTTLRRAMT & ", " & varRRDS1 & ", " & varRRDS_DESC1 & ", " & varRRDS_AMT1 & ", " & varRRNETRRAMT & ", " & varRRSTATUS & ", " & varRRLISTED & ", " & varRRUSERCODE & ", " & varRRLASTUPDATE & ", " & varRRREMARKS & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete from PMIS_RR_Hd where id = " & varRRID
    End If
    Set rsNewRR_HD = Nothing
End Sub

Sub MovePOhd(aydi As Long)
    Dim MoveSql                        As String
    Dim I                              As Integer

    Dim varPOID                        As Long
    Dim varPOPONO, varPOPODATE, varPOPPNO, varPOORDERTYPE As String
    Dim varPODON, varPOSUPCODE, varPOSUPNAME, varPOSUP_ADDRS As String
    Dim varPODEALERCODE, varPOSHIPTO, varPOSHP_ADDRS As String
    Dim varPOPO_AMOUNT, varPODS1       As Double
    Dim varPODS_DESC1                  As String
    Dim varPODS_AMT1, varPONETPOAMT    As Double
    Dim varPOSTATUS, varPOLISTED, varPOUSERCODE As String
    Dim varPOLASTUPDATE, varPOREMARKS  As String

    Dim rsNewPO_HD                     As ADODB.Recordset
    Set rsNewPO_HD = New ADODB.Recordset
    rsNewPO_HD.Open "select * from PMIS_PO_Hd where id =" & aydi, gconDMIS
    If Not rsNewPO_HD.EOF And Not rsNewPO_HD.BOF Then
        DoEvents
        varPOID = rsNewPO_HD!ID
        varPOPONO = N2Str2Null(rsNewPO_HD!PONO)
        varPOPODATE = N2Str2Null(rsNewPO_HD!podate)
        varPOPPNO = N2Str2Null(rsNewPO_HD!ppno)
        varPOORDERTYPE = N2Str2Null(rsNewPO_HD!ORDERTYPE)
        varPODON = N2Str2Null(rsNewPO_HD!DON)
        varPOSUPCODE = N2Str2Null(rsNewPO_HD!SupCode)
        varPOSUPNAME = N2Str2Null(rsNewPO_HD!supname)
        varPOSUP_ADDRS = N2Str2Null(rsNewPO_HD!sup_addrs)
        varPODEALERCODE = N2Str2Null(rsNewPO_HD!dealercode)
        varPOSHIPTO = N2Str2Null(rsNewPO_HD!Shipto)
        varPOSHP_ADDRS = N2Str2Null(rsNewPO_HD!shp_addrs)
        varPOPO_AMOUNT = N2Str2Zero(rsNewPO_HD!po_amount)
        varPODS1 = N2Str2IntZero(rsNewPO_HD!ds1)
        varPODS_DESC1 = N2Str2Null(rsNewPO_HD!ds_desc1)
        varPODS_AMT1 = N2Str2Zero(rsNewPO_HD!ds_amt1)
        varPONETPOAMT = N2Str2Zero(rsNewPO_HD!netpoamt)
        varPOSTATUS = N2Str2Null(rsNewPO_HD!Status)
        varPOLISTED = N2Str2Null(rsNewPO_HD!listed)
        varPOUSERCODE = N2Str2Null(rsNewPO_HD!usercode)
        varPOLASTUPDATE = N2Str2Null(rsNewPO_HD!lastupdate)
        varPOREMARKS = N2Str2Null(rsNewPO_HD!remarks)

        MoveSql = "INSERT into PMIS_PO_Hist " & _
                  "(PONO,PODATE,PPNO,ORDERTYPE,DON,SUPCODE,SUPNAME,SUP_ADDRS,DEALERCODE,SHIPTO,SHP_ADDRS,PO_AMOUNT,DS1,DS_DESC1,DS_AMT1,NETPOAMT,STATUS,LISTED,USERCODE,LASTUPDATE,REMARKS)" & _
                " values (" & varPOPONO & ", " & varPOPODATE & ", " & varPOPPNO & ", " & varPOORDERTYPE & ", " & varPODON & ", " & varPOSUPCODE & ", " & varPOSUPNAME & ", " & varPOSUP_ADDRS & ", " & varPODEALERCODE & ", " & varPOSHIPTO & ", " & varPOSHP_ADDRS & ", " & varPOPO_AMOUNT & ", " & varPODS1 & ", " & varPODS_DESC1 & ", " & varPODS_AMT1 & ", " & varPONETPOAMT & ", " & varPOSTATUS & ", " & varPOLISTED & ", " & varPOUSERCODE & ", " & varPOLASTUPDATE & ", " & varPOREMARKS & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete from PMIS_PO_Hd where id = " & varPOID
    End If
    Set rsNewPO_HD = Nothing
End Sub

Sub MovePPhd(aydi As Long)
    Dim MoveSql                        As String
    Dim I                              As Integer

    Dim varPPID                        As Long
    Dim varPPPPNO, varPPPPDATE, varPPORDERTYPE, varPPDON As String
    Dim varPPSUPCODE, varPPSUPNAME, varPPSUP_ADDRS As String
    Dim varPPDEALERCODE, varPPSHIPTO, varPPSHP_ADDRS As String
    Dim varPPPP_AMOUNT, varPPDS1       As Double
    Dim varPPDS_DESC1                  As String
    Dim varPPDS_AMT1, varPPNETPPAMT    As Double
    Dim varPPSTATUS, varPPLISTED, varPPUSERCODE, varPPLASTUPDATE As String
    Dim varPPREMARKS                   As String

    Dim rsNewPP_HD                     As ADODB.Recordset
    Set rsNewPP_HD = New ADODB.Recordset
    rsNewPP_HD.Open "select * from PMIS_PP_Hd where id =" & aydi, gconDMIS
    If Not rsNewPP_HD.EOF And Not rsNewPP_HD.BOF Then
        DoEvents
        varPPID = rsNewPP_HD!ID
        varPPPPNO = N2Str2Null(rsNewPP_HD!ppno)
        varPPPPDATE = N2Str2Null(rsNewPP_HD!ppdate)
        varPPORDERTYPE = N2Str2Null(rsNewPP_HD!ORDERTYPE)
        varPPDON = N2Str2Null(rsNewPP_HD!DON)
        varPPSUPCODE = N2Str2Null(rsNewPP_HD!SupCode)
        varPPSUPNAME = N2Str2Null(rsNewPP_HD!supname)
        varPPSUP_ADDRS = N2Str2Null(rsNewPP_HD!sup_addrs)
        varPPDEALERCODE = N2Str2Null(rsNewPP_HD!dealercode)
        varPPSHIPTO = N2Str2Null(rsNewPP_HD!Shipto)
        varPPSHP_ADDRS = N2Str2Null(rsNewPP_HD!shp_addrs)
        varPPPP_AMOUNT = N2Str2Zero(rsNewPP_HD!pp_amount)
        varPPDS1 = N2Str2IntZero(rsNewPP_HD!ds1)
        varPPDS_DESC1 = N2Str2Null(rsNewPP_HD!ds_desc1)
        varPPDS_AMT1 = N2Str2Zero(rsNewPP_HD!ds_amt1)
        varPPNETPPAMT = N2Str2Zero(rsNewPP_HD!netppamt)
        varPPSTATUS = N2Str2Null(rsNewPP_HD!Status)
        varPPLISTED = N2Str2Null(rsNewPP_HD!listed)
        varPPUSERCODE = N2Str2Null(rsNewPP_HD!usercode)
        varPPLASTUPDATE = N2Str2Null(rsNewPP_HD!lastupdate)
        varPPREMARKS = N2Str2Null(rsNewPP_HD!remarks)

        MoveSql = "INSERT INTO PP_HIST " & _
                  "(PPNO,PPDATE,ORDERTYPE,DON,SUPCODE,SUPNAME,SUP_ADDRS,DEALERCODE,SHIPTO,SHP_ADDRS,PP_AMOUNT,DS1,DS_DESC1,DS_AMT1,NETPPAMT,STATUS,LISTED,USERCODE,LASTUPDATE,REMARKS)" & _
                " values (" & varPPPPNO & ", " & varPPPPDATE & ", " & varPPORDERTYPE & ", " & varPPDON & ", " & varPPSUPCODE & ", " & varPPSUPNAME & ", " & varPPSUP_ADDRS & ", " & varPPDEALERCODE & ", " & varPPSHIPTO & ", " & varPPSHP_ADDRS & ", " & varPPPP_AMOUNT & ", " & varPPDS1 & ", " & varPPDS_DESC1 & ", " & varPPDS_AMT1 & ", " & varPPNETPPAMT & ", " & varPPSTATUS & ", " & varPPLISTED & ", " & varPPUSERCODE & ", " & varPPLASTUPDATE & ", " & varPPREMARKS & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete from PMIS_PP_Hd where id = " & varPPID
    End If
    Set rsNewPP_HD = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISBatchPosting = Nothing
    UnloadForm Me
End Sub

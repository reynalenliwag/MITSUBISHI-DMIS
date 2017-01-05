VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmCSMSMatBatchPosting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Batch Posting"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MatBatchPosting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1740
   ScaleWidth      =   5835
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
      MouseIcon       =   "MatBatchPosting.frx":27A2
      MousePointer    =   99  'Custom
      Picture         =   "MatBatchPosting.frx":28F4
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
      Left            =   4980
      MouseIcon       =   "MatBatchPosting.frx":2C19
      MousePointer    =   99  'Custom
      Picture         =   "MatBatchPosting.frx":2D6B
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   765
      Width           =   705
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   60
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
            ToolTipText     =   "Process progress"
            Top             =   -30
            Width           =   3525
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   0
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
            MICON           =   "MatBatchPosting.frx":30D1
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
         Picture         =   "MatBatchPosting.frx":30ED
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "MatBatchPosting.frx":3109
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
Attribute VB_Name = "frmCSMSMatBatchPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTDAYTRAN, rsMatMas, rsShipping As ADODB.Recordset
Attribute rsMatMas.VB_VarUserMemId = 1073938432
Attribute rsShipping.VB_VarUserMemId = 1073938432
Dim rsMATREC, rsMATISS, rsMATISS_HIST As ADODB.Recordset
Attribute rsMATREC.VB_VarUserMemId = 1073938435
Attribute rsMATISS.VB_VarUserMemId = 1073938435
Attribute rsMATISS_HIST.VB_VarUserMemId = 1073938435
Dim rsMATREC_HIST        As ADODB.Recordset
Attribute rsMATREC_HIST.VB_VarUserMemId = 1073938438
Dim rsDAYTRAN, rsNOHeader As ADODB.Recordset
Attribute rsDAYTRAN.VB_VarUserMemId = 1073938439
Attribute rsNOHeader.VB_VarUserMemId = 1073938439
Dim rsNODetail, rsNO_Mstr, rsSupplier As ADODB.Recordset
Attribute rsNODetail.VB_VarUserMemId = 1073938441
Attribute rsNO_Mstr.VB_VarUserMemId = 1073938441
Attribute rsSupplier.VB_VarUserMemId = 1073938441

Dim vSupplier, vVatAmt, AddSql, upsql As String
Attribute vSupplier.VB_VarUserMemId = 1073938444
Attribute vVatAmt.VB_VarUserMemId = 1073938444
Attribute AddSql.VB_VarUserMemId = 1073938444
Attribute upsql.VB_VarUserMemId = 1073938444
Dim vTDTranno, vTDMatOrd, vTDTranType As String
Attribute vTDTranno.VB_VarUserMemId = 1073938448
Attribute vTDMatOrd.VB_VarUserMemId = 1073938448
Attribute vTDTranType.VB_VarUserMemId = 1073938448
Dim vTDInOut, vTDStatus  As String
Attribute vTDInOut.VB_VarUserMemId = 1073938451
Attribute vTDStatus.VB_VarUserMemId = 1073938451
Dim vTotTranCost, vCOST  As Double
Attribute vTotTranCost.VB_VarUserMemId = 1073938453
Attribute vCOST.VB_VarUserMemId = 1073938453
Dim vTDRecNo, vMatRecNo  As Long
Attribute vTDRecNo.VB_VarUserMemId = 1073938455
Attribute vMatRecNo.VB_VarUserMemId = 1073938455
Dim vMatOnhand, vMatTrecqty, vPMTissqty As Integer
Attribute vMatOnhand.VB_VarUserMemId = 1073938457
Attribute vMatTrecqty.VB_VarUserMemId = 1073938457
Attribute vPMTissqty.VB_VarUserMemId = 1073938457
Dim vMatLast_Recd, vTDTranDate As String
Attribute vMatLast_Recd.VB_VarUserMemId = 1073938460
Attribute vTDTranDate.VB_VarUserMemId = 1073938460
Dim vMatReceipts, vMatIssuances, vTDTranQTY As Integer
Attribute vMatReceipts.VB_VarUserMemId = 1073938462
Attribute vMatIssuances.VB_VarUserMemId = 1073938462
Attribute vTDTranQTY.VB_VarUserMemId = 1073938462
Dim vTDNetPrice, vTDNetCost, vTDTranucost As Double
Attribute vTDNetPrice.VB_VarUserMemId = 1073938465
Attribute vTDNetCost.VB_VarUserMemId = 1073938465
Attribute vTDTranucost.VB_VarUserMemId = 1073938465
Dim vORDTotPrice, vTDTranuprice As Double
Attribute vORDTotPrice.VB_VarUserMemId = 1073938468
Attribute vTDTranuprice.VB_VarUserMemId = 1073938468
Dim vShCurrMonth         As Integer
Attribute vShCurrMonth.VB_VarUserMemId = 1073938470
Dim vShRecNo             As Long
Attribute vShRecNo.VB_VarUserMemId = 1073938471
Dim vNetPrice, vNetCost  As Double
Attribute vNetPrice.VB_VarUserMemId = 1073938472
Attribute vNetCost.VB_VarUserMemId = 1073938472
Dim vMatIssRecNo, vMatRecRecNo As Long
Attribute vMatIssRecNo.VB_VarUserMemId = 1073938474
Attribute vMatRecRecNo.VB_VarUserMemId = 1073938474

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
    Dim i                As Integer
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "Select id,in_out,trantype,tranno,MatCde,status,tranqty,netcost,tranucost,trandate,tranuprice from PMIS_TdayTran where status <> 'C' order by id asc", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        rsTDAYTRAN.MoveFirst
        i = 0
        Screen.MousePointer = 11
        MsgSpeech "Posting Transactions from Daily Transactions File..."
        Me.Caption = "Posting Transactions from PMIS_TdayTran File..."
        DoEvents
        Do While Not rsTDAYTRAN.EOF
            vTDRecNo = rsTDAYTRAN!ID
            vTDInOut = Null2String(rsTDAYTRAN!IN_OUT)
            vTDTranType = Null2String(rsTDAYTRAN!TRANTYPE)
            vTDTranno = Null2String(rsTDAYTRAN!Tranno)
            vTDMatOrd = Null2String(rsTDAYTRAN!MATCDE)
            vTDStatus = Null2String(rsTDAYTRAN!Status)
            vTDTranQTY = N2Str2IntZero(rsTDAYTRAN!tranqty)
            vTDNetCost = N2Str2Zero(rsTDAYTRAN!netcost)
            vTDTranucost = N2Str2Zero(rsTDAYTRAN!TRANUCOST)
            vTDTranDate = Null2Date(rsTDAYTRAN!trandate)
            vTotTranCost = vTDTranucost * vTDTranQTY
            vTDTranuprice = N2Str2Zero(rsTDAYTRAN!TRANUPRICE)
            labProcessing.Caption = "Processing: " & vTDTranType & " #" & vTDTranno
            DoEvents
            Set rsMatMas = New ADODB.Recordset
            rsMatMas.Open "Select id,onhand,trecqty,last_recd,receipts,tissqty,issuances,lastm_MAC,COST from MatMas where MatCde = '" & vTDMatOrd & "'", gconDMIS
            If Not rsMatMas.EOF And Not rsMatMas.BOF Then
                If vTDTranType <> "ADJ" And vTDTranType <> "PO" And (vTDInOut = "I" Or vTDInOut = "O") And vTDTranQTY <> 0 And vTDStatus <> "C" Then
                    If vTDTranType = "RR" Then
                        Set rsMATREC = New ADODB.Recordset
                        rsMATREC.Open "Select recvd_code,ds1,status,classcode,rrno from MATREC where rrno = '" & Format(rsTDAYTRAN!Tranno, "000000") & "'", gconDMIS
                        If Not rsMATREC.EOF And Not rsMATREC.BOF Then
                            vSupplier = Null2String(rsMATREC!recvd_code)
                            vVatAmt = N2Str2IntZero(rsMATREC!ds1)
                            If rsMATREC!classcode = "PCG" Or rsMATREC!classcode = "PCS" Then
                                If vSupplier <> vPAMCOR And vVatAmt <= 0 Then
                                    vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(N2Str2Zero(rsMATREC!ds1))
                                End If
                            End If
                            vMatRecNo = rsMatMas!ID
                            vMatOnhand = N2Str2IntZero(rsMatMas!ONHAND)
                            vMatTrecqty = N2Str2IntZero(rsMatMas!trecqty)
                            vMatLast_Recd = Null2Date(rsMatMas!last_recd)
                            vMatReceipts = N2Str2IntZero(rsMatMas!receipts)
                            vCOST = N2Str2Zero(rsMatMas!COST)
                            gconDMIS.Execute "update MatMas set " & _
                                             "trecqty = " & vMatTrecqty - vTDTranQTY & ", " & _
                                             "last_recd = " & N2Str2Null(vTDTranDate) & _
                                           " where id =" & vMatRecNo
                            gconDMIS.Execute "update PMIS_TdayTran set MAC = " & vCOST & ", status = 'P', netcost = " & vTotTranCost & " where id = " & vTDRecNo
                        Else
                            gconDMIS.Execute "insert into PMIS_NoHeader " & _
                                             "(trantype,tranno,recno,stat_h)" & _
                                           " values ('" & "RR" & "', '" & vTDTranno & "', " & vTDRecNo & ", '" & vTDStatus & "')"
                            MsgSpeechBox "Error in " & vTDTranType & "-" & vTDTranno & " does not have header File"
                        End If
                    End If
                    If vTDInOut = "O" Then
                        Set rsMATISS = New ADODB.Recordset
                        rsMATISS.Open "Select trantype,tranno from MATISS where trantype = '" & vTDTranType & "' and tranno = '" & Format(rsTDAYTRAN!Tranno, "000000") & "'", gconDMIS
                        If Not rsMATISS.EOF And Not rsMATISS.BOF Then
                            If vTDTranType = "CHG" Or vTDTranType = "CSH" Or vTDTranType = "MRIS" Then
                                vORDTotPrice = (vTDTranuprice * vTDTranQTY) / ConvertToBIRDecimalFormat(VAT_RATE)
                            Else
                                vORDTotPrice = (vTDTranuprice * vTDTranQTY)
                            End If
                            vMatRecNo = rsMatMas!ID
                            vPMTissqty = N2Str2IntZero(rsMatMas!TISSQTY)
                            vMatIssuances = N2Str2IntZero(rsMatMas!issuances)
                            vCOST = N2Str2Zero(rsMatMas!COST)
                            vTotTranCost = vTDTranucost * vTDTranQTY
                            gconDMIS.Execute "update MatMas set " & _
                                             "tissqty = " & vPMTissqty - vTDTranQTY & _
                                           " where id =" & vMatRecNo
                            gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & vCOST & ", netcost = " & vTotTranCost & ", netprice = " & vORDTotPrice & ", status = 'P' where id = " & vTDRecNo
                            Set rsShipping = New ADODB.Recordset
                            rsShipping.Open "select * from PMIS_Shipping where MatCde = '" & vTDMatOrd & "'", gconDMIS
                            If Not rsShipping.EOF And Not rsShipping.BOF Then
                                vShRecNo = rsShipping!ID
                                vShCurrMonth = N2Str2IntZero(rsShipping!curr_month)
                                gconDMIS.Execute "update PMIS_Shipping set curr_month = " & vShCurrMonth + vTDTranQTY & ", " & _
                                                 "freq_curr = 1 where id = " & vShRecNo
                            Else
                                gconDMIS.Execute "insert into PMIS_Shipping (MatCde,curr_month,freq_curr)" & _
                                               " values ('" & vTDMatOrd & "', " & vTDTranQTY & ", 1)"
                            End If
                        Else
                            gconDMIS.Execute "insert into PMIS_NoHeader " & _
                                             "(trantype,tranno,recno,stat_h)" & _
                                           " values ('" & vTDTranType & "', '" & vTDTranno & "', " & vTDRecNo & ", '" & vTDStatus & "')"
                            MsgSpeechBox "Error in " & vTDTranType & "-" & vTDTranno & " does not have header File"
                        End If
                    End If
                End If

                If vTDTranType = "ADJ" And vTDInOut = "I" And vTDTranQTY <> 0 And vTDStatus <> "C" Then
                    gconDMIS.Execute "update PMIS_TdayTran set " & _
                                     "tranucost = " & N2Str2Zero(rsMatMas!COST) & "," & _
                                     "netcost = " & N2Str2Zero(rsMatMas!COST) * vTDTranQTY & _
                                   " where id = " & vTDRecNo
                    vTotTranCost = N2Str2Zero(rsMatMas!COST) * vTDTranQTY

                    vCOST = N2Str2Zero(rsMatMas!COST)
                    vMatRecNo = rsMatMas!ID
                    vMatOnhand = N2Str2IntZero(rsMatMas!ONHAND)
                    vMatTrecqty = N2Str2IntZero(rsMatMas!trecqty)

                    gconDMIS.Execute "update MatMas set " & _
                                     "trecqty = " & vMatTrecqty - vTDTranQTY & ", " & _
                                     "last_recd = " & N2Str2Null(vTDTranDate) & _
                                   " where id =" & vMatRecNo
                    gconDMIS.Execute "update PMIS_TdayTran set MAC = " & vCOST & ", status = 'P', netcost = " & vTotTranCost & " where id = " & vTDRecNo
                End If

                If vTDTranType = "ADJ" And vTDInOut = "O" And vTDTranQTY <> 0 And vTDStatus <> "C" Then
                    gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & N2Str2Zero(rsMatMas!COST) & _
                                   " where id = " & vTDRecNo
                    vTotTranCost = N2Str2Zero(rsMatMas!COST) * vTDTranQTY

                    vMatRecNo = rsMatMas!ID
                    vCOST = N2Str2Zero(rsMatMas!COST)
                    vORDTotPrice = (vCOST * vTDTranQTY)
                    vPMTissqty = N2Str2IntZero(rsMatMas!TISSQTY)
                    vMatIssuances = N2Str2IntZero(rsMatMas!issuances)

                    gconDMIS.Execute "update MatMas set " & _
                                     "tissqty = " & vPMTissqty - vTDTranQTY & ", " & _
                                     "issuances = " & vMatIssuances - vTDTranQTY & _
                                   " where id =" & vMatRecNo
                    gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & vCOST & ", netcost = " & vTotTranCost & ", netprice = " & vORDTotPrice & ", status = 'P' where id = " & vTDRecNo
                End If

            Else
                gconDMIS.Execute "insert into PMIS_No_Mstr " & _
                                 "(trantype,tranno,recno)" & _
                               " values ('" & vTDInOut & "', '" & vTDTranno & "', " & vTDRecNo & ")"
                MsgSpeechBox "Error in " & vTDTranType & "-" & vTDTranno & " Part Number: " & vTDMatOrd & " is not in Master File"
            End If
            i = i + 1
            progCPB.Value = (i / rsTDAYTRAN.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTDAYTRAN.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsMATISS = New ADODB.Recordset
    rsMATISS.Open "select id,trantype,tranno,status from MATISS order by trantype,tranno asc", gconDMIS
    If Not rsMATISS.EOF And Not rsMATISS.BOF Then
        rsMATISS.MoveFirst
        i = 0
        MsgSpeech "Computing Issuances Netcost and Netprice..."
        Me.Caption = "Computing Order Netcost and Netprice..."
        Screen.MousePointer = 11
        DoEvents
        Do While Not rsMATISS.EOF
            vMatIssRecNo = rsMATISS!ID
            labProcessing.Caption = "Processing: " & Null2String(rsMATISS!TRANTYPE) & " #" & Null2String(rsMATISS!Tranno)
            DoEvents
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select id,trantype,tranno,netprice,netcost,status,itemno from PMIS_TdayTran where trantype = " & N2Str2Null(rsMATISS!TRANTYPE) & " and tranno = " & N2Str2Null(rsMATISS!Tranno) & " order by itemno asc", gconDMIS
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
                    MoveTdaytran (rsTDAYTRAN!ID)
                    rsTDAYTRAN.MoveNext
                Loop
                If Null2String(rsMATISS!Status) <> "C" Then
                    gconDMIS.Execute "update MATISS set netcost = " & vNetCost & ", netinvamt2 = " & vNetPrice & ", status = 'P' where id = " & vMatIssRecNo
                End If
                MoveMatIss (vMatIssRecNo)
            End If
            i = i + 1
            progCPB.Value = (i / rsMATISS.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsMATISS.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsMATREC = New ADODB.Recordset
    rsMATREC.Open "select id,rrno,status from MATREC order by rrno asc", gconDMIS
    If Not rsMATREC.EOF And Not rsMATREC.BOF Then
        rsMATREC.MoveFirst
        i = 0
        Screen.MousePointer = 11
        MsgSpeech "Checking if details of receipts are already posted..."
        Me.Caption = "Checking if details of receipts are already posted..."
        DoEvents
        Do While Not rsMATREC.EOF
            vMatRecRecNo = rsMATREC!ID
            labProcessing.Caption = "Processing: RR #" & Null2String(rsMATREC!rrno)
            DoEvents
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select id,status,trantype,tranno,itemno from PMIS_TdayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC!rrno) & " order by itemno asc", gconDMIS
            If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
                rsTDAYTRAN.MoveFirst
                Do While Not rsTDAYTRAN.EOF
                    vTDRecNo = rsTDAYTRAN!ID
                    vTDStatus = Null2String(rsTDAYTRAN!Status)
                    If vTDStatus <> "C" Then
                        gconDMIS.Execute "update PMIS_TdayTran set status = 'P' where id =" & vTDRecNo
                    End If
                    If Null2String(rsMATREC!Status) <> "C" Then
                        gconDMIS.Execute "update MATREC set status = 'P' where id = " & vMatRecRecNo
                    End If
                    MoveTdaytran (rsTDAYTRAN!ID)
                    rsTDAYTRAN.MoveNext
                Loop
                MoveMatRec (vMatRecRecNo)
            End If
            i = i + 1
            progCPB.Value = (i / rsMATREC.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsMATREC.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select id,status,trantype,tranno,itemno from PMIS_TdayTran where trantype = 'ADJ' order by id asc", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        rsTDAYTRAN.MoveFirst
        labProcessing.Caption = "Processing: ADJ #" & Null2String(rsTDAYTRAN!Tranno)
        DoEvents
        Do While Not rsTDAYTRAN.EOF
            vTDRecNo = rsTDAYTRAN!ID
            vTDStatus = Null2String(rsTDAYTRAN!Status)
            If vTDStatus = "N" Then
                gconDMIS.Execute "update PMIS_TdayTran set status = 'P' where id =" & vTDRecNo
            End If
            MoveTdaytran (rsTDAYTRAN!ID)
            rsTDAYTRAN.MoveNext
        Loop
    End If

    MsgSpeechBox "Posting of Transactions Completed..."
    'UNDER MAINTAINANCE
    'frmMain.mnuMatBatchPosting.Enabled = False

    cmdPost.Enabled = False
    Set rsTDAYTRAN = Nothing
    Set rsMatMas = Nothing
    Set rsShipping = Nothing
    Set rsMATISS = Nothing
    Set rsMATREC = Nothing
End Sub

Sub MoveTdaytran(aydi As Long)
    Dim MoveSql          As String
    Dim i                As Integer

    Dim varTRANID, varTRANDATE, varTRANTYPE, varTRANNO As String
    Dim varITEMNO, varMatCde, varMatDsc As String
    Dim varTRANQTY       As Integer
    Dim varUNIT          As String
    Dim varTRANUCOST, varTRANUPRICE, varNETCOST, varNETPRICE As Double
    Dim varSTATUS, varIN_OUT, varMATCH, varLISTED As String
    Dim varCOST, varTRANINVAMT As Double
    Dim varUSERCODE, varLASTUPDATE, varTREMARKS As String

    Dim rsNewTdaytran    As ADODB.Recordset
    Set rsNewTdaytran = New ADODB.Recordset
    rsNewTdaytran.Open "select * from PMIS_TdayTran where id =" & aydi, gconDMIS
    If Not rsNewTdaytran.EOF And Not rsNewTdaytran.BOF Then
        varTRANID = rsNewTdaytran!ID
        varTRANDATE = N2Str2Null(rsNewTdaytran!trandate)
        varTRANTYPE = N2Str2Null(rsNewTdaytran!TRANTYPE)
        varTRANNO = N2Str2Null(rsNewTdaytran!Tranno)
        varITEMNO = N2Str2Null(rsNewTdaytran!itemno)
        varMatCde = N2Str2Null(rsNewTdaytran!MATCDE)
        varMatDsc = N2Str2Null(rsNewTdaytran!MatDsc)
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
        varCOST = N2Str2Zero(rsNewTdaytran!Mac)
        varTRANINVAMT = N2Str2Zero(rsNewTdaytran!TRANINVAMT)
        varUSERCODE = N2Str2Null(rsNewTdaytran!usercode)
        varLASTUPDATE = N2Str2Null(rsNewTdaytran!lastupdate)
        varTREMARKS = N2Str2Null(rsNewTdaytran!tremarks)

        MoveSql = "INSERT into PMIS_DayTran " & _
                  "(TRANDATE,TRANTYPE,TRANNO,ITEMNO,MatCde,MatDsc,TRANQTY,UNIT,TRANUCOST,TRANUPRICE,NETCOST,NETPRICE,STATUS,IN_OUT,LISTED,MAC,TRANINVAMT,USERCODE,LASTUPDATE,TREMARKS)" & _
                " values (" & varTRANDATE & "," & varTRANTYPE & "," & varTRANNO & "," & varITEMNO & "," & varMatCde & "," & varMatDsc & "," & varTRANQTY & "," & varUNIT & "," & varTRANUCOST & "," & varTRANUPRICE & "," & varNETCOST & "," & varNETPRICE & "," & varSTATUS & "," & varIN_OUT & "," & varLISTED & "," & varCOST & "," & varTRANINVAMT & "," & varUSERCODE & "," & varLASTUPDATE & "," & varTREMARKS & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete from PMIS_TdayTran where id = " & varTRANID
    End If
    Set rsNewTdaytran = Nothing
End Sub

Sub MoveMatIss(aydi As Long)
    Dim MoveSql          As String
    Dim i                As Integer

    Dim varOHID          As Long
    Dim varOHTRANTYPE, varOHTRANNO, varOHTRANDATE As String
    Dim varOHCANCDATE, varOHCUSTCODE, varOHCUSTNAME As String
    Dim varOHCHARGETO, varOHRONO, varOHSALESMAN As String
    Dim varOHSMNAME, varOHTERMS As String
    Dim varOHTTLINVAMT, varOHDS1 As Double
    Dim varOHDS_DESC1    As String
    Dim varOHDS_AMT1, varOHNETINVAMT, varOHNETCOST As Double
    Dim varOHSTATUS      As String
    Dim varOHNETINVAMT2, varOHNETCOST2 As Double
    Dim varOHLISTED, varOHUSERCODE, varOHLASTUPDATE As String
    Dim varOHTOTINVAMT, varOHDISCOUNT, varOHVAT As Double
    Dim varOHNETINVOICE, varOHTOTALCOST As Double
    Dim varOHREMARKS     As String

    Dim rsNewMATISS      As ADODB.Recordset
    Set rsNewMATISS = New ADODB.Recordset
    rsNewMATISS.Open "select * from MATISS where id =" & aydi, gconDMIS
    If Not rsNewMATISS.EOF And Not rsNewMATISS.BOF Then
        DoEvents
        varOHID = rsMATISS!ID
        varOHTRANTYPE = N2Str2Null(rsNewMATISS!TRANTYPE)
        varOHTRANNO = N2Str2Null(rsNewMATISS!Tranno)
        varOHTRANDATE = N2Str2Null(rsNewMATISS!trandate)
        varOHCANCDATE = N2Str2Null(rsNewMATISS!cancdate)
        varOHCUSTCODE = N2Str2Null(rsNewMATISS!custcode)
        varOHCUSTNAME = N2Str2Null(rsNewMATISS!custname)
        varOHCHARGETO = N2Str2Null(rsNewMATISS!chargeto)
        varOHRONO = N2Str2Null(rsNewMATISS!rono)
        varOHSALESMAN = N2Str2Null(rsNewMATISS!salesman)
        varOHSMNAME = N2Str2Null(rsNewMATISS!smname)
        varOHTERMS = N2Str2Null(rsNewMATISS!terms)
        varOHTTLINVAMT = N2Str2Zero(rsNewMATISS!ttlinvamt)
        varOHDS1 = N2Str2IntZero(rsNewMATISS!ds1)
        varOHDS_DESC1 = N2Str2Null(rsNewMATISS!ds_desc1)
        varOHDS_AMT1 = N2Str2Zero(rsNewMATISS!ds_amt1)
        varOHNETINVAMT = N2Str2Zero(rsNewMATISS!netinvamt)
        varOHNETCOST = N2Str2Zero(rsNewMATISS!netcost)
        varOHSTATUS = N2Str2Null(rsNewMATISS!Status)
        varOHNETINVAMT2 = N2Str2Zero(rsNewMATISS!NETINVAMT2)
        varOHNETCOST2 = N2Str2Zero(rsNewMATISS!NETCOST2)
        varOHLISTED = N2Str2Null(rsNewMATISS!listed)
        varOHUSERCODE = N2Str2Null(rsNewMATISS!usercode)
        varOHLASTUPDATE = N2Str2Null(rsNewMATISS!lastupdate)
        varOHTOTINVAMT = N2Str2Zero(rsNewMATISS!TOTINVAMT)
        varOHDISCOUNT = N2Str2Zero(rsNewMATISS!DISCOUNT)
        varOHVAT = N2Str2Zero(rsNewMATISS!Vat)
        varOHNETINVOICE = N2Str2Zero(rsNewMATISS!NETINVOICE)
        varOHTOTALCOST = N2Str2Zero(rsNewMATISS!TotalCost)
        varOHREMARKS = N2Str2Null(rsNewMATISS!remarks)

        MoveSql = "INSERT INTO MATISS_HIST " & _
                  "(TRANTYPE,TRANNO,TRANDATE,CANCDATE,CUSTCODE,CUSTNAME,CHARGETO,RONO,SALESMAN,SMNAME,TERMS,TTLINVAMT,DS1,DS_DESC1,DS_AMT1,NETINVAMT,NETCOST,STATUS,NETINVAMT2,NETCOST2,LISTED,USERCODE,LASTUPDATE,TOTINVAMT,DISCOUNT,VAT,NETINVOICE,TOTALCOST,REMARKS)" & _
                " values (" & varOHTRANTYPE & ", " & varOHTRANNO & ", " & varOHTRANDATE & ", " & varOHCANCDATE & ", " & varOHCUSTCODE & ", " & varOHCUSTNAME & ", " & varOHCHARGETO & ", " & varOHRONO & ", " & varOHSALESMAN & ", " & varOHSMNAME & ", " & varOHTERMS & ", " & varOHTTLINVAMT & ", " & varOHDS1 & ", " & varOHDS_DESC1 & ", " & varOHDS_AMT1 & ", " & varOHNETINVAMT & ", " & varOHNETCOST & ", " & varOHSTATUS & ", " & varOHNETINVAMT2 & ", " & varOHNETCOST2 & ", " & varOHLISTED & ", " & varOHUSERCODE & ", " & varOHLASTUPDATE & ", " & varOHTOTINVAMT & ", " & varOHDISCOUNT & ", " & varOHVAT & ", " & varOHNETINVOICE & ", " & varOHTOTALCOST & ", " & varOHREMARKS & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete from MATISS where id = " & varOHID
    End If
    Set rsNewMATISS = Nothing
End Sub

Sub MoveMatRec(aydi As Long)
    Dim MoveSql          As String
    Dim i                As Integer

    Dim varRRID          As Long
    Dim varRRRRNO, varRRRRDATE, varRRCANCDATE, varRRPONO As String
    Dim varRRPODATE, varRRRECVD_CODE, varRRRECVD_FROM As String
    Dim varRRADDRESS, varRRDRNO, varRRINVNO As String
    Dim varRRCLASSCODE, varRRTERMS As String
    Dim varRRTTLRRAMT, varRRDS1 As Double
    Dim varRRDS_DESC1    As String
    Dim varRRDS_AMT1, varRRNETRRAMT As Double
    Dim varRRSTATUS, varRRLISTED, varRRUSERCODE As String
    Dim varRRLASTUPDATE, varRRREMARKS As String

    Dim rsNewMATREC      As ADODB.Recordset
    Set rsNewMATREC = New ADODB.Recordset
    rsNewMATREC.Open "select * from MATREC where id =" & aydi, gconDMIS
    If Not rsNewMATREC.EOF And Not rsNewMATREC.BOF Then
        DoEvents
        varRRID = rsMATREC!ID
        varRRRRNO = N2Str2Null(rsNewMATREC!rrno)
        varRRRRDATE = N2Str2Null(rsNewMATREC!rrdate)
        varRRCANCDATE = N2Str2Null(rsNewMATREC!cancdate)
        varRRPONO = N2Str2Null(rsNewMATREC!pono)
        varRRPODATE = N2Str2Null(rsNewMATREC!podate)
        varRRRECVD_CODE = N2Str2Null(rsNewMATREC!recvd_code)
        varRRRECVD_FROM = N2Str2Null(rsNewMATREC!recvd_from)
        varRRADDRESS = N2Str2Null(rsNewMATREC!Address)
        varRRDRNO = N2Str2Null(rsNewMATREC!drno)
        varRRINVNO = N2Str2Null(rsNewMATREC!invno)
        varRRCLASSCODE = N2Str2Null(rsNewMATREC!classcode)
        varRRTERMS = N2Str2Null(rsNewMATREC!terms)
        varRRTTLRRAMT = N2Str2Zero(rsNewMATREC!ttlrramt)
        varRRDS1 = N2Str2IntZero(rsNewMATREC!ds1)
        varRRDS_DESC1 = N2Str2Null(rsNewMATREC!ds_desc1)
        varRRDS_AMT1 = N2Str2Zero(rsNewMATREC!ds_amt1)
        varRRNETRRAMT = N2Str2Zero(rsNewMATREC!netrramt)
        varRRSTATUS = N2Str2Null(rsNewMATREC!Status)
        varRRLISTED = N2Str2Null(rsNewMATREC!listed)
        varRRUSERCODE = N2Str2Null(rsNewMATREC!usercode)
        varRRLASTUPDATE = N2Str2Null(rsNewMATREC!lastupdate)
        varRRREMARKS = N2Str2Null(rsNewMATREC!remarks)

        MoveSql = "INSERT INTO MATREC_HIST " & _
                  "(RRNO,RRDATE,CANCDATE,PONO,PODATE,RECVD_CODE,RECVD_FROM,ADDRESS,DRNO,INVNO,CLASSCODE,TERMS,TTLRRAMT,DS1,DS_DESC1,DS_AMT1,NETRRAMT,STATUS,LISTED,USERCODE,LASTUPDATE,REMARKS)" & _
                " values (" & varRRRRNO & ", " & varRRRRDATE & ", " & varRRCANCDATE & ", " & varRRPONO & ", " & varRRPODATE & ", " & varRRRECVD_CODE & ", " & varRRRECVD_FROM & ", " & varRRADDRESS & ", " & varRRDRNO & ", " & varRRINVNO & ", " & varRRCLASSCODE & ", " & varRRTERMS & ", " & varRRTTLRRAMT & ", " & varRRDS1 & ", " & varRRDS_DESC1 & ", " & varRRDS_AMT1 & ", " & varRRNETRRAMT & ", " & varRRSTATUS & ", " & varRRLISTED & ", " & varRRUSERCODE & ", " & varRRLASTUPDATE & ", " & varRRREMARKS & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete from MATREC where id = " & varRRID
    End If
    Set rsNewMATREC = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISBatchPosting = Nothing
    UnloadForm Me
End Sub

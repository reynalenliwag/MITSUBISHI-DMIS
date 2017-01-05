VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmOSMSProcessBatchPosting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplies Batch Posting"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H8000000F&
   Icon            =   "BatchPosting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1590
   ScaleWidth      =   5820
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
      MouseIcon       =   "BatchPosting.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "BatchPosting.frx":0594
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   660
      Width           =   735
   End
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
      Left            =   4260
      MaskColor       =   &H0000FFFF&
      MouseIcon       =   "BatchPosting.frx":08FA
      MousePointer    =   99  'Custom
      Picture         =   "BatchPosting.frx":0A4C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Press F11 for Posting By Range"
      Top             =   660
      Width           =   735
   End
   Begin VB.PictureBox picCPB 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   0
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
            MICON           =   "BatchPosting.frx":0D71
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
         Picture         =   "BatchPosting.frx":0D8D
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "BatchPosting.frx":0DA9
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
Attribute VB_Name = "frmOSMSProcessBatchPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsISSUANCE_DETAILS As ADODB.Recordset
Dim rsrrDETAILS As ADODB.Recordset
Dim rsSupply As ADODB.Recordset
Dim rsShipping As ADODB.Recordset
Dim rsrrHEADER As ADODB.Recordset
Dim rsIssuance_Header As ADODB.Recordset
Attribute rsIssuance_Header.VB_VarUserMemId = 1073938436

Dim vTDTRANS_NO As String
Attribute vTDTRANS_NO.VB_VarUserMemId = 1073938449
Dim vTDSupplyCode As String
Dim vTDStatus As String
Dim vTotTranCost As Double
Attribute vTotTranCost.VB_VarUserMemId = 1073938454
Dim vCOST As Double
Dim vTDRecNo As Long
Attribute vTDRecNo.VB_VarUserMemId = 1073938456
Dim vSUPPLYID As Long
Dim vSupplyOnHand As Integer
Attribute vSupplyOnHand.VB_VarUserMemId = 1073938458
Dim vSupplyTrecqty As Integer
Dim vPMTissqty As Integer
Dim vMatlastrrdate As String
Attribute vMatlastrrdate.VB_VarUserMemId = 1073938461
Dim vTDTranDate As String
Dim vRRHEADEReceipts As Integer
Attribute vRRHEADEReceipts.VB_VarUserMemId = 1073938463
Dim vISSUANCE_HEADERuances As Integer
Dim vTDID_QUANTITY As Integer
Dim vTDNetCost As Double
Dim vTDTranucost As Double
Dim vORDTotPrice As Double
Attribute vORDTotPrice.VB_VarUserMemId = 1073938469
Dim vTDTranuprice As Double
Dim vShCurrMonth As Integer
Attribute vShCurrMonth.VB_VarUserMemId = 1073938471
Dim vShRecNo As Long
Attribute vShRecNo.VB_VarUserMemId = 1073938472
Dim vNetPrice As Double
Attribute vNetPrice.VB_VarUserMemId = 1073938473
Dim vNetCost As Double
Dim vISSUANCE_HEADERRecNo As Long
Attribute vISSUANCE_HEADERRecNo.VB_VarUserMemId = 1073938475
Dim vRRHEADERRecNo As Long

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
    Set rsISSUANCE_DETAILS = New ADODB.Recordset

    rsISSUANCE_DETAILS.Open "Select id,ID_ITEM_NO,TRANS_NO,Supply_Code,status,ID_QUANTITY,cost from OSMS_ISSUANCE_DETAILS where status <> 'C' order by id asc", gconDMIS

    If Not rsISSUANCE_DETAILS.EOF And Not rsISSUANCE_DETAILS.BOF Then
        rsISSUANCE_DETAILS.MoveFirst
        i = 0
        Screen.MousePointer = 11
        MsgSpeech "Posting Transactions from Issuance Daily Transactions File..."
        DoEvents
        Do While Not rsISSUANCE_DETAILS.EOF
            vTDRecNo = rsISSUANCE_DETAILS!ID
            vTDTRANS_NO = Null2String(rsISSUANCE_DETAILS!Trans_No)
            vTDSupplyCode = Null2String(rsISSUANCE_DETAILS!Supply_Code)
            vTDStatus = Null2String(rsISSUANCE_DETAILS!Status)
            vTDID_QUANTITY = N2Str2IntZero(rsISSUANCE_DETAILS!ID_Quantity)
            vTDTranucost = N2Str2Zero(rsISSUANCE_DETAILS!Cost)
            vTotTranCost = vTDTranucost * vTDID_QUANTITY
            labProcessing.Caption = "Processing: Issuance #" & vTDTRANS_NO
            DoEvents
            Set rsSupply = New ADODB.Recordset

            rsSupply.Open "Select id,onhand,trecqty,lastrrdate,receipts,tissqty,issuances,lastm_cost,COST from OSMS_SUPPLY where Supply_Code = '" & vTDSupplyCode & "'", gconDMIS

            If Not rsSupply.EOF And Not rsSupply.BOF Then
                Set rsIssuance_Header = New ADODB.Recordset
                rsIssuance_Header.Open "Select TRANS_NO FROM OSMS_ISSUANCE_HEADER where TRANS_NO = '" & Format(rsISSUANCE_DETAILS!Trans_No, "000000") & "'", gconDMIS
                If Not rsIssuance_Header.EOF And Not rsIssuance_Header.BOF Then
                    vORDTotPrice = (vTDTranuprice * vTDID_QUANTITY)
                    vSUPPLYID = rsSupply!ID
                    vPMTissqty = N2Str2IntZero(rsSupply!tissqty)
                    vISSUANCE_HEADERuances = N2Str2IntZero(rsSupply!issuances)
                    vCOST = N2Str2Zero(rsSupply!Cost)
                    vTotTranCost = vTDTranucost * vTDID_QUANTITY
                    gconDMIS.Execute "update OSMS_SUPPLY set " & _
                                     "tissqty = " & vPMTissqty - vTDID_QUANTITY & _
                                   " where id =" & vSUPPLYID
                    gconDMIS.Execute "update OSMS_ISSUANCE_DETAILS set status = 'P' where id = " & vTDRecNo
                    Set rsShipping = New ADODB.Recordset
                    rsShipping.Open "select * FROM OSMS_SHIPPING  where Supply_Code = '" & vTDSupplyCode & "'", gconDMIS
                    If Not rsShipping.EOF And Not rsShipping.BOF Then
                        vShRecNo = rsShipping!ID
                        vShCurrMonth = N2Str2IntZero(rsShipping!curr_month)
                        gconDMIS.Execute "update OSMS_shipping set curr_month = " & vShCurrMonth + vTDID_QUANTITY & ", " & _
                                         "freq_curr = 1 where id = " & vShRecNo
                    Else
                        gconDMIS.Execute "insert into osms_shipping (Supply_Code,curr_month,freq_curr)" & _
                                       " values ('" & vTDSupplyCode & "', " & vTDID_QUANTITY & ", 1)"
                    End If
                Else
                    If vTDTRANS_NO <> "000000" And Null2String(rsISSUANCE_DETAILS!id_item_no) <> "0000" Then
                        gconDMIS.Execute "insert into osms_noheader " & _
                                         "(trantype,TRANNO,recno,stat_h)" & _
                                       " values ('ISS', '" & vTDTRANS_NO & "', " & vTDRecNo & ", '" & vTDStatus & "')"
                    End If
                End If
            Else
                gconDMIS.Execute "insert into osms_no_mstr " & _
                                 "(trantype,TRANS_NO,recno)" & _
                               " values ('ISS', '" & vTDTRANS_NO & "', " & vTDRecNo & ")"
            End If
            i = i + 1
            progCPB.Value = (i / rsISSUANCE_DETAILS.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsISSUANCE_DETAILS.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsrrDETAILS = New ADODB.Recordset
    rsrrDETAILS.Open "Select id,item_no,RRNumber,Supply_Code,status,rrQuantity,cost from OSMS_RRDETAILS  where status <> 'C' order by id asc", gconDMIS
    If Not rsrrDETAILS.EOF And Not rsrrDETAILS.BOF Then
        rsrrDETAILS.MoveFirst
        i = 0
        Screen.MousePointer = 11
        MsgSpeech "Posting Transactions from Receipts Daily Transactions File..."
        DoEvents
        Do While Not rsrrDETAILS.EOF
            vTDRecNo = rsrrDETAILS!ID
            vTDTRANS_NO = Null2String(rsrrDETAILS!rrnumber)
            vTDSupplyCode = Null2String(rsrrDETAILS!Supply_Code)
            vTDStatus = Null2String(rsrrDETAILS!Status)
            vTDID_QUANTITY = N2Str2IntZero(rsrrDETAILS!rrQUANTITY)
            vTDTranucost = N2Str2Zero(rsrrDETAILS!Cost)
            vTotTranCost = vTDTranucost * vTDID_QUANTITY
            labProcessing.Caption = "Processing: Receipts #" & vTDTRANS_NO
            DoEvents
            Set rsSupply = New ADODB.Recordset
            rsSupply.Open "Select id,onhand,trecqty,lastrrdate,receipts,tissqty,issuances,lastm_cost,COST from OSMS_SUPPLY where Supply_Code = '" & vTDSupplyCode & "'", gconDMIS
            If Not rsSupply.EOF And Not rsSupply.BOF Then
                Set rsrrHEADER = New ADODB.Recordset
                rsrrHEADER.Open "Select status,rrnumber from OSMS_RRHEADER  where rrnumber = '" & Format(rsrrDETAILS!rrnumber, "000000") & "'", gconDMIS
                If Not rsrrHEADER.EOF And Not rsrrHEADER.BOF Then
                    vSUPPLYID = rsSupply!ID
                    vSupplyOnHand = N2Str2IntZero(rsSupply!ONHAND)
                    vSupplyTrecqty = N2Str2IntZero(rsSupply!trecqty)
                    vMatlastrrdate = Null2Date(rsSupply!lastrrdate)
                    vRRHEADEReceipts = N2Str2IntZero(rsSupply!receipts)
                    vCOST = N2Str2Zero(rsSupply!Cost)
                    gconDMIS.Execute "update OSMS_SUPPLY set " & _
                                     "trecqty = " & vSupplyTrecqty - vTDID_QUANTITY & ", " & _
                                     "lastrrdate = " & N2Str2Null(vTDTranDate) & _
                                   " where id =" & vSUPPLYID
                    gconDMIS.Execute "update OSMS_RRDETAILS  set status = 'P' where id = " & vTDRecNo
                Else
                    If vTDTRANS_NO <> "111111" And Null2String(rsrrDETAILS!item_no) <> "1111" Then
                        gconDMIS.Execute "insert into OSMS_noheader " & _
                                         "(trantype,TRANNO,recno,stat_h)" & _
                                       " values ('RR', '" & vTDTRANS_NO & "', " & vTDRecNo & ", '" & vTDStatus & "')"
                    End If
                End If
            Else

                gconDMIS.Execute "insert into OSMS_no_mstr " & _
                                 "(trantype,TRANNO,recno)" & _
                               " values ('RR', '" & vTDTRANS_NO & "', " & vTDRecNo & ")"

            End If

            i = i + 1

            progCPB.Value = (i / rsrrDETAILS.RecordCount) * 100

            labCPB.Caption = Int(progCPB.Value) & "% Completed"

            DoEvents

            rsrrDETAILS.MoveNext

        Loop

        labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsIssuance_Header = New ADODB.Recordset
    rsIssuance_Header.Open "select TRANS_NO,status FROM OSMS_ISSUANCE_HEADER order by TRANS_NO asc", gconDMIS
    If Not rsIssuance_Header.EOF And Not rsIssuance_Header.BOF Then
        rsIssuance_Header.MoveFirst
        i = 0
        MsgSpeech "Computing Issuances Netcost and Netprice..."
        Me.Caption = "Computing Order Netcost and Netprice..."
        Screen.MousePointer = 11
        DoEvents
        Do While Not rsIssuance_Header.EOF
            vISSUANCE_HEADERRecNo = rsIssuance_Header!Trans_No
            labProcessing.Caption = "Processing: ISS #" & Null2String(rsIssuance_Header!Trans_No)
            DoEvents
            Set rsISSUANCE_DETAILS = New ADODB.Recordset

            rsISSUANCE_DETAILS.Open "select id,TRANS_NO,cost,status,id_item_no,id_quantity from OSMS_ISSUANCE_DETAILS where TRANS_NO = " & N2Str2Null(rsIssuance_Header!Trans_No) & " order by id_item_no asc", gconDMIS

            If Not rsISSUANCE_DETAILS.EOF And Not rsISSUANCE_DETAILS.BOF Then
                rsISSUANCE_DETAILS.MoveFirst
                vNetPrice = 0: vNetCost = 0
                Do While Not rsISSUANCE_DETAILS.EOF
                    vTDNetCost = N2Str2Zero(rsISSUANCE_DETAILS!ID_Quantity) * N2Str2Zero(rsISSUANCE_DETAILS!Cost)
                    vTDStatus = Null2String(rsISSUANCE_DETAILS!Status)
                    If vTDStatus <> "C" Then vNetCost = vNetCost + vTDNetCost
                    MoveISSUANCE_DETAILS (rsISSUANCE_DETAILS!Trans_No)
                    rsISSUANCE_DETAILS.MoveNext
                Loop
                If Null2String(rsIssuance_Header!Status) <> "C" Then
                    gconDMIS.Execute "update OSMS_Issuance_Header set netcost = " & vNetCost & ", status = 'P' where trans_no = '" & vISSUANCE_HEADERRecNo & "'"
                End If
                MoveISSUANCE_HEADER (vISSUANCE_HEADERRecNo)
            End If
            i = i + 1
            progCPB.Value = (i / rsIssuance_Header.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsIssuance_Header.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsrrHEADER = New ADODB.Recordset
    rsrrHEADER.Open "select rrnumber,status from OSMS_RRHEADER  order by rrnumber asc", gconDMIS
    If Not rsrrHEADER.EOF And Not rsrrHEADER.BOF Then
        rsrrHEADER.MoveFirst
        i = 0
        Screen.MousePointer = 11
        MsgSpeech "Checking if details of receipts are already posted..."
        Me.Caption = "Checking if details of receipts are already posted..."
        DoEvents
        Do While Not rsrrHEADER.EOF
            vRRHEADERRecNo = rsrrHEADER!rrnumber
            labProcessing.Caption = "Processing: RR #" & Null2String(rsrrHEADER!rrnumber)
            DoEvents
            Set rsrrDETAILS = New ADODB.Recordset
            rsrrDETAILS.Open "select id,status,RRNUMBER,item_no from OSMS_RRDETAILS  where RRNUMBER = " & N2Str2Null(rsrrHEADER!rrnumber) & " order by item_no asc", gconDMIS
            If Not rsrrDETAILS.EOF And Not rsrrDETAILS.BOF Then
                rsrrDETAILS.MoveFirst
                Do While Not rsrrDETAILS.EOF
                    vTDRecNo = rsrrDETAILS!ID
                    vTDStatus = Null2String(rsrrDETAILS!Status)
                    If vTDStatus <> "C" Then
                        gconDMIS.Execute "update OSMS_RRDETAILS  set status = 'P' where id =" & vTDRecNo
                    End If
                    MoverrDETAILS (rsrrDETAILS!ID)
                    rsrrDETAILS.MoveNext
                Loop
            End If
            If Null2String(rsrrHEADER!Status) <> "C" Then
                gconDMIS.Execute "UPDATE OSMS_RRHEADER  set status = 'P' where rrnumber = '" & vRRHEADERRecNo & "'"
            End If
            MoveRRHEADER (vRRHEADERRecNo)
            i = i + 1
            progCPB.Value = (i / rsrrHEADER.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsrrHEADER.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsISSUANCE_DETAILS = New ADODB.Recordset
    rsISSUANCE_DETAILS.Open "select id,status,TRANS_NO,id_item_no from OSMS_ISSUANCE_DETAILS where trans_no = '000000' order by id asc", gconDMIS
    If Not rsISSUANCE_DETAILS.EOF And Not rsISSUANCE_DETAILS.BOF Then
        rsISSUANCE_DETAILS.MoveFirst
        labProcessing.Caption = "Processing: ADJ #" & Null2String(rsISSUANCE_DETAILS!Trans_No)
        DoEvents
        Do While Not rsISSUANCE_DETAILS.EOF
            vTDRecNo = rsISSUANCE_DETAILS!ID
            vTDStatus = Null2String(rsISSUANCE_DETAILS!Status)
            If vTDStatus = "N" Then
                gconDMIS.Execute "update OSMS_ISSUANCE_DETAILS set status = 'P' where id =" & vTDRecNo
            End If
            MoveISSUANCE_DETAILS (rsISSUANCE_DETAILS!ID)
            rsISSUANCE_DETAILS.MoveNext
        Loop
    End If

    Set rsrrDETAILS = New ADODB.Recordset
    rsrrDETAILS.Open "select id,status,rrnumber,item_no from OSMS_RRDETAILS  where rrnumber = '111111' order by id asc", gconDMIS
    If Not rsrrDETAILS.EOF And Not rsrrDETAILS.BOF Then
        rsrrDETAILS.MoveFirst
        labProcessing.Caption = "Processing: ADJ #" & Null2String(rsrrDETAILS!rrnumber)
        DoEvents
        Do While Not rsrrDETAILS.EOF
            vTDRecNo = rsrrDETAILS!ID
            vTDStatus = Null2String(rsrrDETAILS!Status)
            If vTDStatus = "N" Then
                gconDMIS.Execute "update OSMS_RRDETAILS  set status = 'P' where id =" & vTDRecNo
            End If
            MoverrDETAILS (rsrrDETAILS!ID)
            rsrrDETAILS.MoveNext
        Loop
    End If
    MsgSpeechBox "Posting of Transactions Completed..."
    'frmMain.mnuBatchPosting.Enabled = False

    cmdPost.Enabled = False
    Set rsISSUANCE_DETAILS = Nothing
    Set rsSupply = Nothing
    Set rsShipping = Nothing
    Set rsIssuance_Header = Nothing
    Set rsrrHEADER = Nothing
End Sub

Sub MoveISSUANCE_DETAILS(aydi As Long)
    Dim MoveSql As String

    Dim varTRANID As String

    Dim varTRANS_NO As String
    Dim varITEMNO As String
    Dim varSupply_Code As String
    Dim varID_QUANTITY As Integer
    Dim varUNIT As String
    Dim varTRANUCOST As Double
    Dim varSTATUS As String
    Dim varUSERCODE As String
    Dim varLASTUPDATE As String
    Dim rsNewISSUANCE_DETAILS As ADODB.Recordset

    Set rsNewISSUANCE_DETAILS = New ADODB.Recordset
    rsNewISSUANCE_DETAILS.Open "select * from OSMS_ISSUANCE_DETAILS where id =" & aydi, gconDMIS
    If Not rsNewISSUANCE_DETAILS.EOF And Not rsNewISSUANCE_DETAILS.BOF Then
        varTRANID = rsNewISSUANCE_DETAILS!ID
        varTRANS_NO = N2Str2Null(rsNewISSUANCE_DETAILS!Trans_No)
        varITEMNO = N2Str2Null(rsNewISSUANCE_DETAILS!id_item_no)
        varSupply_Code = N2Str2Null(rsNewISSUANCE_DETAILS!Supply_Code)
        varID_QUANTITY = N2Str2IntZero(rsNewISSUANCE_DETAILS!ID_Quantity)
        varUNIT = N2Str2Null(rsNewISSUANCE_DETAILS!ID_Unit)
        varTRANUCOST = N2Str2Zero(rsNewISSUANCE_DETAILS!Cost)
        varSTATUS = N2Str2Null(rsNewISSUANCE_DETAILS!Status)
        varUSERCODE = N2Str2Null(rsNewISSUANCE_DETAILS!USERCODE)
        varLASTUPDATE = N2Str2Null(rsNewISSUANCE_DETAILS!lastupdate)

        MoveSql = "INSERT INTO OSMS_ISSUANCE_DETHIST  " & _
                  "(TRANS_NO,ID_ITEM_NO,Supply_Code,ID_QUANTITY,id_UNIT,COST,STATUS,USERCODE,LASTUPDATE)" & _
                " values (" & varTRANS_NO & "," & varITEMNO & "," & varSupply_Code & "," & varID_QUANTITY & "," & varUNIT & "," & varTRANUCOST & "," & varSTATUS & "," & varUSERCODE & "," & varLASTUPDATE & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete from OSMS_ISSUANCE_DETAILS where id = " & varTRANID
    End If
    Set rsNewISSUANCE_DETAILS = Nothing
End Sub

Sub MoverrDETAILS(aydi As Long)
    Dim MoveSql As String

    Dim varTRANID As String
    Dim varTRANS_NO As String
    Dim varITEMNO As String
    Dim varSupply_Code As String
    Dim varID_QUANTITY As Integer
    Dim varUNIT As String
    Dim varNETCOST As Double
    Dim varSTATUS As String
    Dim varUSERCODE As String
    Dim varLASTUPDATE As String

    Dim rsNewRRDETAILS As ADODB.Recordset
    Set rsNewRRDETAILS = New ADODB.Recordset
    rsNewRRDETAILS.Open "select * from OSMS_RRDETAILS  where id =" & aydi, gconDMIS
    If Not rsNewRRDETAILS.EOF And Not rsNewRRDETAILS.BOF Then
        varTRANID = rsNewRRDETAILS!ID
        varTRANS_NO = N2Str2Null(rsNewRRDETAILS!rrnumber)
        varITEMNO = N2Str2Null(rsNewRRDETAILS!item_no)
        varSupply_Code = N2Str2Null(rsNewRRDETAILS!Supply_Code)
        varID_QUANTITY = N2Str2IntZero(rsNewRRDETAILS!rrQUANTITY)
        varUNIT = N2Str2Null(rsNewRRDETAILS!rrunit)
        varNETCOST = N2Str2Zero(rsNewRRDETAILS!Cost)
        varSTATUS = N2Str2Null(rsNewRRDETAILS!Status)
        varUSERCODE = N2Str2Null(rsNewRRDETAILS!USERCODE)
        varLASTUPDATE = N2Str2Null(rsNewRRDETAILS!lastupdate)

        MoveSql = "INSERT INTO OSMS_RRDETAILS_HIST " & _
                  "(RRNUMBER,ITEM_NO,Supply_Code,RRQUANTITY,rrUNIT,COST,STATUS,USERCODE,LASTUPDATE)" & _
                " values (" & varTRANS_NO & "," & varITEMNO & "," & varSupply_Code & "," & varID_QUANTITY & "," & varUNIT & "," & varNETCOST & "," & varSTATUS & "," & varUSERCODE & "," & varLASTUPDATE & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete from OSMS_RRDETAILS  where id = " & varTRANID
    End If
    Set rsNewRRDETAILS = Nothing
End Sub

Sub MoveISSUANCE_HEADER(aydi As Long)
    Dim MoveSql As String

    Dim varOHID As Long
    Dim varOHTRANS_NO As String
    Dim varOHTRANDATE As String
    Dim varOHSALESMAN As String
    Dim varOHSMNAME As String
    Dim varOHTTLINVAMT As Double
    Dim varOHNETCOST As Double
    Dim varOHSTATUS As String
    Dim varOHUSERCODE As String
    Dim varOHLASTUPDATE As String

    Dim rsNewISSUANCE_HEADER As ADODB.Recordset
    Set rsNewISSUANCE_HEADER = New ADODB.Recordset
    rsNewISSUANCE_HEADER.Open "select * FROM OSMS_ISSUANCE_HEADER where trans_no = ' " & aydi & "'", gconDMIS
    If Not rsNewISSUANCE_HEADER.EOF And Not rsNewISSUANCE_HEADER.BOF Then
        varOHID = rsIssuance_Header!ID
        varOHTRANS_NO = N2Str2Null(rsNewISSUANCE_HEADER!Trans_No)
        varOHTRANDATE = N2Str2Null(rsNewISSUANCE_HEADER!TRANS_DATE)
        varOHSALESMAN = N2Str2Null(rsNewISSUANCE_HEADER!Issued_by)
        varOHSMNAME = N2Str2Null(rsNewISSUANCE_HEADER!ISSUED_TO)
        varOHTTLINVAMT = N2Str2Zero(rsNewISSUANCE_HEADER!Total_Amount)
        varOHNETCOST = N2Str2Zero(rsNewISSUANCE_HEADER!netCOST)
        varOHSTATUS = N2Str2Null(rsNewISSUANCE_HEADER!Status)
        varOHUSERCODE = N2Str2Null(rsNewISSUANCE_HEADER!USERCODE)
        varOHLASTUPDATE = N2Str2Null(rsNewISSUANCE_HEADER!lastupdate)

        MoveSql = "INSERT INTO ISSUANCE_HEADHIST " & _
                  "(TRANS_NO,TRANS_DATE,ISSUED_BY,ISSUED_TO,TOTAL_AMOUNT,NETCOST,STATUS,USERCODE,LASTUPDATE)" & _
                " values (" & varOHTRANS_NO & ", " & varOHTRANDATE & ", " & varOHSALESMAN & ", " & varOHSMNAME & ", " & varOHTTLINVAMT & ", " & varOHNETCOST & ", " & varOHSTATUS & ", " & varOHUSERCODE & ", " & varOHLASTUPDATE & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete FROM OSMS_ISSUANCE_HEADER where trans_no = " & varOHTRANS_NO
    End If
    Set rsNewISSUANCE_HEADER = Nothing
End Sub

Sub MoveRRHEADER(aydi As Long)
    Dim MoveSql As String

    Dim varRRRRNO As String
    Dim varRRRRDATE As String
    Dim varRRPONO As String
    Dim varRRPODATE As String
    Dim varRRTTLRRAMT As Double
    Dim varRRNETRRAMT As Double
    Dim varRRSTATUS As String
    Dim varRRUSERCODE As String
    Dim varRRLASTUPDATE As String
    Dim varRRREMARKS As String

    Dim varRRINV_NO As String
    Dim varRRINV_DATE As String
    Dim varRRSupplier_Code As String
    Dim varRRReceivedby_Code As String

    Dim rsNewRRHEADER As ADODB.Recordset
    Set rsNewRRHEADER = New ADODB.Recordset
    rsNewRRHEADER.Open "select * from OSMS_RRHEADER  where rrnumber = '" & aydi & "'", gconDMIS
    If Not rsNewRRHEADER.EOF And Not rsNewRRHEADER.BOF Then
        'varRRID = rsRRHEADER!Id
        varRRRRNO = N2Str2Null(rsNewRRHEADER!rrnumber)
        varRRRRDATE = N2Str2Null(rsNewRRHEADER!RRDATE)
        varRRINV_NO = N2Str2Null(rsNewRRHEADER!inv_no)
        varRRINV_DATE = N2Str2Null(rsNewRRHEADER!inv_date)
        varRRPONO = N2Str2Null(rsNewRRHEADER!PONO)
        varRRPODATE = N2Str2Null(rsNewRRHEADER!PODATE)
        varRRSupplier_Code = N2Str2Null(rsNewRRHEADER!Supplier_code)
        varRRReceivedby_Code = N2Str2Null(rsNewRRHEADER!Receivedby_Code)
        varRRTTLRRAMT = N2Str2Zero(rsNewRRHEADER!Total_Amount)
        varRRNETRRAMT = N2Str2Zero(rsNewRRHEADER!netCOST)
        varRRSTATUS = N2Str2Null(rsNewRRHEADER!Status)
        varRRUSERCODE = N2Str2Null(rsNewRRHEADER!USERCODE)
        varRRLASTUPDATE = N2Str2Null(rsNewRRHEADER!lastupdate)
        varRRREMARKS = N2Str2Null(rsNewRRHEADER!PURPOSE)

        MoveSql = "INSERT INTO OSMS_RRHEADER_HIST " & _
                  "(RRNO,RRDATE,INV_NO,INV_DATE,PONO,PODATE,SUPPLIER_CODE,RECEIVEDBY_CODE,TOTAL_AMOUNT,NETCOST,STATUS,USERCODE,LASTUPDATE,PURPOSE)" & _
                " values (" & varRRRRNO & ", " & varRRRRDATE & ", " & varRRINV_NO & "," & varRRINV_DATE & ", " & varRRPONO & ", " & varRRPODATE & ", " & varRRSupplier_Code & ", " & varRRReceivedby_Code & ", " & varRRTTLRRAMT & ", " & varRRSTATUS & ", " & varRRUSERCODE & ", " & varRRLASTUPDATE & ", " & varRRREMARKS & ")"
        gconDMIS.Execute MoveSql
        gconDMIS.Execute "delete from OSMS_RRHEADER  where rrnumber = " & varRRRRNO
    End If
    Set rsNewRRHEADER = Nothing
End Sub

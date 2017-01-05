VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmPMISProcess_UpdateMaster 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Master File"
   ClientHeight    =   6240
   ClientLeft      =   270
   ClientTop       =   360
   ClientWidth     =   8940
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FF8080&
   Icon            =   "UpdateMaster.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8940
   Begin VB.Frame fraCurrentActivity 
      BackColor       =   &H00FF8080&
      Caption         =   "Current Activity"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      Left            =   60
      TabIndex        =   13
      Top             =   1740
      Width           =   8805
      Begin RichTextLib.RichTextBox txtCurrentActivity 
         Height          =   3855
         Left            =   120
         TabIndex        =   14
         Top             =   330
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   6800
         _Version        =   393217
         BackColor       =   8454143
         TextRTF         =   $"UpdateMaster.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
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
      Left            =   8100
      MouseIcon       =   "UpdateMaster.frx":0388
      MousePointer    =   99  'Custom
      Picture         =   "UpdateMaster.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Exit Window"
      Top             =   930
      Width           =   735
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Update"
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
      Left            =   7380
      MouseIcon       =   "UpdateMaster.frx":0840
      MousePointer    =   99  'Custom
      Picture         =   "UpdateMaster.frx":0992
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Update Master File"
      Top             =   930
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update Last Month Details"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   8070
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update Last Month MAC"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   7650
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
      TabIndex        =   0
      Top             =   7350
      Value           =   1  'Checked
      Width           =   5685
   End
   Begin VB.PictureBox picCPB 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1095
      Left            =   30
      ScaleHeight     =   1095
      ScaleWidth      =   8805
      TabIndex        =   1
      Top             =   60
      Width           =   8805
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         ScaleHeight     =   255
         ScaleWidth      =   7125
         TabIndex        =   2
         Top             =   720
         Width           =   7125
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
            Height          =   195
            Left            =   30
            TabIndex        =   3
            Top             =   30
            Width           =   7065
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         ToolTipText     =   "Update progress"
         Top             =   300
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   556
         Picture         =   "UpdateMaster.frx":0C2D
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "UpdateMaster.frx":0C49
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
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   7215
         TabIndex        =   4
         Top             =   660
         Width           =   7215
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
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   30
         Width           =   5595
      End
   End
   Begin wizProgBar.Prg Prg1 
      Height          =   315
      Left            =   90
      TabIndex        =   8
      ToolTipText     =   "Update progress"
      Top             =   7680
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      Picture         =   "UpdateMaster.frx":0C65
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "UpdateMaster.frx":0C81
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
      TabIndex        =   10
      ToolTipText     =   "Update progress"
      Top             =   8100
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      Picture         =   "UpdateMaster.frx":0C9D
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "UpdateMaster.frx":0CB9
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
End
Attribute VB_Name = "frmPMISProcess_UpdateMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function CheckIfNonVatSup(SupplierCode As String) As Boolean
    Dim rsSupplierMaster                               As ADODB.Recordset
    Set rsSupplierMaster = New ADODB.Recordset
    rsSupplierMaster.Open "Select supcode,supname,NONVAT from PMIS_vw_Supplier where supcode = '" & SupplierCode & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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

Sub UpdateMaster()
    Dim RSPARTMAS, rsCURPartmas, RSTDAYTRAN, RSPO_HD, rsRR_HD, RSORD_HD As ADODB.Recordset
    Dim I                                              As Long
    Dim vTotTranCost, vTotTranInvAmt, vTDTranQTY       As Double
    Dim vTDTranDate, vTDTranType, vTDTranno, vTDType   As String
    Dim vVatAmt, vMAC                                  As Double
    Dim vPMOnhand                                      As Integer
    Dim vSTOCKDESC                                     As String

    Dim vTotalQty                                      As Long
    Dim vOrdHDRecNo                                    As Long
    Dim INCLUDE_UPDATE_MAC                             As Boolean

    INCLUDE_UPDATE_MAC = False
    Set RSPO_HD = New ADODB.Recordset
    RSPO_HD.Open "select id,pono,status,TYPE from PMIS_PO_Hd order by pono asc", gconDMIS
    If Not RSPO_HD.EOF And Not RSPO_HD.BOF Then
        RSPO_HD.MoveFirst: I = 0: Screen.MousePointer = 11
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Computing Total Quantity of Purchases......": DoEvents
        Do While Not RSPO_HD.EOF
            vOrdHDRecNo = RSPO_HD!ID
            labProcessing.Caption = "Processing: PO #" & Null2String(RSPO_HD!Type) & "-" & Null2String(RSPO_HD!PONO)
            DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,status,itemno from PMIS_TdayTran where [TYPE] = '" & Null2String(RSPO_HD!Type) & "' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst: vTotalQty = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!tranqty)
                    gconDMIS.Execute "Update PMIS_TdayTran SET STATUS = '" & RSPO_HD!STATUS & "' where ID = " & RSTDAYTRAN!ID
                    RSTDAYTRAN.MoveNext
                Loop
                gconDMIS.Execute "update PMIS_PO_Hd set TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
            End If
            I = I + 1
            progCPB.Value = (I / RSPO_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            RSPO_HD.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents
        Screen.MousePointer = 0
    End If
    Set RSPO_HD = Nothing
    Set RSTDAYTRAN = Nothing


    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select id,rrno,status,TYPE from PMIS_RR_Hd order by rrno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        rsRR_HD.MoveFirst: I = 0: Screen.MousePointer = 11
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Computing Total Quantity of Receiving......": DoEvents
        Do While Not rsRR_HD.EOF
            vOrdHDRecNo = rsRR_HD!ID
            labProcessing.Caption = "Processing: RR #" & Null2String(rsRR_HD!Type) & "-" & Null2String(rsRR_HD!RRNO)
            DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,status,itemno from PMIS_TdayTran where [TYPE] = '" & Null2String(rsRR_HD!Type) & "' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!RRNO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst
                vTotalQty = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!tranqty)
                    gconDMIS.Execute "Update PMIS_TdayTran SET STATUS = '" & Null2String(rsRR_HD!STATUS) & "' where ID = " & RSTDAYTRAN!ID
                    RSTDAYTRAN.MoveNext
                Loop
                gconDMIS.Execute "update PMIS_RR_Hd set TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
            End If
            I = I + 1
            progCPB.Value = (I / rsRR_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            rsRR_HD.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents
        Screen.MousePointer = 0
    End If
    Set rsRR_HD = Nothing
    Set RSTDAYTRAN = Nothing


    Dim vTotalTranCost                                 As Double
    Set RSORD_HD = New ADODB.Recordset
    RSORD_HD.Open "select id,trantype,tranno,status,TYPE from PMIS_Ord_Hd order by trantype,tranno asc", gconDMIS
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        RSORD_HD.MoveFirst: I = 0: Screen.MousePointer = 11
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Computing Total Quantity of Request and Issuance......": DoEvents
        Do While Not RSORD_HD.EOF
            vOrdHDRecNo = RSORD_HD!ID
            labProcessing.Caption = "Processing: " & Null2String(RSORD_HD!Type) & "-" & Null2String(RSORD_HD!TranType) & " #" & Null2String(RSORD_HD!TRANNO): DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,tranucost,status,itemno from PMIS_TdayTran where [TYPE] = '" & Null2String(RSORD_HD!Type) & "' AND trantype = " & N2Str2Null(RSORD_HD!TranType) & " and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst
                vTotalQty = 0: vTotalTranCost = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!tranqty)
                    vTotalTranCost = vTotalTranCost + (N2Str2Zero(RSTDAYTRAN!TRANUCOST) * N2Str2Zero(RSTDAYTRAN!tranqty))
                    gconDMIS.Execute "Update PMIS_TdayTran SET STATUS = '" & Null2String(RSORD_HD!STATUS) & "' where ID = " & RSTDAYTRAN!ID
                    RSTDAYTRAN.MoveNext
                Loop
                gconDMIS.Execute "update PMIS_Ord_Hd set NETCOST = " & vTotalTranCost & ", TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
            End If
            I = I + 1
            progCPB.Value = (I / RSORD_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            RSORD_HD.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents
        Screen.MousePointer = 0
    End If
    Set RSORD_HD = Nothing
    Set RSTDAYTRAN = Nothing

    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select id,ItemNo,trantype,tranno,TYPE,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIS_TdayTran where trantype <> 'ADB' and (status = 'P' OR status = 'B') order by id asc", gconDMIS
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst: Screen.MousePointer = 11
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Storing Stocks Master Beginning Balances...": DoEvents
        If Month(Null2String(RSTDAYTRAN!TRANDATE)) = 1 Then
            If INCLUDE_UPDATE_MAC = True Then
                gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                               " lasty_oh = ISNULL(lastm_oh,0)," & _
                               " lasty_mac = ISNULL(lastm_mac,0)," & _
                               " lasty_oo = ISNULL(lastm_oo,0)," & _
                               " onhand = ISNULL(lastm_oh,0)," & _
                               " mac = ISNULL(lastm_mac,0)," & _
                               " onorder = ISNULL(lastm_oo,0)," & _
                               " tpoqty = 0," & _
                               " tissqty = 0," & _
                               " trecqty = 0," & _
                               " purchases = 0," & _
                               " receipts = 0," & _
                               " issuances = 0 WHERE ACTIVE = 'Y'"
            Else
                gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                               " lasty_oh = ISNULL(lastm_oh,0)," & _
                               " lasty_oo = ISNULL(lastm_oo,0)," & _
                               " onhand = ISNULL(lastm_oh,0)," & _
                               " onorder = ISNULL(lastm_oo,0)," & _
                               " tpoqty = 0," & _
                               " tissqty = 0," & _
                               " trecqty = 0," & _
                               " purchases = 0," & _
                               " receipts = 0," & _
                               " issuances = 0 WHERE ACTIVE = 'Y'"
            End If
        Else
            gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                           " purchases = ISNULL(purchases,0) - ISNULL(tpoqty,0)," & _
                           " receipts = ISNULL(receipts,0) - ISNULL(trecqty,0)," & _
                           " issuances = ISNULL(issuances,0) - ISNULL(tissqty,0)"
            If INCLUDE_UPDATE_MAC = True Then
                gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                               " onhand = ISNULL(lastm_oh,0)," & _
                               " mac = ISNULL(lastm_mac,0)," & _
                               " onorder = ISNULL(lastm_oo,0)," & _
                               " tpoqty = 0," & _
                               " tissqty = 0," & _
                               " trecqty = 0 WHERE ACTIVE = 'Y'"
            Else
                gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                               " onhand = ISNULL(lastm_oh,0)," & _
                               " onorder = ISNULL(lastm_oo,0)," & _
                               " tpoqty = 0," & _
                               " tissqty = 0," & _
                               " trecqty = 0 WHERE ACTIVE = 'Y'"
            End If
        End If
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " ONREQUEST = 0," & _
                       " REQSERVED = 0," & _
                       " REQUNSERVED = 0," & _
                       " REQFILLRATE = 0," & _
                       " S_ONREQUEST = 0," & _
                       " S_REQSERVED = 0," & _
                       " S_REQUNSERVED = 0," & _
                       " S_REQFILLRATE = 0," & _
                       " ORDERED = 0," & _
                       " ONORDER = 0," & _
                       " SERVED = 0," & _
                       " UNSERVED = 0," & _
                       " BACKORDER = 0," & _
                       " FILLRATE = 0," & _
                       " EMERGENCY_PO = 0 WHERE ACTIVE = 'Y'"
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " STOCKTYPE = 'GJ' WHERE (STOCKTYPE <> 'BP' AND LEFT(STOCKNO,2) <> '08')"
        gconDMIS.Execute "update PMIS_STKSTAT set" & _
                       " STOCKTYPE = 'GJ' WHERE (STOCKTYPE <> 'BP' AND LEFT(STOCKNO,2) <> '08')"
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " NON_HARI = 'N' WHERE (NON_HARI IS NULL OR (NON_HARI <> 'Y' AND NON_HARI <> 'N' AND NON_HARI <> 'O'))"
        gconDMIS.Execute "update PMIS_STKSTAT set" & _
                       " NON_HARI = 'N' WHERE (NON_HARI IS NULL OR (NON_HARI <> 'Y' AND NON_HARI <> 'N' AND NON_HARI <> 'O'))"

        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Updating Stock Master from Transactions Made...": DoEvents
        I = 0
        Do While Not RSTDAYTRAN.EOF
            gconDMIS.Execute "update PMIS_TdayTran set ItemNo = '" & Format(Null2String(RSTDAYTRAN!ITEMNO), "0000") & "' where ID = " & RSTDAYTRAN!ID
            vTDType = Null2String(RSTDAYTRAN![Type])
            vTDTranDate = N2Date2Null(RSTDAYTRAN!TRANDATE)
            vTDTranType = Null2String(RSTDAYTRAN!TranType)
            vTDTranno = Null2String(RSTDAYTRAN!TRANNO)
            vTDTranQTY = N2Str2Zero(RSTDAYTRAN!tranqty)
            vTotTranCost = N2Str2Zero(RSTDAYTRAN!TRANUCOST) * vTDTranQTY
            vTotTranInvAmt = N2Str2Zero(RSTDAYTRAN!TRANINVAMT) * vTDTranQTY
            labProcessing.Caption = "Processing: " & Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(RSTDAYTRAN!TRANNO): DoEvents

            Set RSPARTMAS = New ADODB.Recordset
            Set RSPARTMAS = gconDMIS.Execute("select STOCKNO from PMIS_STOCKMAS where STOCKNO = " & N2Str2Null(RSTDAYTRAN!STOCK_ORD))
            If RSPARTMAS.EOF And RSPARTMAS.BOF Then
                If vTDType = "P" Then
                    Set rsCURPartmas = New ADODB.Recordset
                    Set rsCURPartmas = gconDMIS.Execute("Select PARTNUMBER,DESCRIPTIO from PMIS_DNPP where PARTNUMBER = " & N2Str2Null(RSTDAYTRAN!STOCK_ORD))
                    If Not rsCURPartmas.EOF And Not rsCURPartmas.BOF Then
                        vSTOCKDESC = N2Str2Null(rsCURPartmas!DESCRIPTIO)
                    Else
                        vSTOCKDESC = "'NO DESCRIPTION'"
                    End If
                Else
                    vSTOCKDESC = "'NO DESCRIPTION'"
                End If
                txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Inserting Found New Stock No. (" & Null2String(RSTDAYTRAN!STOCK_ORD) & ")": DoEvents
                gconDMIS.Execute ("Insert into PMIS_STOCKMAS (TYPE,STOCKNO,STOCKDESC,date_entered) values ('" & vTDType & "'," & N2Str2Null(RSTDAYTRAN!STOCK_ORD) & "," & vSTOCKDESC & "," & N2Str2Null(RSTDAYTRAN!TRANDATE) & ")")
                txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Continued: Updating Stock Master from Transactions Made...": DoEvents
            Else
                gconDMIS.Execute ("Update PMIS_STOCKMAS SET ACTIVE = 'Y', TYPE = '" & vTDType & "' WHERE STOCKNO = " & N2Str2Null(RSTDAYTRAN!STOCK_ORD))
            End If
            Set RSPARTMAS = New ADODB.Recordset
            Set RSPARTMAS = gconDMIS.Execute("select id,STOCKNO,mac,Onhand,NON_HARI from PMIS_STOCKMAS where TYPE = '" & vTDType & "' AND STOCKNO = " & N2Str2Null(RSTDAYTRAN!STOCK_ORD))
            If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                vMAC = N2Str2Zero(RSPARTMAS!Mac)
                vPMOnhand = N2Str2IntZero(RSPARTMAS!ONHAND)
                If Null2String(RSTDAYTRAN!IN_OUT) = "R" And vTDTranQTY <> 0 Then
                    Set RSORD_HD = New ADODB.Recordset
                    Set RSORD_HD = gconDMIS.Execute("Select sales_origin from PMIS_Ord_HD where [TYPE] = '" & vTDType & "' AND trantype = '" & vTDTranType & "' and tranno = '" & Format(vTDTranno, "000000") & "'")
                    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                        If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "ONREQUEST = ISNULL(ONREQUEST,0) + " & vTDTranQTY & _
                                           " where id = " & RSPARTMAS!ID
                        Else
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "S_ONREQUEST = ISNULL(S_ONREQUEST,0) + " & vTDTranQTY & _
                                           " where id = " & RSPARTMAS!ID
                        End If
                    End If
                End If
                If Null2String(RSTDAYTRAN!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    Set RSORD_HD = New ADODB.Recordset
                    Set RSORD_HD = gconDMIS.Execute("Select sales_origin from PMIS_Ord_HD where [TYPE] = '" & vTDType & "' AND trantype = '" & vTDTranType & "' and tranno = '" & Format(vTDTranno, "000000") & "'")
                    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                        If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "onhand = ISNULL(onhand,0) - " & vTDTranQTY & ", " & _
                                             "REQSERVED = ISNULL(REQSERVED,0) + " & vTDTranQTY & ", " & _
                                             "tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
                                             "issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
                                           " where id = " & RSPARTMAS!ID
                        Else
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "onhand = ISNULL(onhand,0) - " & vTDTranQTY & ", " & _
                                             "S_REQSERVED = ISNULL(S_REQSERVED,0) + " & vTDTranQTY & ", " & _
                                             "tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
                                             "issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
                                           " where id = " & RSPARTMAS!ID
                        End If
                    Else
                        If vTDTranType = "ADJ" Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "onhand = ISNULL(onhand,0) - " & vTDTranQTY & ", " & _
                                             "tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
                                             "issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
                                           " where id = " & RSPARTMAS!ID
                        End If
                    End If
                    'EAP: remove the code that update the unitcost in pmis_tdaytran
                    'gconDMIS.Execute "update PMIS_TdayTran set NON_HARI = " & N2Str2Null(rsPartMas!NON_HARI) & ", tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TdayTran set NON_HARI = " & N2Str2Null(RSPARTMAS!NON_HARI) & " where ID = " & RSTDAYTRAN!ID
                End If

                If Null2String(RSTDAYTRAN!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    If vPMOnhand <= 0 Then
                        vMAC = vTotTranCost / vTDTranQTY
                        If INCLUDE_UPDATE_MAC = True Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "mac = " & vMAC & ", " & _
                                             "Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & ", " & _
                                             "ONORDER = ISNULL(ONORDER,0) - " & vTDTranQTY & ", " & _
                                             "SERVED = ISNULL(SERVED,0) + " & vTDTranQTY & ", " & _
                                             "last_recd = " & vTDTranDate & ", " & _
                                             "trecqty = ISNULL(trecqty,0) + " & vTDTranQTY & ", " & _
                                             "receipts = ISNULL(receipts,0) + " & vTDTranQTY & _
                                           " where id =" & RSPARTMAS!ID
                        Else
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & ", " & _
                                             "ONORDER = ISNULL(ONORDER,0) - " & vTDTranQTY & ", " & _
                                             "SERVED = ISNULL(SERVED,0) + " & vTDTranQTY & ", " & _
                                             "last_recd = " & vTDTranDate & ", " & _
                                             "trecqty = ISNULL(trecqty,0) + " & vTDTranQTY & ", " & _
                                             "receipts = ISNULL(receipts,0) + " & vTDTranQTY & _
                                           " where id =" & RSPARTMAS!ID
                        End If
                    Else

                        vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                        If INCLUDE_UPDATE_MAC = True Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "mac = " & vMAC & ", " & _
                                             "ONORDER = ISNULL(ONORDER,0) - " & vTDTranQTY & ", " & _
                                             "SERVED = ISNULL(SERVED,0) + " & vTDTranQTY & ", " & _
                                             "Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & ", " & _
                                             "last_recd = " & vTDTranDate & ", " & _
                                             "trecqty = ISNULL(trecqty,0) + " & vTDTranQTY & ", " & _
                                             "receipts = ISNULL(receipts,0) + " & vTDTranQTY & _
                                           " where id =" & RSPARTMAS!ID
                        Else
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "ONORDER = ISNULL(ONORDER,0) - " & vTDTranQTY & ", " & _
                                             "SERVED = ISNULL(SERVED,0) + " & vTDTranQTY & ", " & _
                                             "Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & ", " & _
                                             "last_recd = " & vTDTranDate & ", " & _
                                             "trecqty = ISNULL(trecqty,0) + " & vTDTranQTY & ", " & _
                                             "receipts = ISNULL(receipts,0) + " & vTDTranQTY & _
                                           " where id =" & RSPARTMAS!ID
                        End If
                    End If
                    If INCLUDE_UPDATE_MAC = True Then
                        gconDMIS.Execute "update PMIS_TdayTran set NON_HARI = " & N2Str2Null(RSPARTMAS!NON_HARI) & ", mac = " & vMAC & " where id = " & RSTDAYTRAN!ID
                    Else
                        gconDMIS.Execute "update PMIS_TdayTran set NON_HARI = " & N2Str2Null(RSPARTMAS!NON_HARI) & " where id = " & RSTDAYTRAN!ID
                    End If
                End If
                If Null2String(RSTDAYTRAN!TranType) = "PO" Then
                    gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                     "purchases = ISNULL(purchases,0) + " & vTDTranQTY & "," & _
                                     "tpoqty = ISNULL(tpoqty,0) + " & vTDTranQTY & "," & _
                                     "ONORDER = ISNULL(ONORDER,0) + " & vTDTranQTY & "," & _
                                     "ORDERED = ISNULL(ORDERED,0) + " & vTDTranQTY & _
                                   " where id = " & RSPARTMAS!ID
                End If
            End If
            I = I + 1
            progCPB.Value = (I / RSTDAYTRAN.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            RSTDAYTRAN.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents
        Screen.MousePointer = 0
    End If
    txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Updating of Stock Master Completed...": DoEvents
    MsgSpeechBox "Updating of Master File Completed..."
End Sub

Private Sub cmdCheck_Click()
    If Function_Access(LOGID, "Acess_Process", "UPDATE MASTER FILE") = False Then Exit Sub
    cmdCheck.Enabled = False
    cmdExit.Enabled = False
    DoEvents
    UpdateMaster
    cmdExit.Enabled = True
    DoEvents
    NEW_LogAudit "R", "UPDATE MASTER FILE", "", "", "", " - ", "", ""

    Exit Sub
Errorcode:
    ShowVBError

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
    Dim RSTDAYTRAN                                     As ADODB.Recordset
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select trantype from PMIS_TdayTran where trantype = 'ADJ'", gconDMIS
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        chkUpdateAdjustment.Enabled = True
        chkUpdateAdjustment.Value = 1
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISProcess_UpdateMaster = Nothing
    UnloadForm Me
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub


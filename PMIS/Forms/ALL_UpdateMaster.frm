VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmPMISALL_UpdateMaster 
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
   Icon            =   "ALL_UpdateMaster.frx":0000
   LinkTopic       =   "Form1"
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
      TabIndex        =   17
      Top             =   1740
      Width           =   8805
      Begin RichTextLib.RichTextBox txtCurrentActivity 
         Height          =   3855
         Left            =   120
         TabIndex        =   18
         Top             =   330
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   6800
         _Version        =   393217
         BackColor       =   8454143
         Enabled         =   -1  'True
         TextRTF         =   $"ALL_UpdateMaster.frx":030A
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
      MouseIcon       =   "ALL_UpdateMaster.frx":0388
      MousePointer    =   99  'Custom
      Picture         =   "ALL_UpdateMaster.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   16
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
      MouseIcon       =   "ALL_UpdateMaster.frx":0840
      MousePointer    =   99  'Custom
      Picture         =   "ALL_UpdateMaster.frx":0992
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Update Master File"
      Top             =   930
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update Last Month Details"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   8070
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update Last Month MAC"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
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
      TabIndex        =   2
      Top             =   7350
      Value           =   1  'Checked
      Width           =   5685
   End
   Begin MSMask.MaskEdBox txtFrom 
      Height          =   345
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "Input starting date"
      Top             =   1230
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
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
      Top             =   1230
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
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
   Begin VB.PictureBox picCPB 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1095
      Left            =   30
      ScaleHeight     =   1095
      ScaleWidth      =   8805
      TabIndex        =   5
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
         TabIndex        =   6
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
            TabIndex        =   7
            Top             =   30
            Width           =   7065
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   9
         ToolTipText     =   "Update progress"
         Top             =   300
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   556
         Picture         =   "ALL_UpdateMaster.frx":0C2D
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "ALL_UpdateMaster.frx":0C49
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
         TabIndex        =   8
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
         TabIndex        =   10
         Top             =   30
         Width           =   5595
      End
   End
   Begin wizProgBar.Prg Prg1 
      Height          =   315
      Left            =   90
      TabIndex        =   12
      ToolTipText     =   "Update progress"
      Top             =   7680
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      Picture         =   "ALL_UpdateMaster.frx":0C65
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "ALL_UpdateMaster.frx":0C81
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
      TabIndex        =   14
      ToolTipText     =   "Update progress"
      Top             =   8100
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      Picture         =   "ALL_UpdateMaster.frx":0C9D
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "ALL_UpdateMaster.frx":0CB9
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
      TabIndex        =   4
      Top             =   1230
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
      TabIndex        =   3
      Top             =   1230
      Width           =   765
   End
End
Attribute VB_Name = "frmPMISALL_UpdateMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
    If Function_Access(LOGID, "Acess_Process", "UPDATE MASTER FILE") = False Then Exit Sub
    Dim rsTdayTran                      As ADODB.Recordset

    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select trandate from PMIS_TdayTran where trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "'", gconDMIS
    If rsTdayTran.EOF And rsTdayTran.BOF Then
        MsgSpeechBox "No Transactions Made from " & txtFrom.Text & " to " & txtTo.Text
        Exit Sub
    End If
    txtFrom.Enabled = False
    txtTo.Enabled = False
    cmdCheck.Enabled = False
    cmdExit.Enabled = False
    DoEvents
    UpdateMaster
    cmdExit.Enabled = True
    LogAudit "G", "UPDATE MASTERFILE", txtFrom & "-" & txtTo

    DoEvents

    Exit Sub
ErrorCode:
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
    'Me.Height = 2025
    Dim rsTdayTran                      As ADODB.Recordset
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select trantype from PMIS_TdayTran where trantype = 'ADJ'", gconDMIS
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        chkUpdateAdjustment.Enabled = True
        chkUpdateAdjustment.Value = 1
    End If
    txtFrom.Text = firstDay(LOGDATE)
    txtTo.Text = LOGDATE
    Screen.MousePointer = 0
End Sub

Sub UpdateMaster()
    Dim rsPartMas, rsCURPartmas, rsTdayTran, rsPO_HD, rsRR_HD, rsOrd_Hd As ADODB.Recordset
    Dim I                               As Integer
    Dim vTotTranCost, vTotTranInvAmt, vTotTranQTY, vTDTranQTY As Double
    Dim vTDTranDate, vTDTranType, vTDTranno, vSupCode, vTDType As String
    Dim vVatAmt, vMAC                   As Double
    Dim vPMOnhand                       As Integer
    Dim vSTOCKDESC                      As String

    Dim vTotalQty                       As Long
    Dim vOrdHDRecNo                     As Long
        
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,ItemNo,trantype,tranno,TYPE,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIS_AlldayTran where trantype <> 'ADB' and (status = 'P' OR status = 'B') AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by id asc", gconDMIS
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        rsTdayTran.MoveFirst: Screen.MousePointer = 11
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Storing Stocks Master Beginning Balances...": DoEvents
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " purchases = 0," & _
                       " receipts = 0," & _
                       " issuances = 0"
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " onhand = 0," & _
                       " mac = 0," & _
                       " onorder = 0," & _
                       " Lastm_OH = 0," & _
                       " Lastm_mac = 0," & _
                       " Lastm_oo = 0," & _
                       " tpoqty = 0," & _
                       " tissqty = 0," & _
                       " trecqty = 0 WHERE ACTIVE = 'Y'"
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
                       
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Updating Stock Master from Transactions Made...": DoEvents
        I = 0
        Do While Not rsTdayTran.EOF
            gconDMIS.Execute "update PMIS_TdayTran set ItemNo = '" & Format(Null2String(rsTdayTran!itemno), "0000") & "' where ID = " & rsTdayTran!ID
            gconDMIS.Execute "update PMIS_dayTran set ItemNo = '" & Format(Null2String(rsTdayTran!itemno), "0000") & "' where ID = " & rsTdayTran!ID
            vTDType = Null2String(rsTdayTran![Type])
            vTDTranDate = N2Date2Null(rsTdayTran!trandate)
            vTDTranType = Null2String(rsTdayTran!TRANTYPE)
            vTDTranno = Null2String(rsTdayTran!tranno)
            vTDTranQTY = N2Str2Zero(rsTdayTran!tranqty)
            vTotTranCost = N2Str2Zero(rsTdayTran!TRANUCOST) * vTDTranQTY
            vTotTranInvAmt = N2Str2Zero(rsTdayTran!TRANINVAMT) * vTDTranQTY
            labProcessing.Caption = "Processing: " & Null2String(rsTdayTran!TRANTYPE) & " #" & Null2String(rsTdayTran!tranno): DoEvents
            
            Set rsPartMas = New ADODB.Recordset
            Set rsPartMas = gconDMIS.Execute("select STOCKNO from PMIS_STOCKMAS where STOCKNO = " & N2Str2Null(rsTdayTran!STOCK_ORD))
            If rsPartMas.EOF And rsPartMas.BOF Then
                If vTDType = "P" Then
                   Set rsCURPartmas = New ADODB.Recordset
                   Set rsCURPartmas = gconDMIS.Execute("Select PARTNUMBER,DESCRIPTIO from PMIS_DNPP where PARTNUMBER = " & N2Str2Null(rsTdayTran!STOCK_ORD))
                   If Not rsCURPartmas.EOF And Not rsCURPartmas.BOF Then
                       vSTOCKDESC = N2Str2Null(rsCURPartmas!DESCRIPTIO)
                   Else
                       vSTOCKDESC = "'NO DESCRIPTION'"
                   End If
                Else
                   vSTOCKDESC = "'NO DESCRIPTION'"
                End If
                txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Inserting Found New Stock No. (" & Null2String(rsTdayTran!STOCK_ORD) & ")": DoEvents
                gconDMIS.Execute ("Insert into PMIS_STOCKMAS (TYPE,STOCKNO,STOCKDESC,date_entered) values ('" & vTDType & "'," & N2Str2Null(rsTdayTran!STOCK_ORD) & "," & vSTOCKDESC & "," & N2Str2Null(rsTdayTran!trandate) & ")")
                txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Continued: Updating Stock Master from Transactions Made...": DoEvents
            Else
                gconDMIS.Execute ("Update PMIS_STOCKMAS SET ACTIVE = 'Y', TYPE = '" & vTDType & "' WHERE STOCKNO = " & N2Str2Null(rsTdayTran!STOCK_ORD))
            End If
            Set rsPartMas = New ADODB.Recordset
            'rsPartMas.Open "select id,STOCKNO,mac,tissqty,trecqty,poqty,tprqty,prqty,issuances,receipts,onhand,ONREQUEST,REQSERVED,SERVED,ONORDER,ORDERED,tpoqty,S_REQSERVED,S_ONREQUEST,purchases from PMIS_STOCKMAS where TYPE = '" & vTDType & "' AND STOCKNO = " & N2Str2Null(rsTdayTran!STOCK_ORD), gconDMIS
            Set rsPartMas = gconDMIS.Execute("select id,STOCKNO,mac,Onhand from PMIS_STOCKMAS where TYPE = '" & vTDType & "' AND STOCKNO = " & N2Str2Null(rsTdayTran!STOCK_ORD))
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                vMAC = N2Str2Zero(rsPartMas!Mac)
                vPMOnhand = N2Str2IntZero(rsPartMas!ONHAND)
                If Null2String(rsTdayTran!IN_OUT) = "R" And vTDTranQTY <> 0 Then
                    Set rsOrd_Hd = New ADODB.Recordset
                    Set rsOrd_Hd = gconDMIS.Execute("Select sales_origin from PMIS_Ord_HD where [TYPE] = '" & vTDType & "' AND trantype = '" & vTDTranType & "' and tranno = '" & Format(vTDTranno, "000000") & "'")
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        If Null2String(rsOrd_Hd!SALES_ORIGIN) = "W" Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "ONREQUEST = ISNULL(ONREQUEST,0) + " & vTDTranQTY & _
                                             " where id = " & rsPartMas!ID
                        Else
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "S_ONREQUEST = ISNULL(S_ONREQUEST,0) + " & vTDTranQTY & _
                                             " where id = " & rsPartMas!ID
                        End If
                    End If
                End If
                If Null2String(rsTdayTran!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    Set rsOrd_Hd = New ADODB.Recordset
                    Set rsOrd_Hd = gconDMIS.Execute("Select sales_origin from PMIS_vw_IS_HISTORY where [TYPE] = '" & vTDType & "' AND trantype = '" & vTDTranType & "' and tranno = '" & Format(vTDTranno, "000000") & "'")
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        If Null2String(rsOrd_Hd!SALES_ORIGIN) = "W" Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "onhand = ISNULL(onhand,0) - " & vTDTranQTY & ", " & _
                                             "REQSERVED = ISNULL(REQSERVED,0) + " & vTDTranQTY & ", " & _
                                             "tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
                                             "issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
                                             " where id = " & rsPartMas!ID
                        Else
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "onhand = ISNULL(onhand,0) - " & vTDTranQTY & ", " & _
                                             "S_REQSERVED = ISNULL(S_REQSERVED,0) + " & vTDTranQTY & ", " & _
                                             "tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
                                             "issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
                                             " where id = " & rsPartMas!ID
                        End If
                    Else
                        If vTDTranType = "ADJ" Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "onhand = ISNULL(onhand,0) - " & vTDTranQTY & ", " & _
                                             "tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
                                             "issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
                                             " where id = " & rsPartMas!ID
                        End If
                    End If
                    gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_dayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                End If

                If Null2String(rsTdayTran!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "Select id,recvd_code,ds1,status,classcode,rrno from PMIS_vw_RR_Trans where [TYPE] = '" & vTDType & "' AND rrno = '" & Format(vTDTranno, "000000") & "'", gconDMIS
                    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        vVatAmt = N2Str2Zero(rsRR_HD!ds1)
                        If Null2String(rsRR_HD!classcode) = "PCG" Or Null2String(rsRR_HD!classcode) = "PCS" Then
                            If CheckIfNonVatSup(Null2String(rsRR_HD!recvd_code)) = False Then
                               vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(vVatAmt)
                            Else
                                vTotTranCost = vTotTranInvAmt
                                gconDMIS.Execute ("update PMIS_TdayTran Set tranucost = " & N2Str2Zero(rsTdayTran!TRANINVAMT) & " Where id = " & rsTdayTran!ID)
                                gconDMIS.Execute ("update PMIS_dayTran Set tranucost = " & N2Str2Zero(rsTdayTran!TRANINVAMT) & " Where id = " & rsTdayTran!ID)
                                gconDMIS.Execute ("update PMIS_RR_HD Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                                gconDMIS.Execute ("update PMIS_REC_HIST Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                            End If
                        End If
                    Else
                        If vTDTranType = "ADJ" Or vTDTranType = "BEG" Then
                            vTotTranCost = vMAC * vTDTranQTY
                        End If
                    End If
                    If vPMOnhand <= 0 Then
                        vMAC = vTotTranCost / vTDTranQTY
                        gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                         "mac = " & vMAC & ", " & _
                                         "Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & ", " & _
                                         "ONORDER = ISNULL(ONORDER,0) - " & vTDTranQTY & ", " & _
                                         "SERVED = ISNULL(SERVED,0) + " & vTDTranQTY & ", " & _
                                         "last_recd = " & vTDTranDate & ", " & _
                                         "trecqty = ISNULL(trecqty,0) + " & vTDTranQTY & ", " & _
                                         "receipts = ISNULL(receipts,0) + " & vTDTranQTY & _
                                         " where id =" & rsPartMas!ID
                    Else
                        vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                        gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                         "mac = " & vMAC & ", " & _
                                         "ONORDER = ISNULL(ONORDER,0) - " & vTDTranQTY & ", " & _
                                         "SERVED = ISNULL(SERVED,0) + " & vTDTranQTY & ", " & _
                                         "Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & ", " & _
                                         "last_recd = " & vTDTranDate & ", " & _
                                         "trecqty = ISNULL(trecqty,0) + " & vTDTranQTY & ", " & _
                                         "receipts = ISNULL(receipts,0) + " & vTDTranQTY & _
                                         " where id =" & rsPartMas!ID
                    End If
                    gconDMIS.Execute "update PMIS_TdayTran set mac = " & vMAC & " where id = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_dayTran set mac = " & vMAC & " where id = " & rsTdayTran!ID
                End If
                If Null2String(rsTdayTran!TRANTYPE) = "PO" Then
                    gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                     "purchases = ISNULL(purchases,0) + " & vTDTranQTY & "," & _
                                     "tpoqty = ISNULL(tpoqty,0) + " & vTDTranQTY & "," & _
                                     "ONORDER = ISNULL(ONORDER,0) + " & vTDTranQTY & "," & _
                                     "ORDERED = ISNULL(ORDERED,0) + " & vTDTranQTY & _
                                     " where id = " & rsPartMas!ID
                End If
            End If
            I = I + 1
            progCPB.Value = (I / rsTdayTran.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            rsTdayTran.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents
        Screen.MousePointer = 0
    End If
    txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Updating of Stock Master Completed...": DoEvents
    MsgSpeechBox "Updating of Master File Completed..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISUpdateMaster = Nothing
    UnloadForm Me
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Function CheckIfNonVatSup(SupplierCode As String) As Boolean
    Dim rsSupplierMaster                As ADODB.Recordset
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

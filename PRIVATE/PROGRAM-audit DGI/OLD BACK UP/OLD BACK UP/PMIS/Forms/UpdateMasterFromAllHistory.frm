VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmPMISUpdateMasterFromAllHistory 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Master File From All History"
   ClientHeight    =   6240
   ClientLeft      =   270
   ClientTop       =   360
   ClientWidth     =   8940
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FF8080&
   Icon            =   "UpdateMasterFromAllHistory.frx":0000
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
         Enabled         =   -1  'True
         TextRTF         =   $"UpdateMasterFromAllHistory.frx":030A
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
      MouseIcon       =   "UpdateMasterFromAllHistory.frx":0388
      MousePointer    =   99  'Custom
      Picture         =   "UpdateMasterFromAllHistory.frx":04DA
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
      MouseIcon       =   "UpdateMasterFromAllHistory.frx":0840
      MousePointer    =   99  'Custom
      Picture         =   "UpdateMasterFromAllHistory.frx":0992
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
         Picture         =   "UpdateMasterFromAllHistory.frx":0C2D
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "UpdateMasterFromAllHistory.frx":0C49
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
      Picture         =   "UpdateMasterFromAllHistory.frx":0C65
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "UpdateMasterFromAllHistory.frx":0C81
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
      Picture         =   "UpdateMasterFromAllHistory.frx":0C9D
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "UpdateMasterFromAllHistory.frx":0CB9
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
Attribute VB_Name = "frmPMISUpdateMasterFromAllHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub UpdateMaster()
    Dim rsPartMas, rsCURPartmas, rsTdayTran                           As ADODB.Recordset
    Dim I                                                             As Long
    Dim vTotTranCost, vTotTranInvAmt, vTDTranQTY                      As Double
    Dim vTDTranDate, vTDTranType, vTDTranno, vTDType                  As String
    Dim vMAC                                                          As Double
    Dim vPMOnhand                                                     As Long
    Dim vSTOCKDESC                                                    As String

    Dim INCLUDE_UPDATE_MAC                                            As Boolean

    'If COMPANY_CODE = "HGC" Then
    '    gconDMIS.Execute ("Update PMIS_DayTran Set TRANUCOST = MAC, TRANINVAMT = MAC WHERE ROUND(TRANUCOST / TRANQTY,0) =  ROUND(MAC,0) AND TRANTYPE = 'BEG'")
    'End If
    If COMPANY_CODE = "HMH" Then
        gconDMIS.Execute ("Update PMIS_DayTran Set TRANDATE = '12/31/2007' WHERE TRANTYPE = 'BEG'")
    End If
    If COMPANY_CODE = "HGC" Then
        gconDMIS.Execute ("Update PMIS_DayTran Set TRANDATE = '4/30/2008' WHERE TRANTYPE = 'BEG'")
    End If
    
    gconDMIS.Execute ("delete from PMIS_StkStat")
    INCLUDE_UPDATE_MAC = True
    Dim Bulan As Integer
    
    Screen.MousePointer = 11
    txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Storing Stocks Master Beginning Balances...": DoEvents
    If INCLUDE_UPDATE_MAC = True Then
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " onhand = 0, mac=0"
    Else
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " onhand = 0"
    End If
    Dim Yir As Integer
    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                   " STOCKTYPE = 'GJ' WHERE (STOCKTYPE <> 'BP' AND LEFT(STOCKNO,2) <> '08')"
    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                   " NON_HARI = 'N' WHERE NON_HARI IS NULL"
    For Yir = 2007 To 2008
        For Bulan = 1 To 12
            Set rsTdayTran = New ADODB.Recordset
                rsTdayTran.Open "select id,ItemNo,trantype,tranno,TYPE,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIS_DayTran where (month(trandate) = " & Bulan & " and year(trandate) = " & Yir & ") and (trantype <> 'ADB' and trantype <> 'ARS' and trantype <> 'MRS' and trantype <> 'PRS') and (status = 'P' OR status = 'B') order by trandate asc, id asc,tranno asc", gconDMIS
            If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
                rsTdayTran.MoveFirst:
                txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Updating Stock Master from Transactions Made...": DoEvents
                I = 0
                Do While Not rsTdayTran.EOF
                    
                    gconDMIS.Execute "Update PMIS_DayTran set ItemNo = '" & Format(Null2String(rsTdayTran!itemno), "0000") & "' where ID = " & rsTdayTran!ID
                    vTDType = Null2String(rsTdayTran![Type])
                    vTDTranDate = N2Date2Null(rsTdayTran!trandate)
                    vTDTranType = Null2String(rsTdayTran!TranType)
                    vTDTranno = Null2String(rsTdayTran!TRANNO)
                    vTDTranQTY = N2Str2Zero(rsTdayTran!tranqty)
                    vTotTranCost = N2Str2Zero(rsTdayTran!TRANUCOST) * vTDTranQTY
                    vTotTranInvAmt = N2Str2Zero(rsTdayTran!TRANINVAMT) * vTDTranQTY
                    labProcessing.Caption = "Processing: " & Null2String(rsTdayTran!TranType) & " #" & Null2String(rsTdayTran!TRANNO): DoEvents
        
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
                    Set rsPartMas = gconDMIS.Execute("select id,STOCKNO,mac,Onhand,NON_HARI from PMIS_STOCKMAS where TYPE = '" & vTDType & "' AND STOCKNO = " & N2Str2Null(rsTdayTran!STOCK_ORD))
                    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                        vMAC = N2Str2Zero(rsPartMas!Mac)
                        If vMAC <= 0 Then vMAC = N2Str2Zero(rsTdayTran!TRANUCOST)
                        vPMOnhand = N2Str2IntZero(rsPartMas!ONHAND)
                        If Null2String(rsTdayTran!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                             "onhand = ISNULL(onhand,0) - " & vTDTranQTY & _
                                           " where id = " & rsPartMas!ID
                            If INCLUDE_UPDATE_MAC = True Then
                                gconDMIS.Execute "update PMIS_DayTran set NON_HARI = " & N2Str2Null(rsPartMas!NON_HARI) & ", tranucost = " & vMAC & ", MAC = " & vMAC & " where ID = " & rsTdayTran!ID
                            Else
                                gconDMIS.Execute "update PMIS_DayTran set NON_HARI = " & N2Str2Null(rsPartMas!NON_HARI) & " where ID = " & rsTdayTran!ID
                            End If
                        End If
        
                        If Null2String(rsTdayTran!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                            If INCLUDE_UPDATE_MAC = True Then
                                If vPMOnhand <= 0 Then
                                    vMAC = vTotTranCost / vTDTranQTY
                                    gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                                     "mac = " & vMAC & ", " & _
                                                     "Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & _
                                                   " where id =" & rsPartMas!ID
                                Else
                                    vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                                    gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                                     "mac = " & vMAC & ", " & _
                                                     "Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & _
                                                   " where id =" & rsPartMas!ID
                                End If
                                gconDMIS.Execute "update PMIS_DayTran set NON_HARI = " & N2Str2Null(rsPartMas!NON_HARI) & ", mac = " & vMAC & ", Tranucost = " & N2Str2Zero(rsTdayTran!TRANUCOST) & " where id = " & rsTdayTran!ID
                            Else
                                If vPMOnhand <= 0 Then
                                    gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                                     "Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & _
                                                   " where id =" & rsPartMas!ID
                                Else
                                    gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                                     "Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & _
                                                   " where id =" & rsPartMas!ID
                                End If
                                gconDMIS.Execute "update PMIS_DayTran set NON_HARI = " & N2Str2Null(rsPartMas!NON_HARI) & ", Tranucost = " & N2Str2Zero(rsTdayTran!TRANUCOST) & " where id = " & rsTdayTran!ID
                            End If
                        End If
                    End If
                    I = I + 1
                    progCPB.Value = (I / rsTdayTran.RecordCount) * 100
                    labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
                    rsTdayTran.MoveNext
                Loop
                labProcessing.Caption = "": DoEvents
                Screen.MousePointer = 0
                If INCLUDE_UPDATE_MAC = True Then
                    gconDMIS.Execute "update PMIS_PARTMAS set" & _
                                   " PMIS_PARTMAS.lastm_oh = ISNULL(PMIS_PARTMAS.onhand,0)," & _
                                   " PMIS_PARTMAS.lastm_mac = ISNULL(PMIS_PARTMAS.Mac,0)," & _
                                   " PMIS_PARTMAS.lastm_mad = ISNULL(PMIS_PARTMAS.Mad,0)," & _
                                   " PMIS_PARTMAS.lastm_oo = ISNULL(PMIS_PARTMAS.onorder,0)" & _
                                   " where PMIS_PARTMAS.ACTIVE = 'Y'"
                    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                                   " PMIS_STOCKMAS.lastm_oh = ISNULL(PMIS_STOCKMAS.onhand,0)," & _
                                   " PMIS_STOCKMAS.lastm_mac = ISNULL(PMIS_STOCKMAS.Mac,0)," & _
                                   " PMIS_STOCKMAS.lastm_mad = ISNULL(PMIS_STOCKMAS.Mad,0)," & _
                                   " PMIS_STOCKMAS.lastm_oo = ISNULL(PMIS_STOCKMAS.onorder,0)" & _
                                   " where PMIS_STOCKMAS.TYPE = 'M' AND PMIS_STOCKMAS.ACTIVE = 'Y'"
                    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                                   " PMIS_STOCKMAS.lastm_oh = ISNULL(PMIS_STOCKMAS.onhand,0)," & _
                                   " PMIS_STOCKMAS.lastm_mac = ISNULL(PMIS_STOCKMAS.Mac,0)," & _
                                   " PMIS_STOCKMAS.lastm_mad = ISNULL(PMIS_STOCKMAS.Mad,0)," & _
                                   " PMIS_STOCKMAS.lastm_oo = ISNULL(PMIS_STOCKMAS.onorder,0)" & _
                                   " where PMIS_STOCKMAS.TYPE = 'A' AND PMIS_STOCKMAS.ACTIVE = 'Y'"
                Else
                    gconDMIS.Execute "update PMIS_PARTMAS set" & _
                                   " PMIS_PARTMAS.lastm_oh = ISNULL(PMIS_PARTMAS.onhand,0)," & _
                                   " PMIS_PARTMAS.lastm_mad = ISNULL(PMIS_PARTMAS.Mad,0)," & _
                                   " PMIS_PARTMAS.lastm_oo = ISNULL(PMIS_PARTMAS.onorder,0)" & _
                                   " where PMIS_PARTMAS.ACTIVE = 'Y'"
                    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                                   " PMIS_STOCKMAS.lastm_oh = ISNULL(PMIS_STOCKMAS.onhand,0)," & _
                                   " PMIS_STOCKMAS.lastm_mad = ISNULL(PMIS_STOCKMAS.Mad,0)," & _
                                   " PMIS_STOCKMAS.lastm_oo = ISNULL(PMIS_STOCKMAS.onorder,0)" & _
                                   " where PMIS_STOCKMAS.TYPE = 'M' AND PMIS_STOCKMAS.ACTIVE = 'Y'"
                    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                                   " PMIS_STOCKMAS.lastm_oh = ISNULL(PMIS_STOCKMAS.onhand,0)," & _
                                   " PMIS_STOCKMAS.lastm_mad = ISNULL(PMIS_STOCKMAS.Mad,0)," & _
                                   " PMIS_STOCKMAS.lastm_oo = ISNULL(PMIS_STOCKMAS.onorder,0)" & _
                                   " where PMIS_STOCKMAS.TYPE = 'A' AND PMIS_STOCKMAS.ACTIVE = 'Y'"
                End If
                txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Updating of Stock Master Completed...": DoEvents
                CreateStockStatus Bulan, Yir
                txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Create Stock Status Completed...": DoEvents
            End If
        Next
    Next
    MsgSpeechBox "Updating of Master File Completed..."
End Sub

Sub CreateStockStatus(xxxManth As Integer, xxxYeer As Integer)
    Screen.MousePointer = 11
    Dim DATE_GEN As String
    progCPB.Value = 0
    Me.Caption = "Updating Part Master File"
    labCPB.Caption = "Updating Stocks Master File for Stock Status... Please Wait..."
    DoEvents
    progCPB.Value = 100
    DATE_GEN = lastDay(DateSerial(xxxYeer, xxxManth, 1))
    DoEvents: Screen.MousePointer = 11: progCPB.Value = 0: Me.Caption = "Creating Stock Status"
    labCPB.Caption = "Create Stock Status Master File... Please Wait...": DoEvents: progCPB.Value = 100
    gconDMIS.Execute "insert into PMIS_StkStat " & _
                     "(TYPE, STOCKTYPE, NON_HARI, STOCKNO,STOCKDESC,onhand,mac,mad,sstock,resservice,onorder,ADJ_ADD,ADJ_MINUS,BACKORD,SOQ,SRP,TD,EM_PO,LS,LOS)" & _
                   " select TYPE, STOCKTYPE, NON_HARI, STOCKNO,STOCKDESC,ISNULL(OnHand,0),ISNULL(Mac,0),ISNULL(Mad,0),ISNULL(SStock,0),ISNULL(ResService,0),ISNULL(OnOrder,0),ISNULL(TADJQTY_IN,0),ISNULL(TADJQTY_OUT,0),ISNULL(BACKORDER,0),ISNULL(SOQ,0),ISNULL(SRP,0),(ISNULL(ONREQUEST,0) + ISNULL(S_ONREQUEST,0)),ISNULL(EMERGENCY_PO,0),ISNULL(LOST_SALES,0),ISNULL(LEVEL_OF_SERVICE,0) from PMIS_STOCKMAS WHERE ACTIVE = 'Y' order by STOCKNO asc"
    gconDMIS.Execute "update PMIS_StkStat set date_gen = " & N2Date2Null(DATE_GEN) & " where date_gen IS NULL"
    MsgSpeech "Create Stock Status Complete!"
    Screen.MousePointer = 0
    DoEvents
End Sub

Private Sub cmdCheck_Click()
    'If Function_Access(LOGID, "Acess_Process", "UPDATE MASTER FILE") = False Then Exit Sub
    cmdCheck.Enabled = False
    cmdExit.Enabled = False
    DoEvents
    UpdateMaster
    cmdExit.Enabled = True
    
    DoEvents
    
    NEW_LogAudit "R", "UPDATE MASTER FILE", "", "", "", " - ", "", ""
    
    Exit Sub
ERRORCODE:
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
    Dim rsTdayTran                                                    As ADODB.Recordset
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select trantype from PMIS_DayTran where trantype = 'ADJ'", gconDMIS
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        chkUpdateAdjustment.Enabled = True
        chkUpdateAdjustment.Value = 1
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISProcess_UpdateMaster = Nothing
    UnloadForm Me
End Sub

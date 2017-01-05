VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Begin VB.Form frmCSMSMatUpdateMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Update Master File"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MatUpdateMaster.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6090
   Begin VB.CheckBox chkUpdateAdjustment 
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
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1500
      Value           =   1  'Checked
      Width           =   5685
   End
   Begin MSMask.MaskEdBox txtFrom 
      Height          =   345
      Left            =   750
      TabIndex        =   0
      Top             =   690
      Width           =   1965
      _ExtentX        =   3466
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
      Left            =   750
      TabIndex        =   1
      Top             =   1080
      Width           =   1965
      _ExtentX        =   3466
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
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   330
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   556
      Picture         =   "MatUpdateMaster.frx":01CA
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "MatUpdateMaster.frx":01E6
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
      Left            =   5220
      MouseIcon       =   "MatUpdateMaster.frx":0202
      MousePointer    =   99  'Custom
      Picture         =   "MatUpdateMaster.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit Window"
      Top             =   720
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
      Left            =   4500
      MouseIcon       =   "MatUpdateMaster.frx":06BA
      MousePointer    =   99  'Custom
      Picture         =   "MatUpdateMaster.frx":080C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Update Materials Master File"
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Left            =   -60
      TabIndex        =   4
      Top             =   1110
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   -60
      TabIndex        =   3
      Top             =   720
      Width           =   765
   End
   Begin VB.Label labCPB 
      BackColor       =   &H8000000D&
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   5835
   End
End
Attribute VB_Name = "frmCSMSMatUpdateMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTDAYTRAN                                         As ADODB.Recordset
Private Sub cmdCheck_Click()
    If Function_Access(LOGID, "Acess_Process", "UPDATE MASTERFILE") = False Then Exit Sub
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select trandate from CSMS_TdayTran where trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "'", gconDMIS
    If rsTDAYTRAN.EOF And rsTDAYTRAN.BOF Then
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
    DoEvents
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    txtFrom.Text = firstDay(LOGDATE)
    txtTo.Text = LOGDATE
    Me.Height = 1965
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select trantype from CSMS_TdayTran where trantype = 'ADJ'", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        Me.Height = 2250
        chkUpdateAdjustment.Enabled = True
    End If
    Screen.MousePointer = 0
End Sub

Sub UpdateMaster()
    Dim rsMatMas, rsTDAYTRAN, rsMATREC                 As ADODB.Recordset
    Dim i                                              As Integer
    Dim vTotTranCost, vTDTranQTY                       As Double
    Dim vTDTranType, vTDTranno, vSupCode               As String
    Dim vVatAmt, vCOST                                 As Double
    Dim vPMOnhand                                      As Integer

    If chkUpdateAdjustment.Value = 1 Then
        Set rsTDAYTRAN = New ADODB.Recordset
        rsTDAYTRAN.Open "select id,ItemNo,trantype,tranno,MatCde,tranqty,status,in_out,tranucost from CSMS_TdayTran where status <> 'C' AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by id asc", gconDMIS
        If rsTDAYTRAN.EOF And rsTDAYTRAN.BOF Then
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select id,ItemNo,trantype,tranno,MatCde,tranqty,status,in_out,tranucost from CSMS_DayTran where status <> 'C' AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by id asc", gconDMIS
        End If
    Else
        Set rsTDAYTRAN = New ADODB.Recordset
        rsTDAYTRAN.Open "select id,ItemNo,trantype,tranno,MatCde,tranqty,status,in_out,tranucost from CSMS_TdayTran where trantype <> 'ADJ' and status <> 'C' AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by id asc", gconDMIS
        If rsTDAYTRAN.EOF And rsTDAYTRAN.BOF Then
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select id,ItemNo,trantype,tranno,MatCde,tranqty,status,in_out,tranucost from CSMS_DayTran where trantype <> 'ADJ' and status <> 'C' AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by id asc", gconDMIS
        End If
    End If
    'Update Master File based on History Of Transactions
    'Set rsTdaytran = New ADODB.Recordset
    '    rsTdaytran.Open "select id,ItemNo,trantype,tranno,MatCde,tranqty,status,in_out,tranucost from CSMS_DayTran where status <> 'C' order by id asc", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        rsTDAYTRAN.MoveFirst
        Screen.MousePointer = 11
        DoEvents
        Me.Caption = "Updating Materials Master File"
        'Update Master File based on History Of Transactions
        gconDMIS.Execute "update CSMS_MatMas set onhand = MatMas.lastm_oh" & _
                         ", Cost = MatMas.lastm_mac, onorder = MatMas.lastm_oo" & _
                         ", tissqty = 0, trecqty = 0, tpoqty = 0, receipts = 0" & _
                         ", issuances = 0"
        'gconDMIS.Execute "update CSMS_MatMas set onhand = 0" & _
         '                  ", Cost = 0, onorder = 0" & _
         '                  ", tissqty = 0, trecqty = 0, tpoqty = 0, receipts = 0" & _
         '                  ", issuances = 0"
        'gconDMIS.Execute "update CSMS_MatMas set MatMas.lastm_oh=onhand" & _
         '                  ", MatMas.lastm_mac=Cost, MatMas.lastm_oo=onorder" & _
         '                  ", tissqty = 0, trecqty = 0, tpoqty = 0, receipts = 0" & _
         '                  ", issuances = 0"
        DoEvents
        MsgSpeech "Updating Transactions to Materials Master File"
        Me.Caption = "Updating Transactions to Materials Master File"
        DoEvents
        i = 0
        Do While Not rsTDAYTRAN.EOF
            vTDTranType = Null2String(rsTDAYTRAN!TRANTYPE)
            vTDTranno = Null2String(rsTDAYTRAN!Tranno)
            vTDTranQTY = N2Str2IntZero(rsTDAYTRAN!tranqty)
            vTotTranCost = N2Str2Zero(rsTDAYTRAN!TRANUCOST) * vTDTranQTY
            labCPB.Caption = "Processing: " & Null2String(rsTDAYTRAN!TRANTYPE) & " #" & Null2String(rsTDAYTRAN!Tranno)
            DoEvents
            gconDMIS.Execute "update CSMS_TdayTran set ItemNo = '" & Format(Null2String(rsTDAYTRAN!itemno), "0000") & "' where id = " & rsTDAYTRAN!ID
            'Update Master File based on History Of Transactions
            'gconDMIS.Execute "update CSMS_DayTran set ItemNo = '" & Format(Null2String(rsTdaytran!itemno), "0000") & "' where id = " & rsTdaytran!ID
            Set rsMatMas = New ADODB.Recordset
            rsMatMas.Open "select id,MatCde,Cost,tissqty,trecqty,tpoqty,tprqty,prqty,issuances,receipts,onhand from CSMS_MatMas where MatCde = " & N2Str2Null(rsTDAYTRAN!MATCDE), gconDMIS
            If Not rsMatMas.EOF And Not rsMatMas.BOF Then
                vCOST = N2Str2Zero(rsMatMas!COST)
                vPMOnhand = N2Str2IntZero(rsMatMas!ONHAND)
                If Null2String(rsTDAYTRAN!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    'If vTDTranType = "ADJ" And Null2String(rsTdaytran!Status) = "P" Then
                    'Else
                    gconDMIS.Execute "update CSMS_MatMas set " & _
                                     "onhand = " & vPMOnhand - vTDTranQTY & ", " & _
                                     "tissqty = " & N2Str2IntZero(rsMatMas!TISSQTY) + vTDTranQTY & ", " & _
                                     "issuances = " & N2Str2IntZero(rsMatMas!issuances) + vTDTranQTY & _
                                   " where id = " & rsMatMas!ID
                    gconDMIS.Execute "update CSMS_TdayTran set tranucost = " & vCOST & " where ID = " & rsTDAYTRAN!ID
                    'End If
                End If

                If Null2String(rsTDAYTRAN!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    Set rsMATREC = New ADODB.Recordset
                    rsMATREC.Open "Select id,recvd_code,ds1,status,classcode,rrno from CSMS_MatRec where rrno = '" & Format(vTDTranno, "000000") & "' and status <> 'C'", gconDMIS
                    If Not rsMATREC.EOF And Not rsMATREC.BOF Then
                        vSupCode = Null2String(rsMATREC!recvd_code)
                        vVatAmt = N2Str2Zero(rsMATREC!ds1)
                        If Null2String(rsMATREC!classcode) = "PCG" Or Null2String(rsMATREC!classcode) = "PCS" Then
                            gconDMIS.Execute "update CSMS_MatRec set ds1 = 10 where id = " & rsMATREC!ID
                            vVatAmt = 10
                            If vSupCode <> vPAMCOR And vVatAmt <= 0 Then
                                vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(VAT_RATE)
                            End If
                        End If
                    Else
                        If vTDTranType = "ADJ" Then
                            vTotTranCost = vCOST * vTDTranQTY
                        End If
                    End If

                    'If vTDTranType = "ADJ" And Null2String(rsTdaytran!Status) = "P" Then
                    'Else
                    If vPMOnhand <= 0 Then
                        gconDMIS.Execute "update CSMS_MatMas set " & _
                                         "Cost = " & vTotTranCost / vTDTranQTY & ", " & _
                                         "Onhand = " & vPMOnhand + vTDTranQTY & ", " & _
                                         "trecqty = " & N2Str2IntZero(rsMatMas!trecqty) + vTDTranQTY & ", " & _
                                         "receipts = " & N2Str2IntZero(rsMatMas!receipts) + vTDTranQTY & _
                                       " where id =" & rsMatMas!ID
                    Else
                        gconDMIS.Execute "update CSMS_MatMas set " & _
                                         "Cost = " & ((vCOST * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand) & ", " & _
                                         "Onhand = " & vPMOnhand + vTDTranQTY & ", " & _
                                         "trecqty = " & N2Str2IntZero(rsMatMas!trecqty) + vTDTranQTY & ", " & _
                                         "receipts = " & N2Str2IntZero(rsMatMas!receipts) + vTDTranQTY & _
                                       " where id =" & rsMatMas!ID
                    End If
                    'End If
                End If

                If Null2String(rsTDAYTRAN!TRANTYPE) = "PO" Then
                    gconDMIS.Execute "update CSMS_MatMas set " & _
                                     "tpoqty = " & N2Str2IntZero(rsMatMas!tpoqty) + vTDTranQTY & _
                                   " where id = " & rsMatMas!ID
                End If
            End If
            gconDMIS.Execute "update CSMS_TdayTran set status = 'N' where ID = " & rsTDAYTRAN!ID
            'Update Master File based on History Of Transactions
            'gconDMIS.Execute "update CSMS_TdayTran set status = 'P' where ID = " & rsTdaytran!ID
            DoEvents
            i = i + 1
            progCPB.Value = (i / rsTDAYTRAN.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTDAYTRAN.MoveNext
        Loop
        labCPB.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "Error Opening Tdaytran File"
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

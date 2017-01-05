VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Begin VB.Form frmOSMSProcessUpdateMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Supplies Master File"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H8000000F&
   Icon            =   "UpdateMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   5775
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
      MouseIcon       =   "UpdateMaster.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "UpdateMaster.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check"
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
      MouseIcon       =   "UpdateMaster.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "UpdateMaster.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.PictureBox picCPB 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   0
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
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "UpdateMaster.frx":0BAF
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "UpdateMaster.frx":0BCB
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
            MICON           =   "UpdateMaster.frx":0BE7
         End
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
Attribute VB_Name = "frmOSMSProcessUpdateMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCheck_Click()
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Height = 1965
    'Set rsISSUANCE_DETAILS = New ADODB.Recordset
    '    rsISSUANCE_DETAILS.Open "select trantype from OSMS_ISSUANCE_DETAILS where trantype = 'ADJ'", gconDMIS
    'If Not rsISSUANCE_DETAILS.EOF And Not rsISSUANCE_DETAILS.BOF Then
    '   Me.Height = 2250
    '   chkUpdateAdjustment.Enabled = True
    'End If
    'txtFrom.Text = firstDay(LOGDATE)
    'txtTo.Text = LOGDATE
    Screen.MousePointer = 0
End Sub

Sub UpdateMaster()
    Dim rsSupply As ADODB.Recordset
    Dim rsISSUANCE_DETAILS As ADODB.Recordset
    Dim rsrrDETAILS As ADODB.Recordset
    Dim i As Integer
    Dim vTotTranCost As Double
    Dim vTDTranQTY As Double
    Dim vTDTRANS_NO As String
    Dim vMAC As Double
    Dim vSUP_ONHAND As Integer

    Me.Caption = "Updating Supply Master File"
    'gconDMIS.Execute "update OSMS_SUPPLY set onhand = SUPPLY.lastm_oh" & _
     '                  ", COST = SUPPLY.LASTM_COST, onorder = SUPPLY.lastm_oo" & _
     '                  ", tissqty = 0, trecqty = 0, tpoqty = 0, receipts = 0" & _
     '                  ", issuances = 0"
    gconDMIS.Execute "update OSMS_SUPPLY set onhand = 0" & _
                     ", tissqty = 0, trecqty = 0, tpoqty = 0, receipts = 0" & _
                     ", issuances = 0"
    DoEvents
    Set rsrrDETAILS = New ADODB.Recordset
    rsrrDETAILS.Open "select id,ITEM_NO,RRNumber,SUPPLY_CODE,rrQUANTITY,status,COST from OSMS_RRDETAILS  where status <> 'C' order by rrnumber,item_no asc", gconDMIS
    If Not rsrrDETAILS.EOF And Not rsrrDETAILS.BOF Then
        rsrrDETAILS.MoveFirst
        Screen.MousePointer = 11
        DoEvents
        Me.Caption = "Updating Receipts Transactions to Supply Master File"
        DoEvents
        i = 0
        Do While Not rsrrDETAILS.EOF
            gconDMIS.Execute "update OSMS_RRDETAILS  set ITEM_NO = '" & Format(Null2String(rsrrDETAILS!item_no), "0000") & "' where ID = " & rsrrDETAILS!Id
            vTDTRANS_NO = Null2String(rsrrDETAILS!rrnumber)
            vTDTranQTY = N2Str2IntZero(rsrrDETAILS!rrQUANTITY)
            vTotTranCost = N2Str2Zero(rsrrDETAILS!Cost) * vTDTranQTY
            labProcessing.Caption = "Processing: RECEIPTS #" & Null2String(rsrrDETAILS!rrnumber)
            DoEvents
            Set rsSupply = New ADODB.Recordset
            rsSupply.Open "select id,SUPPLY_CODE,COST,onhand from OSMS_SUPPLY where SUPPLY_CODE = " & N2Str2Null(rsrrDETAILS!Supply_Code), gconDMIS
            If Not rsSupply.EOF And Not rsSupply.BOF Then
                vMAC = N2Str2Zero(rsSupply!Cost)
                vSUP_ONHAND = N2Str2IntZero(rsSupply!Onhand)
                gconDMIS.Execute "update OSMS_SUPPLY set " & _
                                 "cost = " & vMAC & ", " & _
                                 "onhand = onhand + " & vTDTranQTY & ", " & _
                                 "trecqty = trecqty + " & vTDTranQTY & ", " & _
                                 "receipts = receipts + " & vTDTranQTY & _
                               " where id = " & rsSupply!Id
                gconDMIS.Execute "update OSMS_RRDETAILS  set cost = " & vMAC & " where ID = " & rsrrDETAILS!Id
            End If
            DoEvents
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
    Set rsISSUANCE_DETAILS = New ADODB.Recordset
    rsISSUANCE_DETAILS.Open "select id,ID_ITEM_NO,trans_no,SUPPLY_CODE,ID_QUANTITY,status,COST from OSMS_ISSUANCE_DETAILS where status <> 'C' order by trans_no,id_item_no asc", gconDMIS
    If Not rsISSUANCE_DETAILS.EOF And Not rsISSUANCE_DETAILS.BOF Then
        rsISSUANCE_DETAILS.MoveFirst
        Screen.MousePointer = 11
        DoEvents
        Me.Caption = "Updating Issuance Transactions to Supply Master File"
        DoEvents
        i = 0
        Do While Not rsISSUANCE_DETAILS.EOF
            gconDMIS.Execute "update OSMS_ISSUANCE_DETAILS set ID_ITEM_NO = '" & Format(Null2String(rsISSUANCE_DETAILS!id_item_no), "0000") & "' where ID = " & rsISSUANCE_DETAILS!Id
            vTDTRANS_NO = Null2String(rsISSUANCE_DETAILS!Trans_No)
            vTDTranQTY = N2Str2IntZero(rsISSUANCE_DETAILS!ID_Quantity)
            vTotTranCost = N2Str2Zero(rsISSUANCE_DETAILS!Cost) * vTDTranQTY
            labProcessing.Caption = "Processing: ISSUANCE #" & Null2String(rsISSUANCE_DETAILS!Trans_No)
            DoEvents
            Set rsSupply = New ADODB.Recordset
            rsSupply.Open "select id,SUPPLY_CODE,COST,onhand from OSMS_SUPPLY where SUPPLY_CODE = " & N2Str2Null(rsISSUANCE_DETAILS!Supply_Code), gconDMIS
            If Not rsSupply.EOF And Not rsSupply.BOF Then
                vMAC = N2Str2Zero(rsSupply!Cost)
                vSUP_ONHAND = N2Str2IntZero(rsSupply!Onhand)
                gconDMIS.Execute "update OSMS_SUPPLY set " & _
                                 "onhand = onhand - " & vTDTranQTY & ", " & _
                                 "tissqty = tissqty + " & vTDTranQTY & ", " & _
                                 "issuances = issuances + " & vTDTranQTY & _
                               " where id = " & rsSupply!Id
                gconDMIS.Execute "update OSMS_ISSUANCE_DETAILS set cost = " & vMAC & " where ID = " & rsISSUANCE_DETAILS!Id
            End If
            DoEvents
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

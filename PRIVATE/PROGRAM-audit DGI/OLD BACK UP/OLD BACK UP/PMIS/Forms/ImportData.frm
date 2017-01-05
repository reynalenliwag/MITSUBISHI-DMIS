VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmPMIOSImportData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATA TRANSFER FROM UNFORMATTED TO STANDARD FORMAT"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   ForeColor       =   &H8000000F&
   Icon            =   "ImportData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "ImportData.frx":27A2
   ScaleHeight     =   5220
   ScaleWidth      =   5745
   Begin VB.TextBox txtMissingPartNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1500
      Width           =   5535
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   765
      Left            =   4740
      Picture         =   "ImportData.frx":54DE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   660
      Width           =   915
   End
   Begin VB.CommandButton cmdPost 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Import"
      Height          =   765
      Left            =   3840
      Picture         =   "ImportData.frx":57E8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   660
      Width           =   915
   End
   Begin VB.PictureBox picCPB 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   0
      Picture         =   "ImportData.frx":5AF2
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   2
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
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "ImportData.frx":882E
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "ImportData.frx":884A
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
         Picture         =   "ImportData.frx":8866
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
            MICON           =   "ImportData.frx":B5A2
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
         TabIndex        =   8
         Top             =   30
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frmPMIOSImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gconOldPMIS As ADODB.Connection

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPost_Click()
cmdPOST.Enabled = False
cmdExit.Enabled = False
UpdateLastM_MAC
'TRANSFER_DATA
'On Error Resume Next
'gconOldPMIS.Close
cmdExit.Enabled = True
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
Screen.MousePointer = 0
End Sub

Private Function OpenOldDb() As Boolean
Dim OLDPMIOS_Connection As String
With wizVar
     If .VerifyCryptoFile(App.Path & "\PMIOS.crp") = True Then
        OLDPMIOS_Connection = .OpenCryptoFile("OLDPMIOS", "CONNECT")
     End If
End With
On Error Resume Next
deOldPMIOS.deConnOldPMIOS.Close
On Error GoTo ConnErr
If OLDPMIOS_Connection <> "" Then
   deOldPMIOS.deConnOldPMIOS.ConnectionString = OLDPMIOS_Connection
   Set gconOldPMIS = New ADODB.Connection
   Set gconOldPMIS = deOldPMIOS.deConnOldPMIOS
   gconOldPMIS.Open
   OpenOldDb = True
Else
   OpenOldDb = False
End If
Exit Function

ConnErr:
  MsgBox Err.Description
End Function

Sub TRANSFER_DATA()
If OpenOldDb Then
   MoveTdaytran
   MoveTdaytranHist
   MoveOrdHd
   MoveOrdHdHist
   MoveRRhd
   MoveRRhdHist
   MovePOhd
   MovePOhdHist
   MovePartmas
   MoveShipping
   MoveRankFle
   MoveSupplier
   MoveCustomer
   MoveCunter
   MoveSalesman
   MoveCustCtl
   MoveLocation
Else
   MsgBoxXP "Cannot Find Old Database File"
End If
End Sub

Sub MoveTdaytran()
Dim MoveSql As String
Dim i As Integer

Dim varTRANDATE As String
Dim varTRANTYPE As String
Dim varTRANNO As String
Dim varITEMNO As String
Dim varPART_ORD As String
Dim varPART_SUP As String
Dim varTRANQTY As Integer
Dim varUNIT As String
Dim varTRANUCOST As Double
Dim varTRANUPRICE As Double
Dim varNETCOST As Double
Dim varNETPRICE As Double
Dim varSTATUS As String
Dim varIN_OUT As String
Dim varMATCH As String
Dim varLISTED As String
Dim varMAC As Double
Dim varTRANINVAMT As Double
Dim varUSERCODE As String
Dim varLASTUPDATE As String
Dim varTREMARKS As String
Dim rsOldTdaytran As ADODB.Recordset
gconPMIOS.Execute "delete from tdaytran"
Set rsOldTdaytran = New ADODB.Recordset
    rsOldTdaytran.Open "select * from tdaytran", gconOldPMIS
If Not rsOldTdaytran.EOF And Not rsOldTdaytran.BOF Then
   rsOldTdaytran.MoveFirst
   Me.Caption = "Currently Converting Tdaytran File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldTdaytran.EOF
      varTRANDATE = N2Str2Null(rsOldTdaytran!trandate)
      varTRANTYPE = N2Str2Null(rsOldTdaytran!trantype)
      varTRANNO = N2Str2Null(rsOldTdaytran!tranno)
      varITEMNO = N2Str2Null(rsOldTdaytran!itemno)
      varPART_ORD = N2Str2Null(rsOldTdaytran!part_ord)
      varPART_SUP = N2Str2Null(rsOldTdaytran!part_sup)
      varTRANQTY = N2Str2Zero(rsOldTdaytran!tranqty)
      varUNIT = N2Str2Null(rsOldTdaytran!unit)
      varTRANUCOST = N2Str2Zero(rsOldTdaytran!tranucost)
      varTRANUPRICE = N2Str2Zero(rsOldTdaytran!tranuprice)
      varNETCOST = N2Str2Zero(rsOldTdaytran!netcost)
      varNETPRICE = N2Str2Zero(rsOldTdaytran!NETprice)
      varSTATUS = N2Str2Null(rsOldTdaytran!Status)
      varIN_OUT = N2Str2Null(rsOldTdaytran!in_out)
      varMATCH = N2Str2Null(rsOldTdaytran!MATCH)
      varLISTED = N2Str2Null(rsOldTdaytran!listed)
      varMAC = N2Str2Zero(rsOldTdaytran!MAC)
      varTRANINVAMT = N2Str2Zero(rsOldTdaytran!traninvamt)
      varUSERCODE = N2Str2Null(rsOldTdaytran!usercode)
      varLASTUPDATE = N2Str2Null(rsOldTdaytran!lastupdate)
   
      MoveSql = "INSERT INTO TDAYTRAN " & _
                "(TRANDATE,TRANTYPE,TRANNO,ITEMNO,PART_ORD,PART_SUP,TRANQTY,UNIT,TRANUCOST,TRANUPRICE,NETCOST,NETPRICE,STATUS,IN_OUT,LISTED,MAC,TRANINVAMT,USERCODE,LASTUPDATE)" & _
                " values (" & varTRANDATE & "," & varTRANTYPE & "," & varTRANNO & "," & varITEMNO & "," & varPART_ORD & "," & varPART_SUP & "," & varTRANQTY & "," & varUNIT & "," & varTRANUCOST & "," & varTRANUPRICE & "," & varNETCOST & "," & varNETPRICE & "," & varSTATUS & "," & varIN_OUT & "," & varLISTED & "," & varMAC & "," & varTRANINVAMT & "," & varUSERCODE & "," & varLASTUPDATE & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldTdaytran.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldTdaytran.MoveNext
   Loop
   Me.Caption = "Tdaytran File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub MoveTdaytranHist()
Dim MoveSql As String
Dim i As Integer

Dim varTRANDATE As String
Dim varTRANTYPE As String
Dim varTRANNO As String
Dim varITEMNO As String
Dim varPART_ORD As String
Dim varPART_SUP As String
Dim varTRANQTY As Integer
Dim varUNIT As String
Dim varTRANUCOST As Double
Dim varTRANUPRICE As Double
Dim varNETCOST As Double
Dim varNETPRICE As Double
Dim varSTATUS As String
Dim varIN_OUT As String
Dim varMATCH As String
Dim varLISTED As String
Dim varMAC As Double
Dim varTRANINVAMT As Double
Dim varUSERCODE As String
Dim varLASTUPDATE As String
Dim varTREMARKS As String
Dim rsOldTdaytranHist As ADODB.Recordset
gconPMIOS.Execute "delete from daytran"
Set rsOldTdaytranHist = New ADODB.Recordset
    rsOldTdaytranHist.Open "select * from daytran", gconOldPMIS
If Not rsOldTdaytranHist.EOF And Not rsOldTdaytranHist.BOF Then
   rsOldTdaytranHist.MoveFirst
   Me.Caption = "Currently Converting Tdaytran History File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldTdaytranHist.EOF
      varTRANDATE = N2Str2Null(rsOldTdaytranHist!trandate)
      varTRANTYPE = N2Str2Null(rsOldTdaytranHist!trantype)
      varTRANNO = N2Str2Null(rsOldTdaytranHist!tranno)
      varITEMNO = N2Str2Null(rsOldTdaytranHist!itemno)
      varPART_ORD = N2Str2Null(rsOldTdaytranHist!part_ord)
      varPART_SUP = N2Str2Null(rsOldTdaytranHist!part_sup)
      varTRANQTY = N2Str2IntZero(rsOldTdaytranHist!tranqty)
      varUNIT = N2Str2Null(rsOldTdaytranHist!unit)
      varTRANUCOST = N2Str2Zero(rsOldTdaytranHist!tranucost)
      varTRANUPRICE = N2Str2Zero(rsOldTdaytranHist!tranuprice)
      varNETCOST = N2Str2Zero(rsOldTdaytranHist!netcost)
      varNETPRICE = N2Str2Zero(rsOldTdaytranHist!NETprice)
      varSTATUS = N2Str2Null(rsOldTdaytranHist!Status)
      varIN_OUT = N2Str2Null(rsOldTdaytranHist!in_out)
      varMATCH = N2Str2Null(rsOldTdaytranHist!MATCH)
      varLISTED = N2Str2Null(rsOldTdaytranHist!listed)
      varMAC = N2Str2Zero(rsOldTdaytranHist!MAC)
      varTRANINVAMT = N2Str2Zero(rsOldTdaytranHist!traninvamt)
      varUSERCODE = N2Str2Null(rsOldTdaytranHist!usercode)
      varLASTUPDATE = N2Str2Null(rsOldTdaytranHist!lastupdate)
   
      MoveSql = "INSERT INTO DAYTRAN " & _
                "(TRANDATE,TRANTYPE,TRANNO,ITEMNO,PART_ORD,PART_SUP,TRANQTY,UNIT,TRANUCOST,TRANUPRICE,NETCOST,NETPRICE,STATUS,IN_OUT,LISTED,MAC,TRANINVAMT,USERCODE,LASTUPDATE)" & _
                " values (" & varTRANDATE & "," & varTRANTYPE & "," & varTRANNO & "," & varITEMNO & "," & varPART_ORD & "," & varPART_SUP & "," & varTRANQTY & "," & varUNIT & "," & varTRANUCOST & "," & varTRANUPRICE & "," & varNETCOST & "," & varNETPRICE & "," & varSTATUS & "," & varIN_OUT & "," & varLISTED & "," & varMAC & "," & varTRANINVAMT & "," & varUSERCODE & "," & varLASTUPDATE & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldTdaytranHist.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldTdaytranHist.MoveNext
   Loop
   Me.Caption = "Tdaytran History File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub MoveOrdHd()
Dim MoveSql As String
Dim i As Integer

Dim varOHTRANTYPE As String
Dim varOHTRANNO As String
Dim varOHTRANDATE As String
Dim varOHCANCDATE As String
Dim varOHCUSTCODE As String
Dim varOHCUSTNAME As String
Dim varOHCHARGETO As String
Dim varOHRONO As String
Dim varOHSALESMAN As String
Dim varOHSMNAME As String
Dim varOHTERMS As String
Dim varOHTTLINVAMT As String
Dim varOHDS1 As String
Dim varOHDS_DESC1 As String
Dim varOHDS_AMT1 As String
Dim varOHNETINVAMT As String
Dim varOHNETCOST As String
Dim varOHSTATUS As String
Dim varOHNETINVAMT2 As String
Dim varOHNETCOST2 As String
Dim varOHLISTED As String
Dim varOHUSERCODE As String
Dim varOHLASTUPDATE As String
Dim varOHTOTINVAMT As String
Dim varOHDISCOUNT As String
Dim varOHVAT As String
Dim varOHNETINVOICE As String
Dim varOHTOTALCOST As String
Dim varOHREMARKS As String
Dim rsOldOrd_HD As ADODB.Recordset
gconPMIOS.Execute "delete from ord_hd"
Set rsOldOrd_HD = New ADODB.Recordset
    rsOldOrd_HD.Open "select * from ord_hd", gconOldPMIS
If Not rsOldOrd_HD.EOF And Not rsOldOrd_HD.BOF Then
   rsOldOrd_HD.MoveFirst
   Me.Caption = "Currently Converting Issuance Header File"
   DoEvents
   Screen.MousePointer = 11
   i = 0
   Do While Not rsOldOrd_HD.EOF
      varOHTRANTYPE = N2Str2Null(rsOldOrd_HD!trantype)
      varOHTRANNO = N2Str2Null(rsOldOrd_HD!tranno)
      varOHTRANDATE = N2Str2Null(rsOldOrd_HD!trandate)
      varOHCANCDATE = N2Str2Null(rsOldOrd_HD!cancdate)
      varOHCUSTCODE = N2Str2Null(rsOldOrd_HD!custcode)
      varOHCUSTNAME = N2Str2Null(Trim(rsOldOrd_HD!custname) & Trim(rsOldOrd_HD!custadrs1) & Trim(rsOldOrd_HD!custadrs2))
      varOHCHARGETO = N2Str2Null(rsOldOrd_HD!chargeto)
      varOHRONO = N2Str2Null(rsOldOrd_HD!rono)
      varOHSALESMAN = N2Str2Null(rsOldOrd_HD!salesman)
      varOHSMNAME = N2Str2Null(rsOldOrd_HD!smname)
      varOHTERMS = N2Str2Null(rsOldOrd_HD!terms)
      varOHTTLINVAMT = N2Str2Zero(rsOldOrd_HD!ttlinvamt)
      varOHDS1 = N2Str2IntZero(rsOldOrd_HD!ds1)
      varOHDS_DESC1 = N2Str2Null(rsOldOrd_HD!ds_desc1)
      varOHDS_AMT1 = N2Str2Zero(rsOldOrd_HD!ds_amt1)
      varOHNETINVAMT = N2Str2Zero(rsOldOrd_HD!netinvamt)
      varOHNETCOST = N2Str2Zero(rsOldOrd_HD!netcost)
      varOHSTATUS = N2Str2Null(rsOldOrd_HD!Status)
      varOHNETINVAMT2 = N2Str2Zero(rsOldOrd_HD!NETINVAMT2)
      varOHNETCOST2 = N2Str2Zero(rsOldOrd_HD!NETCOST2)
      varOHLISTED = N2Str2Null(rsOldOrd_HD!listed)
      varOHUSERCODE = N2Str2Null(rsOldOrd_HD!usercode)
      varOHLASTUPDATE = N2Str2Null(rsOldOrd_HD!lastupdate)
      varOHTOTINVAMT = N2Str2Zero(rsOldOrd_HD!TOTINVAMT)
      varOHDISCOUNT = N2Str2Zero(rsOldOrd_HD!DISCOUNT)
      varOHVAT = N2Str2Zero(rsOldOrd_HD!Vat)
      varOHNETINVOICE = N2Str2Zero(rsOldOrd_HD!NETINVOICE)
      varOHTOTALCOST = N2Str2Zero(rsOldOrd_HD!TotalCost)
      
      MoveSql = "INSERT INTO ORD_HD " & _
                "(TRANTYPE,TRANNO,TRANDATE,CANCDATE,CUSTCODE,CUSTNAME,CHARGETO,RONO,SALESMAN,SMNAME,TERMS,TTLINVAMT,DS1,DS_DESC1,DS_AMT1,NETINVAMT,NETCOST,STATUS,NETINVAMT2,NETCOST2,LISTED,USERCODE,LASTUPDATE,TOTINVAMT,DISCOUNT,VAT,NETINVOICE,TOTALCOST)" & _
                " values (" & varOHTRANTYPE & ", " & varOHTRANNO & ", " & varOHTRANDATE & ", " & varOHCANCDATE & ", " & varOHCUSTCODE & ", " & varOHCUSTNAME & ", " & varOHCHARGETO & ", " & varOHRONO & ", " & varOHSALESMAN & ", " & varOHSMNAME & ", " & varOHTERMS & ", " & varOHTTLINVAMT & ", " & varOHDS1 & ", " & varOHDS_DESC1 & ", " & varOHDS_AMT1 & ", " & varOHNETINVAMT & ", " & varOHNETCOST & ", " & varOHSTATUS & ", " & varOHNETINVAMT2 & ", " & varOHNETCOST2 & ", " & varOHLISTED & ", " & varOHUSERCODE & ", " & varOHLASTUPDATE & ", " & varOHTOTINVAMT & ", " & varOHDISCOUNT & ", " & varOHVAT & ", " & varOHNETINVOICE & ", " & varOHTOTALCOST & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldOrd_HD.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldOrd_HD.MoveNext
   Loop
   Me.Caption = "Issuance Header File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub MoveOrdHdHist()
Dim MoveSql As String
Dim i As Integer

Dim varOHTRANTYPE As String
Dim varOHTRANNO As String
Dim varOHTRANDATE As String
Dim varOHCANCDATE As String
Dim varOHCUSTCODE As String
Dim varOHCUSTNAME As String
Dim varOHCHARGETO As String
Dim varOHRONO As String
Dim varOHSALESMAN As String
Dim varOHSMNAME As String
Dim varOHTERMS As String
Dim varOHTTLINVAMT As String
Dim varOHDS1 As String
Dim varOHDS_DESC1 As String
Dim varOHDS_AMT1 As String
Dim varOHNETINVAMT As String
Dim varOHNETCOST As String
Dim varOHSTATUS As String
Dim varOHNETINVAMT2 As String
Dim varOHNETCOST2 As String
Dim varOHLISTED As String
Dim varOHUSERCODE As String
Dim varOHLASTUPDATE As String
Dim varOHTOTINVAMT As String
Dim varOHDISCOUNT As String
Dim varOHVAT As String
Dim varOHNETINVOICE As String
Dim varOHTOTALCOST As String
Dim varOHREMARKS As String
Dim rsOldOrd_HDHist As ADODB.Recordset
gconPMIOS.Execute "delete from ord_hist"
Set rsOldOrd_HDHist = New ADODB.Recordset
    rsOldOrd_HDHist.Open "select * from ord_hist", gconOldPMIS
If Not rsOldOrd_HDHist.EOF And Not rsOldOrd_HDHist.BOF Then
   rsOldOrd_HDHist.MoveFirst
   Me.Caption = "Currently Converting Issuance Header History File"
   DoEvents
   Screen.MousePointer = 11
   i = 0
   Do While Not rsOldOrd_HDHist.EOF
      varOHTRANTYPE = N2Str2Null(rsOldOrd_HDHist!trantype)
      varOHTRANNO = N2Str2Null(rsOldOrd_HDHist!tranno)
      varOHTRANDATE = N2Str2Null(rsOldOrd_HDHist!trandate)
      varOHCANCDATE = N2Str2Null(rsOldOrd_HDHist!cancdate)
      varOHCUSTCODE = N2Str2Null(rsOldOrd_HDHist!custcode)
      varOHCUSTNAME = N2Str2Null(Trim(rsOldOrd_HDHist!custname) & Trim(rsOldOrd_HDHist!custadrs1) & Trim(rsOldOrd_HDHist!custadrs2))
      varOHCHARGETO = N2Str2Null(rsOldOrd_HDHist!chargeto)
      varOHRONO = N2Str2Null(rsOldOrd_HDHist!rono)
      varOHSALESMAN = N2Str2Null(rsOldOrd_HDHist!salesman)
      varOHSMNAME = N2Str2Null(rsOldOrd_HDHist!smname)
      varOHTERMS = N2Str2Null(rsOldOrd_HDHist!terms)
      varOHTTLINVAMT = N2Str2Zero(rsOldOrd_HDHist!ttlinvamt)
      varOHDS1 = N2Str2IntZero(rsOldOrd_HDHist!ds1)
      varOHDS_DESC1 = N2Str2Null(rsOldOrd_HDHist!ds_desc1)
      varOHDS_AMT1 = N2Str2Zero(rsOldOrd_HDHist!ds_amt1)
      varOHNETINVAMT = N2Str2Zero(rsOldOrd_HDHist!netinvamt)
      varOHNETCOST = N2Str2Zero(rsOldOrd_HDHist!netcost)
      varOHSTATUS = N2Str2Null(rsOldOrd_HDHist!Status)
      varOHNETINVAMT2 = N2Str2Zero(rsOldOrd_HDHist!NETINVAMT2)
      varOHNETCOST2 = N2Str2Zero(rsOldOrd_HDHist!NETCOST2)
      varOHLISTED = N2Str2Null(rsOldOrd_HDHist!listed)
      varOHUSERCODE = N2Str2Null(rsOldOrd_HDHist!usercode)
      varOHLASTUPDATE = N2Str2Null(rsOldOrd_HDHist!lastupdate)
      varOHTOTINVAMT = N2Str2Zero(rsOldOrd_HDHist!TOTINVAMT)
      varOHDISCOUNT = N2Str2Zero(rsOldOrd_HDHist!DISCOUNT)
      varOHVAT = N2Str2Zero(rsOldOrd_HDHist!Vat)
      varOHNETINVOICE = N2Str2Zero(rsOldOrd_HDHist!NETINVOICE)
      varOHTOTALCOST = N2Str2Zero(rsOldOrd_HDHist!TotalCost)
      
      MoveSql = "INSERT INTO ORD_HIST " & _
                "(TRANTYPE,TRANNO,TRANDATE,CANCDATE,CUSTCODE,CUSTNAME,CHARGETO,RONO,SALESMAN,SMNAME,TERMS,TTLINVAMT,DS1,DS_DESC1,DS_AMT1,NETINVAMT,NETCOST,STATUS,NETINVAMT2,NETCOST2,LISTED,USERCODE,LASTUPDATE,TOTINVAMT,DISCOUNT,VAT,NETINVOICE,TOTALCOST)" & _
                " values (" & varOHTRANTYPE & ", " & varOHTRANNO & ", " & varOHTRANDATE & ", " & varOHCANCDATE & ", " & varOHCUSTCODE & ", " & varOHCUSTNAME & ", " & varOHCHARGETO & ", " & varOHRONO & ", " & varOHSALESMAN & ", " & varOHSMNAME & ", " & varOHTERMS & ", " & varOHTTLINVAMT & ", " & varOHDS1 & ", " & varOHDS_DESC1 & ", " & varOHDS_AMT1 & ", " & varOHNETINVAMT & ", " & varOHNETCOST & ", " & varOHSTATUS & ", " & varOHNETINVAMT2 & ", " & varOHNETCOST2 & ", " & varOHLISTED & ", " & varOHUSERCODE & ", " & varOHLASTUPDATE & ", " & varOHTOTINVAMT & ", " & varOHDISCOUNT & ", " & varOHVAT & ", " & varOHNETINVOICE & ", " & varOHTOTALCOST & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldOrd_HDHist.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldOrd_HDHist.MoveNext
   Loop
   Me.Caption = "Issuance Header History File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub MoveRRhd()
Dim MoveSql As String
Dim i As Integer

Dim varRRRRNO As String
Dim varRRRRDATE As String
Dim varRRCANCDATE As String
Dim varRRPONO As String
Dim varRRPODATE As String
Dim varRRRECVD_CODE As String
Dim varRRRECVD_FROM As String
Dim varRRADDRESS As String
Dim varRRDRNO As String
Dim varRRINVNO As String
Dim varRRCLASSCODE As String
Dim varRRTERMS As String
Dim varRRTTLRRAMT As String
Dim varRRDS1 As String
Dim varRRDS_DESC1 As String
Dim varRRDS_AMT1 As String
Dim varRRNETRRAMT As String
Dim varRRSTATUS As String
Dim varRRLISTED As String
Dim varRRUSERCODE As String
Dim varRRLASTUPDATE As String
Dim varRRREMARKS As String
Dim rsOldRR_HD As ADODB.Recordset
gconPMIOS.Execute "delete from rr_hd"
Set rsOldRR_HD = New ADODB.Recordset
    rsOldRR_HD.Open "select * from rr_hd", gconOldPMIS
If Not rsOldRR_HD.EOF And Not rsOldRR_HD.BOF Then
   Me.Caption = "Currently Converting Receipts Header File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldRR_HD.EOF
      varRRRRNO = N2Str2Null(rsOldRR_HD!rrno)
      varRRRRDATE = N2Str2Null(rsOldRR_HD!rrdate)
      varRRCANCDATE = N2Str2Null(rsOldRR_HD!cancdate)
      varRRPONO = N2Str2Null(rsOldRR_HD!pono)
      varRRPODATE = N2Str2Null(rsOldRR_HD!podate)
      varRRRECVD_CODE = N2Str2Null(rsOldRR_HD!recvd_code)
      varRRRECVD_FROM = N2Str2Null(Trim(rsOldRR_HD!recvd_from))
      varRRADDRESS = N2Str2Null(Trim(rsOldRR_HD!address1) & Trim(rsOldRR_HD!address2))
      varRRDRNO = N2Str2Null(rsOldRR_HD!drno)
      varRRINVNO = N2Str2Null(rsOldRR_HD!invno)
      varRRCLASSCODE = N2Str2Null(rsOldRR_HD!classcode)
      varRRTERMS = N2Str2Null(rsOldRR_HD!terms)
      varRRTTLRRAMT = N2Str2Zero(rsOldRR_HD!ttlrramt)
      varRRDS1 = N2Str2IntZero(rsOldRR_HD!ds1)
      varRRDS_DESC1 = N2Str2Null(rsOldRR_HD!ds_desc1)
      varRRDS_AMT1 = N2Str2Zero(rsOldRR_HD!ds_amt1)
      varRRNETRRAMT = N2Str2Zero(rsOldRR_HD!netrramt)
      varRRSTATUS = N2Str2Null(rsOldRR_HD!Status)
      varRRLISTED = N2Str2Null(rsOldRR_HD!listed)
      varRRUSERCODE = N2Str2Null(rsOldRR_HD!usercode)
      varRRLASTUPDATE = N2Str2Null(rsOldRR_HD!lastupdate)
      varRRREMARKS = N2Str2Null(rsOldRR_HD!remarks)
  
      MoveSql = "INSERT INTO RR_HD " & _
                "(RRNO,RRDATE,CANCDATE,PONO,PODATE,RECVD_CODE,RECVD_FROM,ADDRESS,DRNO,INVNO,CLASSCODE,TERMS,TTLRRAMT,DS1,DS_DESC1,DS_AMT1,NETRRAMT,STATUS,LISTED,USERCODE,LASTUPDATE,REMARKS)" & _
                " values (" & varRRRRNO & ", " & varRRRRDATE & ", " & varRRCANCDATE & ", " & varRRPONO & ", " & varRRPODATE & ", " & varRRRECVD_CODE & ", " & varRRRECVD_FROM & ", " & varRRADDRESS & ", " & varRRDRNO & ", " & varRRINVNO & ", " & varRRCLASSCODE & ", " & varRRTERMS & ", " & varRRTTLRRAMT & ", " & varRRDS1 & ", " & varRRDS_DESC1 & ", " & varRRDS_AMT1 & ", " & varRRNETRRAMT & ", " & varRRSTATUS & ", " & varRRLISTED & ", " & varRRUSERCODE & ", " & varRRLASTUPDATE & ", " & varRRREMARKS & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldRR_HD.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldRR_HD.MoveNext
   Loop
   Me.Caption = "Receipts Header File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub MoveRRhdHist()
Dim MoveSql As String
Dim i As Integer

Dim varRRRRNO As String
Dim varRRRRDATE As String
Dim varRRCANCDATE As String
Dim varRRPONO As String
Dim varRRPODATE As String
Dim varRRRECVD_CODE As String
Dim varRRRECVD_FROM As String
Dim varRRADDRESS As String
Dim varRRDRNO As String
Dim varRRINVNO As String
Dim varRRCLASSCODE As String
Dim varRRTERMS As String
Dim varRRTTLRRAMT As String
Dim varRRDS1 As String
Dim varRRDS_DESC1 As String
Dim varRRDS_AMT1 As String
Dim varRRNETRRAMT As String
Dim varRRSTATUS As String
Dim varRRLISTED As String
Dim varRRUSERCODE As String
Dim varRRLASTUPDATE As String
Dim varRRREMARKS As String
Dim rsOldRR_HDHist As ADODB.Recordset
gconPMIOS.Execute "delete from rec_hist"
Set rsOldRR_HDHist = New ADODB.Recordset
    rsOldRR_HDHist.Open "select * from rec_hist", gconOldPMIS
If Not rsOldRR_HDHist.EOF And Not rsOldRR_HDHist.BOF Then
   Me.Caption = "Currently Converting Receipts Header History File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldRR_HDHist.EOF
      varRRRRNO = N2Str2Null(rsOldRR_HDHist!rrno)
      varRRRRDATE = N2Str2Null(rsOldRR_HDHist!rrdate)
      varRRCANCDATE = N2Str2Null(rsOldRR_HDHist!cancdate)
      varRRPONO = N2Str2Null(rsOldRR_HDHist!pono)
      varRRPODATE = N2Str2Null(rsOldRR_HDHist!podate)
      varRRRECVD_CODE = N2Str2Null(rsOldRR_HDHist!recvd_code)
      varRRRECVD_FROM = N2Str2Null(Trim(rsOldRR_HDHist!recvd_from))
      varRRADDRESS = N2Str2Null(Trim(rsOldRR_HDHist!address1) & Trim(rsOldRR_HDHist!address2))
      varRRDRNO = N2Str2Null(rsOldRR_HDHist!drno)
      varRRINVNO = N2Str2Null(rsOldRR_HDHist!invno)
      varRRCLASSCODE = N2Str2Null(rsOldRR_HDHist!classcode)
      varRRTERMS = N2Str2Null(rsOldRR_HDHist!terms)
      varRRTTLRRAMT = N2Str2Zero(rsOldRR_HDHist!ttlrramt)
      varRRDS1 = N2Str2IntZero(rsOldRR_HDHist!ds1)
      varRRDS_DESC1 = N2Str2Null(rsOldRR_HDHist!ds_desc1)
      varRRDS_AMT1 = N2Str2Zero(rsOldRR_HDHist!ds_amt1)
      varRRNETRRAMT = N2Str2Zero(rsOldRR_HDHist!netrramt)
      varRRSTATUS = N2Str2Null(rsOldRR_HDHist!Status)
      varRRLISTED = N2Str2Null(rsOldRR_HDHist!listed)
      varRRUSERCODE = N2Str2Null(rsOldRR_HDHist!usercode)
      varRRLASTUPDATE = N2Str2Null(rsOldRR_HDHist!lastupdate)
      varRRREMARKS = N2Str2Null(rsOldRR_HDHist!remarks)
  
      MoveSql = "INSERT INTO REC_HIST " & _
                "(RRNO,RRDATE,CANCDATE,PONO,PODATE,RECVD_CODE,RECVD_FROM,ADDRESS,DRNO,INVNO,CLASSCODE,TERMS,TTLRRAMT,DS1,DS_DESC1,DS_AMT1,NETRRAMT,STATUS,LISTED,USERCODE,LASTUPDATE,REMARKS)" & _
                " values (" & varRRRRNO & ", " & varRRRRDATE & ", " & varRRCANCDATE & ", " & varRRPONO & ", " & varRRPODATE & ", " & varRRRECVD_CODE & ", " & varRRRECVD_FROM & ", " & varRRADDRESS & ", " & varRRDRNO & ", " & varRRINVNO & ", " & varRRCLASSCODE & ", " & varRRTERMS & ", " & varRRTTLRRAMT & ", " & varRRDS1 & ", " & varRRDS_DESC1 & ", " & varRRDS_AMT1 & ", " & varRRNETRRAMT & ", " & varRRSTATUS & ", " & varRRLISTED & ", " & varRRUSERCODE & ", " & varRRLASTUPDATE & ", " & varRRREMARKS & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldRR_HDHist.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldRR_HDHist.MoveNext
   Loop
   Me.Caption = "Receipts Header History File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub MovePOhd()
Dim MoveSql As String
Dim i As Integer

Dim varPOPONO As String
Dim varPOPODATE As String
Dim varPOORDERTYPE As String
Dim varPODON As String
Dim varPOSUPCODE As String
Dim varPOSUPNAME As String
Dim varPOSUP_ADDRS As String
Dim varPODEALERCODE As String
Dim varPOSHIPTO As String
Dim varPOSHP_ADDRS As String
Dim varPOPO_AMOUNT As String
Dim varPODS1 As String
Dim varPODS_DESC1 As String
Dim varPODS_AMT1 As String
Dim varPONETPOAMT As String
Dim varPOSTATUS As String
Dim varPOLISTED As String
Dim varPOUSERCODE As String
Dim varPOLASTUPDATE As String
Dim varPOREMARKS As String
Dim rsOldPO_HD As ADODB.Recordset
gconPMIOS.Execute "delete from po_hd"
Set rsOldPO_HD = New ADODB.Recordset
    rsOldPO_HD.Open "select * from po_hd", gconOldPMIS
If Not rsOldPO_HD.EOF And Not rsOldPO_HD.BOF Then
   Me.Caption = "Currently Converting Purchase Header File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldPO_HD.EOF
      varPOPONO = N2Str2Null(rsOldPO_HD!pono)
      varPOPODATE = N2Str2Null(rsOldPO_HD!podate)
      varPOORDERTYPE = N2Str2Null(rsOldPO_HD!ORDERTYPE)
      varPODON = N2Str2Null(rsOldPO_HD!don)
      varPOSUPCODE = N2Str2Null(rsOldPO_HD!SupCode)
      varPOSUPNAME = N2Str2Null(Trim(rsOldPO_HD!supname))
      varPOSUP_ADDRS = N2Str2Null(Trim(rsOldPO_HD!Sup_Addrs1) & Trim(rsOldPO_HD!Sup_Addrs2))
      varPODEALERCODE = N2Str2Null(rsOldPO_HD!dealercode)
      varPOSHIPTO = N2Str2Null(Trim(rsOldPO_HD!Shipto))
      varPOSHP_ADDRS = N2Str2Null(Trim(rsOldPO_HD!shp_addrs1) & Trim(rsOldPO_HD!shp_addrs2))
      varPOPO_AMOUNT = N2Str2Zero(rsOldPO_HD!po_amount)
      varPODS1 = N2Str2IntZero(rsOldPO_HD!ds1)
      varPODS_DESC1 = N2Str2Null(rsOldPO_HD!ds_desc1)
      varPODS_AMT1 = N2Str2Zero(rsOldPO_HD!ds_amt1)
      varPONETPOAMT = N2Str2Zero(rsOldPO_HD!netpoamt)
      varPOSTATUS = N2Str2Null(rsOldPO_HD!Status)
      varPOLISTED = N2Str2Null(rsOldPO_HD!listed)
      varPOUSERCODE = N2Str2Null(rsOldPO_HD!usercode)
      varPOLASTUPDATE = N2Str2Null(rsOldPO_HD!lastupdate)
      
      MoveSql = "INSERT INTO PO_HD " & _
                "(PONO,PODATE,ORDERTYPE,DON,SUPCODE,SUPNAME,SUP_ADDRS,DEALERCODE,SHIPTO,SHP_ADDRS,PO_AMOUNT,DS1,DS_DESC1,DS_AMT1,NETPOAMT,STATUS,LISTED,USERCODE,LASTUPDATE)" & _
                " values (" & varPOPONO & ", " & varPOPODATE & ", " & varPOORDERTYPE & ", " & varPODON & ", " & varPOSUPCODE & ", " & varPOSUPNAME & ", " & varPOSUP_ADDRS & ", " & varPODEALERCODE & ", " & varPOSHIPTO & ", " & varPOSHP_ADDRS & ", " & varPOPO_AMOUNT & ", " & varPODS1 & ", " & varPODS_DESC1 & ", " & varPODS_AMT1 & ", " & varPONETPOAMT & ", " & varPOSTATUS & ", " & varPOLISTED & ", " & varPOUSERCODE & ", " & varPOLASTUPDATE & ")"
      On Error GoTo ErrorCode
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldPO_HD.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldPO_HD.MoveNext
   Loop
   Me.Caption = "Purchase Header File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MovePOhdHist()
On Error GoTo ErrorCode
Dim MoveSql As String
Dim i As Integer

Dim varPOPONO As String
Dim varPOPODATE As String
Dim varPOORDERTYPE As String
Dim varPODON As String
Dim varPOSUPCODE As String
Dim varPOSUPNAME As String
Dim varPOSUP_ADDRS As String
Dim varPODEALERCODE As String
Dim varPOSHIPTO As String
Dim varPOSHP_ADDRS As String
Dim varPOPO_AMOUNT As String
Dim varPODS1 As String
Dim varPODS_DESC1 As String
Dim varPODS_AMT1 As String
Dim varPONETPOAMT As String
Dim varPOSTATUS As String
Dim varPOLISTED As String
Dim varPOUSERCODE As String
Dim varPOLASTUPDATE As String
Dim varPOREMARKS As String
Dim rsOldPO_HDHist As ADODB.Recordset
gconPMIOS.Execute "delete from po_hist"
Set rsOldPO_HDHist = New ADODB.Recordset
    rsOldPO_HDHist.Open "select * from po_hist", gconOldPMIS
If Not rsOldPO_HDHist.EOF And Not rsOldPO_HDHist.BOF Then
   Me.Caption = "Currently Converting Purchase Header History File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldPO_HDHist.EOF
      varPOPONO = N2Str2Null(rsOldPO_HDHist!pono)
      varPOPODATE = N2Str2Null(rsOldPO_HDHist!podate)
      varPOORDERTYPE = N2Str2Null(rsOldPO_HDHist!ORDERTYPE)
      varPODON = N2Str2Null(rsOldPO_HDHist!don)
      varPOSUPCODE = N2Str2Null(rsOldPO_HDHist!SupCode)
      varPOSUPNAME = N2Str2Null(Trim(rsOldPO_HDHist!supname))
      varPOSUP_ADDRS = N2Str2Null(Trim(rsOldPO_HDHist!Sup_Addrs1) & Trim(rsOldPO_HDHist!Sup_Addrs2))
      varPODEALERCODE = N2Str2Null(rsOldPO_HDHist!dealercode)
      varPOSHIPTO = N2Str2Null(Trim(rsOldPO_HDHist!Shipto))
      varPOSHP_ADDRS = N2Str2Null(Trim(rsOldPO_HDHist!shp_addrs1) & Trim(rsOldPO_HDHist!shp_addrs2))
      varPOPO_AMOUNT = N2Str2Zero(rsOldPO_HDHist!po_amount)
      varPODS1 = N2Str2IntZero(rsOldPO_HDHist!ds1)
      varPODS_DESC1 = N2Str2Null(rsOldPO_HDHist!ds_desc1)
      varPODS_AMT1 = N2Str2Zero(rsOldPO_HDHist!ds_amt1)
      varPONETPOAMT = N2Str2Zero(rsOldPO_HDHist!netpoamt)
      varPOSTATUS = N2Str2Null(rsOldPO_HDHist!Status)
      varPOLISTED = N2Str2Null(rsOldPO_HDHist!listed)
      varPOUSERCODE = N2Str2Null(rsOldPO_HDHist!usercode)
      varPOLASTUPDATE = N2Str2Null(rsOldPO_HDHist!lastupdate)
      
      MoveSql = "INSERT INTO PO_HIST " & _
                "(PONO,PODATE,ORDERTYPE,DON,SUPCODE,SUPNAME,SUP_ADDRS,DEALERCODE,SHIPTO,SHP_ADDRS,PO_AMOUNT,DS1,DS_DESC1,DS_AMT1,NETPOAMT,STATUS,LISTED,USERCODE,LASTUPDATE)" & _
                " values (" & varPOPONO & ", " & varPOPODATE & ", " & varPOORDERTYPE & ", " & varPODON & ", " & varPOSUPCODE & ", " & varPOSUPNAME & ", " & varPOSUP_ADDRS & ", " & varPODEALERCODE & ", " & varPOSHIPTO & ", " & varPOSHP_ADDRS & ", " & varPOPO_AMOUNT & ", " & varPODS1 & ", " & varPODS_DESC1 & ", " & varPODS_AMT1 & ", " & varPONETPOAMT & ", " & varPOSTATUS & ", " & varPOLISTED & ", " & varPOUSERCODE & ", " & varPOLASTUPDATE & ")"
      On Error GoTo ErrorCode
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldPO_HDHist.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldPO_HDHist.MoveNext
   Loop
   Me.Caption = "Purchase Header History File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MovePartmas()
Dim MoveSql As String
Dim i As Integer

Dim varPmasPARTNO As String
Dim varPmasPARTDESC As String
Dim varPmasINVCLASS As String
Dim varPmasVEHTYPE As String
Dim varPmasMODELCODE As String
Dim varPmasLOCATION As String
Dim varPmasMAC As Double
Dim varPmasMAD As Integer
Dim varPmasOLDNO As String
Dim varPmasNEWNO As String
Dim varPmasGENNO As String
Dim varPmasSRP As Double
Dim varPmasNOSHIP As Double
Dim varPmasLASTM_MAC As Double
Dim varPmasLASTM_MAD As Double
Dim varPmasLASTM_SELL As Double
Dim varPmasLASTM_OH As Integer
Dim varPmasLASTM_OO As Integer
Dim varPmasOnhand As Integer
Dim varPmasTrecqty As Double
Dim varPmasTISSQTY As Double
Dim varPmasOnOrder As Integer
Dim varPmasTpoqty As Integer
Dim varPmasPRQTY As Integer
Dim varPmasTPRQTY As Integer
Dim varPmasLAST_RECQ As Integer
Dim varPmasLAST_RECD As String
Dim varPmasLASTY_OH As Integer
Dim varPmasLASTY_MAC As Double
Dim varPmasLASTY_OO As Integer
Dim varPmasLASTY_ADJ As Integer
Dim varPmasHOLD As Integer
Dim varPmasSUPCODE As String
Dim varPmasVARIANCE As Integer
Dim varPmasSUBINVCLASS As String
Dim varPmasPHYCOUNT As Integer
Dim varPmasADJPHYCOUNT As Integer
Dim varPmasCUTOFFQTY As Integer
Dim varPmasCUTOFFMAC As Double
Dim varPmasRECEIPTS As Integer
Dim varPmasISSUANCES As Integer
Dim varPmasUSERCODE As String
Dim varPmasLASTUPDATE As String
Dim varPmasDNP As Double
Dim varPmasVALID_ICC As String
Dim varPmasSStock As Long
Dim varPmasResService As Long
Dim rsOldPartmas As ADODB.Recordset
gconPMIOS.Execute "delete from partmas"
Set rsOldPartmas = New ADODB.Recordset
    rsOldPartmas.Open "select * from partmas order by partdesc asc", gconOldPMIS
If Not rsOldPartmas.EOF And Not rsOldPartmas.BOF Then
   Me.Caption = "Currently Converting Part Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldPartmas.EOF
      varPmasPARTNO = N2Str2Null(rsOldPartmas!PartNo)
      varPmasPARTDESC = N2Str2Null(rsOldPartmas!PartDesc)
      varPmasINVCLASS = N2Str2Null(rsOldPartmas!InvClass)
      varPmasMAD = N2Str2IntZero(rsOldPartmas!mad)
      If varPmasINVCLASS = "'A'" Or varPmasINVCLASS = "'B'" Or varPmasINVCLASS = "'C'" Then
         varPmasSStock = varPmasMAD * 2
         varPmasResService = varPmasMAD
      Else
         varPmasSStock = varPmasMAD
         varPmasResService = varPmasMAD / 2
      End If
      varPmasVEHTYPE = N2Str2Null(rsOldPartmas!vehtype)
      varPmasMODELCODE = N2Str2Null(rsOldPartmas!modelcode)
      varPmasLOCATION = N2Str2Null(rsOldPartmas!location)
      varPmasMAC = N2Str2Zero(rsOldPartmas!MAC)
      varPmasOLDNO = N2Str2Null(rsOldPartmas!oldno)
      varPmasNEWNO = N2Str2Null(rsOldPartmas!newno)
      varPmasGENNO = N2Str2Null(rsOldPartmas!genno)
      varPmasSRP = N2Str2Zero(rsOldPartmas!srp)
      varPmasNOSHIP = N2Str2Zero(rsOldPartmas!noship)
      varPmasLASTM_MAC = N2Str2Zero(rsOldPartmas!lastm_mac)
      varPmasLASTM_MAD = N2Str2Zero(rsOldPartmas!lastm_mad)
      varPmasLASTM_SELL = N2Str2Zero(rsOldPartmas!lastm_sell)
      varPmasLASTM_OH = N2Str2IntZero(rsOldPartmas!lastm_oh)
      varPmasLASTM_OO = N2Str2IntZero(rsOldPartmas!lastm_oo)
      If varPmasLASTM_OO < 0 Then varPmasLASTM_OO = 0
      varPmasOnhand = N2Str2IntZero(rsOldPartmas!Onhand)
      varPmasTrecqty = N2Str2IntZero(rsOldPartmas!trecqty)
      varPmasTISSQTY = N2Str2IntZero(rsOldPartmas!tissqty)
      varPmasOnOrder = N2Str2IntZero(rsOldPartmas!onorder)
      If varPmasOnOrder < 0 Then varPmasOnOrder = 0
      varPmasTpoqty = N2Str2IntZero(rsOldPartmas!tpoqty)
      varPmasPRQTY = N2Str2IntZero(rsOldPartmas!prqty)
      varPmasTPRQTY = N2Str2IntZero(rsOldPartmas!tprqty)
      varPmasLAST_RECQ = N2Str2IntZero(rsOldPartmas!last_recq)
      varPmasLAST_RECD = N2Date2Null(rsOldPartmas!Last_Recd)
      varPmasLASTY_OH = N2Str2IntZero(rsOldPartmas!lasty_oh)
      varPmasLASTY_MAC = N2Str2Zero(rsOldPartmas!lasty_mac)
      varPmasLASTY_OO = N2Str2IntZero(rsOldPartmas!lasty_oo)
      varPmasLASTY_ADJ = N2Str2IntZero(rsOldPartmas!lasty_adj)
      varPmasHOLD = N2Str2IntZero(rsOldPartmas!hold)
      varPmasSUPCODE = N2Str2Null(rsOldPartmas!SupCode)
      varPmasVARIANCE = N2Str2IntZero(rsOldPartmas!variance)
      varPmasSUBINVCLASS = N2Str2Null(rsOldPartmas!SubInvClas)
      varPmasPHYCOUNT = N2Str2IntZero(rsOldPartmas!phycount)
      varPmasADJPHYCOUNT = N2Str2IntZero(rsOldPartmas!adjphycnt)
      varPmasCUTOFFQTY = N2Str2IntZero(rsOldPartmas!CUTOFFQTY)
      varPmasCUTOFFMAC = N2Str2Zero(rsOldPartmas!CUTOFFMAC)
      varPmasRECEIPTS = N2Str2IntZero(rsOldPartmas!receipts)
      varPmasISSUANCES = N2Str2IntZero(rsOldPartmas!issuances)
      varPmasUSERCODE = N2Str2Null(rsOldPartmas!usercode)
      varPmasLASTUPDATE = N2Date2Null(rsOldPartmas!lastupdate)
      varPmasDNP = N2Str2Zero(rsOldPartmas!dnp)
      varPmasVALID_ICC = N2Str2Null(rsOldPartmas!valid_icc)
      If varPmasPARTNO <> "NULL" Then
         MoveSql = "INSERT INTO partmas " & _
                   "(PARTNO,PARTDESC,INVCLASS,VEHTYPE,MODELCODE,LOCATION,MAC,MAD,OLDNO,NEWNO,GENNO,SRP,NOSHIP,LASTM_MAC,LASTM_MAD,LASTM_SELL,LASTM_OH,LASTM_OO,ONHAND,TRECQTY,TISSQTY,ONORDER,TPOQTY,PRQTY,TPRQTY,LAST_RECQ,LAST_RECD,LASTY_OH,LASTY_MAC,LASTY_OO,LASTY_ADJ,HOLD,SUPCODE,VARIANCE,SUBINVCLAS,PHYCOUNT,ADJPHYCNT,CUTOFFQTY,CUTOFFMAC,RECEIPTS,ISSUANCES,USERCODE,LASTUPDATE,DNP,VALID_ICC,DATE_ENTERED,sstock,resservice)" & _
                   " values (" & varPmasPARTNO & "," & varPmasPARTDESC & "," & varPmasINVCLASS & "," & varPmasVEHTYPE & "," & varPmasMODELCODE & "," & varPmasLOCATION & "," & varPmasMAC & "," & varPmasMAD & "," & varPmasOLDNO & "," & varPmasNEWNO & "," & varPmasGENNO & "," & varPmasSRP & "," & varPmasNOSHIP & "," & varPmasLASTM_MAC & "," & varPmasLASTM_MAD & "," & varPmasLASTM_SELL & "," & varPmasLASTM_OH & "," & varPmasLASTM_OO & "," & varPmasOnhand & "," & varPmasTrecqty & "," & varPmasTISSQTY & "," & varPmasOnOrder & "," & varPmasTpoqty & "," & varPmasPRQTY & "," & varPmasTPRQTY & "," & varPmasLAST_RECQ & "," & varPmasLAST_RECD & "," & varPmasLASTY_OH & "," & varPmasLASTY_MAC & "," & varPmasLASTY_OO & "," & varPmasLASTY_ADJ & "," & varPmasHOLD & "," & _
                   " " & varPmasSUPCODE & "," & varPmasVARIANCE & "," & varPmasSUBINVCLASS & "," & varPmasPHYCOUNT & "," & varPmasADJPHYCOUNT & "," & varPmasCUTOFFQTY & "," & varPmasCUTOFFMAC & "," & varPmasRECEIPTS & "," & varPmasISSUANCES & "," & varPmasUSERCODE & "," & varPmasLASTUPDATE & "," & varPmasDNP & "," & varPmasVALID_ICC & ", " & "NULL" & ", " & varPmasSStock & ", " & varPmasResService & ")"
         On Error GoTo ErrorCode
         gconPMIOS.Execute MoveSql
      End If
      i = i + 1
      progCPB.Value = (i / rsOldPartmas.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldPartmas.MoveNext
   Loop
   Me.Caption = "Part Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveShipping()
Dim MoveSql As String
Dim i As Integer

Dim SMonths_60, SMonths_59, SMonths_58, SMonths_57 As Integer
Dim SMonths_56, SMonths_55, SMonths_54, SMonths_53 As Integer
Dim SMonths_52, SMonths_51, SMonths_50, SMonths_49 As Integer
Dim SMonths_48, SMonths_47, SMonths_46, SMonths_45 As Integer
Dim SMonths_44, SMonths_43, SMonths_42, SMonths_41  As Integer
Dim SMonths_40, SMonths_39, SMonths_38, SMonths_37 As Integer
Dim SMonths_36, SMonths_35, SMonths_34, SMonths_33 As Integer
Dim SMonths_32, SMonths_31, SMonths_30, SMonths_29 As Integer
Dim SMonths_28, SMonths_27, SMonths_26, SMonths_25 As Integer
Dim SMonths_24, SMonths_23, SMonths_22, SMonths_21 As Integer
Dim SMonths_20, SMonths_19, SMonths_18, SMonths_17 As Integer
Dim SMonths_16, SMonths_15, SMonths_14, SMonths_13 As Integer
Dim SMonths_12, SMonths_11, SMonths_10, SMonths_9  As Integer
Dim SMonths_8, SMonths_7, SMonths_6, SMonths_5     As Integer
Dim SMonths_4, SMonths_3, SMonths_2, SPrev_Month As Integer
Dim SCurr_Month, SFreq_Curr As Integer
Dim SPartno As String
Dim rsOldShipping As ADODB.Recordset
gconPMIOS.Execute "delete from shipping"
Set rsOldShipping = New ADODB.Recordset
    rsOldShipping.Open "Select * from Shipping order by partno asc", gconOldPMIS
If Not rsOldShipping.EOF And Not rsOldShipping.BOF Then
   rsOldShipping.MoveFirst
   Me.Caption = "Currently Converting Shipping File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldShipping.EOF
      SMonths_60 = N2Str2IntZero(rsOldShipping!months_60)
      SMonths_59 = N2Str2IntZero(rsOldShipping!months_59)
      SMonths_58 = N2Str2IntZero(rsOldShipping!months_58)
      SMonths_57 = N2Str2IntZero(rsOldShipping!months_57)
      SMonths_56 = N2Str2IntZero(rsOldShipping!Months_56)
      SMonths_55 = N2Str2IntZero(rsOldShipping!months_55)
      SMonths_54 = N2Str2IntZero(rsOldShipping!months_54)
      SMonths_53 = N2Str2IntZero(rsOldShipping!months_53)
      SMonths_52 = N2Str2IntZero(rsOldShipping!months_52)
      SMonths_51 = N2Str2IntZero(rsOldShipping!months_51)
      SMonths_50 = N2Str2IntZero(rsOldShipping!months_50)
      SMonths_49 = N2Str2IntZero(rsOldShipping!months_49)
      SMonths_48 = N2Str2IntZero(rsOldShipping!months_48)
      SMonths_47 = N2Str2IntZero(rsOldShipping!months_47)
      SMonths_46 = N2Str2IntZero(rsOldShipping!months_46)
      SMonths_45 = N2Str2IntZero(rsOldShipping!months_45)
      SMonths_44 = N2Str2IntZero(rsOldShipping!months_44)
      SMonths_43 = N2Str2IntZero(rsOldShipping!months_43)
      SMonths_42 = N2Str2IntZero(rsOldShipping!months_42)
      SMonths_41 = N2Str2IntZero(rsOldShipping!months_41)
      SMonths_40 = N2Str2IntZero(rsOldShipping!months_40)
      SMonths_39 = N2Str2IntZero(rsOldShipping!months_39)
      SMonths_38 = N2Str2IntZero(rsOldShipping!months_38)
      SMonths_37 = N2Str2IntZero(rsOldShipping!months_37)
      SMonths_36 = N2Str2IntZero(rsOldShipping!months_36)
      SMonths_35 = N2Str2IntZero(rsOldShipping!months_35)
      SMonths_34 = N2Str2IntZero(rsOldShipping!months_34)
      SMonths_33 = N2Str2IntZero(rsOldShipping!months_33)
      SMonths_32 = N2Str2IntZero(rsOldShipping!months_32)
      SMonths_31 = N2Str2IntZero(rsOldShipping!months_31)
      SMonths_30 = N2Str2IntZero(rsOldShipping!months_30)
      SMonths_29 = N2Str2IntZero(rsOldShipping!months_29)
      SMonths_28 = N2Str2IntZero(rsOldShipping!months_28)
      SMonths_27 = N2Str2IntZero(rsOldShipping!months_27)
      SMonths_26 = N2Str2IntZero(rsOldShipping!months_26)
      SMonths_25 = N2Str2IntZero(rsOldShipping!months_25)
      SMonths_24 = N2Str2IntZero(rsOldShipping!months_24)
      SMonths_23 = N2Str2IntZero(rsOldShipping!months_23)
      SMonths_22 = N2Str2IntZero(rsOldShipping!months_22)
      SMonths_21 = N2Str2IntZero(rsOldShipping!months_21)
      SMonths_20 = N2Str2IntZero(rsOldShipping!months_20)
      SMonths_19 = N2Str2IntZero(rsOldShipping!months_19)
      SMonths_18 = N2Str2IntZero(rsOldShipping!months_18)
      SMonths_17 = N2Str2IntZero(rsOldShipping!months_17)
      SMonths_16 = N2Str2IntZero(rsOldShipping!months_16)
      SMonths_15 = N2Str2IntZero(rsOldShipping!months_15)
      SMonths_14 = N2Str2IntZero(rsOldShipping!months_14)
      SMonths_13 = N2Str2IntZero(rsOldShipping!months_13)
      SMonths_12 = N2Str2IntZero(rsOldShipping!Months_12)
      SMonths_11 = N2Str2IntZero(rsOldShipping!Months_11)
      SMonths_10 = N2Str2IntZero(rsOldShipping!Months_10)
      SMonths_9 = N2Str2IntZero(rsOldShipping!Months_9)
      SMonths_8 = N2Str2IntZero(rsOldShipping!Months_8)
      SMonths_7 = N2Str2IntZero(rsOldShipping!Months_7)
      SMonths_6 = N2Str2IntZero(rsOldShipping!Months_6)
      SMonths_5 = N2Str2IntZero(rsOldShipping!Months_5)
      SMonths_4 = N2Str2IntZero(rsOldShipping!Months_4)
      SMonths_3 = N2Str2IntZero(rsOldShipping!Months_3)
      SMonths_2 = N2Str2IntZero(rsOldShipping!Months_2)
      SPrev_Month = N2Str2IntZero(rsOldShipping!Prev_Month)
      SCurr_Month = N2Str2IntZero(rsOldShipping!curr_month)
      SPartno = N2Str2Null(rsOldShipping!PartNo)
      SFreq_Curr = N2Str2IntZero(rsOldShipping!freq_curr)
      
      MoveSql = "insert into shipping" & _
                "(months_60,months_59,months_58,months_57,months_56,months_55,months_54,months_53,months_52,months_51,months_50,months_49," & _
                "months_48,months_47,months_46,months_45,months_44,months_43,months_42,months_41,months_40,months_39,months_38,months_37," & _
                "months_36,months_35,months_34,months_33,months_32,months_31,months_30,months_29,months_28,months_27,months_26,months_25," & _
                "months_24,months_23,months_22,months_21,months_20,months_19,months_18,months_17,months_16,months_15,months_14,months_13," & _
                "months_12,months_11,months_10,months_9,months_8,months_7,months_6,months_5,months_4,months_3,months_2,prev_month,curr_Month,partno,freq_curr)" & _
                " values (" & SMonths_60 & ", " & SMonths_59 & ", " & SMonths_58 & "," & SMonths_57 & "," & SMonths_56 & _
                "," & SMonths_55 & "," & SMonths_54 & "," & SMonths_53 & "," & SMonths_52 & "," & SMonths_51 & "," & SMonths_50 & "," & SMonths_49 & "," & SMonths_48 & _
                "," & SMonths_47 & "," & SMonths_46 & "," & SMonths_45 & "," & SMonths_44 & "," & SMonths_43 & "," & SMonths_42 & "," & SMonths_41 & "," & SMonths_40 & _
                "," & SMonths_39 & "," & SMonths_38 & "," & SMonths_37 & "," & SMonths_36 & "," & SMonths_35 & "," & SMonths_34 & "," & SMonths_33 & "," & SMonths_32 & _
                "," & SMonths_31 & "," & SMonths_30 & "," & SMonths_29 & "," & SMonths_28 & "," & SMonths_27 & "," & SMonths_26 & "," & SMonths_25 & "," & SMonths_24 & _
                "," & SMonths_23 & "," & SMonths_22 & "," & SMonths_21 & "," & SMonths_20 & "," & SMonths_19 & "," & SMonths_18 & "," & SMonths_17 & "," & SMonths_16 & _
                "," & SMonths_15 & "," & SMonths_14 & "," & SMonths_13 & "," & SMonths_12 & "," & SMonths_11 & "," & SMonths_10 & "," & SMonths_9 & "," & SMonths_8 & _
                "," & SMonths_7 & "," & SMonths_6 & "," & SMonths_5 & "," & SMonths_4 & "," & SMonths_3 & "," & SMonths_2 & "," & SPrev_Month & "," & SCurr_Month & "," & SPartno & "," & SFreq_Curr & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldShipping.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldShipping.MoveNext
   Loop
   Me.Caption = "Shipping File Successfully Converted"
   Screen.MousePointer = 0
End If
End Sub

Sub MoveRankFle()
Dim MoveSql As String
Dim i As Integer

Dim PartNo, PartDesc, InvClass, SubInvClas, Last_Recd As String
Dim Onhand, MAC, MAD12, SALES12 As Double
Dim PrevClass, PrevSClas As String
Dim Month_Gen As Integer
Dim Months_12, Months_11, Months_10, Months_9  As Integer
Dim Months_8, Months_7, Months_6, Months_5     As Integer
Dim Months_4, Months_3, Months_2, Prev_Month As Integer
Dim rsOldRankFle As ADODB.Recordset
gconPMIOS.Execute "delete from rankfle"
Set rsOldRankFle = New ADODB.Recordset
    rsOldRankFle.Open "Select * from RankFle order by partno asc", gconOldPMIS
If Not rsOldRankFle.EOF And Not rsOldRankFle.BOF Then
   rsOldRankFle.MoveFirst
   Me.Caption = "Currently Converting Ranking File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldRankFle.EOF
      Months_12 = N2Str2IntZero(rsOldRankFle!Months_12)
      Months_11 = N2Str2IntZero(rsOldRankFle!Months_11)
      Months_10 = N2Str2IntZero(rsOldRankFle!Months_10)
      Months_9 = N2Str2IntZero(rsOldRankFle!Months_9)
      Months_8 = N2Str2IntZero(rsOldRankFle!Months_8)
      Months_7 = N2Str2IntZero(rsOldRankFle!Months_7)
      Months_6 = N2Str2IntZero(rsOldRankFle!Months_6)
      Months_5 = N2Str2IntZero(rsOldRankFle!Months_5)
      Months_4 = N2Str2IntZero(rsOldRankFle!Months_4)
      Months_3 = N2Str2IntZero(rsOldRankFle!Months_3)
      Months_2 = N2Str2IntZero(rsOldRankFle!Months_2)
      Prev_Month = N2Str2IntZero(rsOldRankFle!Prev_Month)
      SALES12 = N2Str2Zero(rsOldRankFle!SALES12)
      Onhand = N2Str2IntZero(rsOldRankFle!Onhand)
      MAC = N2Str2Zero(rsOldRankFle!MAC)
      MAD12 = N2Str2Zero(rsOldRankFle!MAD12)
      PartNo = N2Str2Null(rsOldRankFle!PartNo)
      PartDesc = N2Str2Null(rsOldRankFle!PartDesc)
      InvClass = N2Str2Null(rsOldRankFle!InvClass)
      SubInvClas = N2Str2Null(rsOldRankFle!SubInvClas)
      Last_Recd = N2Str2Null(rsOldRankFle!Last_Recd)
      'PrevClass = N2Str2Null(rsOldRankFle!PrevClass)
      'PrevSClas = N2Str2Null(rsOldRankFle!PrevSClas)
      
      MoveSql = "insert into rankfle" & _
                "(partno,partdesc,invclass,subinvclas,last_recd,prevclass,prevsclas,mad12,mac,onhand,sales12," & _
                "months_12,months_11,months_10,months_9,months_8,months_7,months_6,months_5,months_4,months_3,months_2,prev_month)" & _
                " values (" & PartNo & ", " & PartDesc & ", " & InvClass & "," & SubInvClas & "," & Last_Recd & _
                "," & PrevClass & "," & PrevSClas & "," & MAD12 & "," & MAC & "," & Onhand & "," & SALES12 & _
                "," & Months_12 & "," & Months_11 & "," & Months_10 & "," & Months_9 & "," & Months_8 & _
                "," & Months_7 & "," & Months_6 & "," & Months_5 & "," & Months_4 & "," & Months_3 & "," & Months_2 & "," & Prev_Month & ")"
      'gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldRankFle.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldRankFle.MoveNext
   Loop
   Me.Caption = "Ranking File Successfully Converted"
   Screen.MousePointer = 0
End If
End Sub

Sub MoveSupplier()
Dim MoveSql As String
Dim i As Integer

Dim varSupSupCode As String
Dim varSupSupName As String
Dim varSupSup_addrs As String
Dim varSupPhoneNo As String
Dim varSupContact As String
Dim varSupDisc_Surch As Double
Dim varSupVat As Double
Dim varSupVat_Percnt As Double
Dim varSupLastUpdate As String
Dim varSupUserCode As String

Dim rsOldSupplier As ADODB.Recordset
gconPMIOS.Execute "delete from supplier"
Set rsOldSupplier = New ADODB.Recordset
    rsOldSupplier.Open "select * from Supplier order by supcode asc", gconOldPMIS
If Not rsOldSupplier.EOF And Not rsOldSupplier.BOF Then
   Me.Caption = "Currently Converting Supplier Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldSupplier.EOF
      varSupSupCode = N2Str2Null(rsOldSupplier!SupCode)
      varSupSupName = N2Str2Null(Trim(rsOldSupplier!supname))
      varSupSup_addrs = N2Str2Null(Trim(rsOldSupplier!Sup_Addrs1) & Trim(rsOldSupplier!Sup_Addrs2))
      varSupPhoneNo = N2Str2Null(rsOldSupplier!phoneno)
      varSupContact = N2Str2Null(rsOldSupplier!contact)
      varSupDisc_Surch = N2Str2Zero(rsOldSupplier!disc_surch)
      varSupVat = N2Str2Zero(rsOldSupplier!Vat)
      varSupVat_Percnt = N2Str2Zero(rsOldSupplier!vat_percnt)
      varSupLastUpdate = N2Str2Null(rsOldSupplier!lastupdate)
      varSupUserCode = N2Str2Null(rsOldSupplier!usercode)
      
      MoveSql = "INSERT INTO Supplier " & _
                "(supcode,supname,sup_addrs,phoneno,contact,disc_surch,vat,vat_percnt,lastupdate,usercode)" & _
                " values (" & varSupSupCode & ", " & varSupSupName & ", " & varSupSup_addrs & ", " & varSupPhoneNo & ", " & varSupContact & ", " & varSupDisc_Surch & ", " & varSupVat & ", " & varSupVat_Percnt & ", " & varSupLastUpdate & ", " & varSupUserCode & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldSupplier.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldSupplier.MoveNext
   Loop
   Me.Caption = "Supplier Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub MoveCustomer()
Dim MoveSql As String
Dim i As Integer

Dim varCustCustCode As String
Dim varCustCustName As String
Dim varCustCustaddrs As String
Dim varCustPhoneNo As String
Dim varCustCR_Amount As Double
Dim varCustCR_Limit As Double
Dim varCustDisc_Surch As Double
Dim varCustPartsPrice As String
Dim varCustLastUpdate As String
Dim varCustUserCode As String

Dim rsOldCustomer As ADODB.Recordset
gconPMIOS.Execute "delete from customer"
Set rsOldCustomer = New ADODB.Recordset
    rsOldCustomer.Open "select * from Customer order by custcode asc", gconOldPMIS
If Not rsOldCustomer.EOF And Not rsOldCustomer.BOF Then
   Me.Caption = "Currently Converting Customer Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldCustomer.EOF
      varCustCustCode = N2Str2Null(rsOldCustomer!custcode)
      varCustCustName = N2Str2Null(Trim(rsOldCustomer!custname))
      varCustCustaddrs = N2Str2Null(Trim(rsOldCustomer!custadrs1) & Trim(rsOldCustomer!custadrs2))
      varCustPhoneNo = N2Str2Null(rsOldCustomer!phoneno)
      varCustCR_Amount = N2Str2Zero(rsOldCustomer!cr_amount)
      varCustCR_Limit = N2Str2Zero(rsOldCustomer!cr_limit)
      varCustDisc_Surch = N2Str2Zero(rsOldCustomer!disc_surch)
      varCustPartsPrice = N2Str2Null(rsOldCustomer!partsprice)
      varCustLastUpdate = N2Str2Null(rsOldCustomer!lastupdate)
      varCustUserCode = N2Str2Null(rsOldCustomer!usercode)
      
      MoveSql = "INSERT INTO Customer " & _
                "(custcode,custname,custadrs,phoneno,cr_amount,cr_limit,disc_surch,partsprice,lastupdate,usercode)" & _
                " values (" & varCustCustCode & ", " & varCustCustName & ", " & varCustCustaddrs & ", " & varCustPhoneNo & ", " & varCustCR_Amount & ", " & varCustCR_Limit & ", " & varCustDisc_Surch & ", " & varCustPartsPrice & ", " & varCustLastUpdate & ", " & varCustUserCode & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldCustomer.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldCustomer.MoveNext
   Loop
   Me.Caption = "Customer Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub MoveCunter()
Dim MoveSql As String
Dim i As Integer

Dim varCuntModul As String
Dim varCuntNextNumber As Long
Dim varCuntLastUpdate As String
Dim varCuntUserCode As String

Dim rsOldCunter As ADODB.Recordset
gconPMIOS.Execute "delete from cunter"
Set rsOldCunter = New ADODB.Recordset
    rsOldCunter.Open "select * from [Counter] order by modul asc", gconOldPMIS
If Not rsOldCunter.EOF And Not rsOldCunter.BOF Then
   Me.Caption = "Currently Converting Counter Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldCunter.EOF
      varCuntModul = N2Str2Null(rsOldCunter![module])
      varCuntNextNumber = N2Str2IntZero(rsOldCunter!nextnumber)
      varCuntLastUpdate = N2Str2Null(rsOldCunter!lastupdate)
      varCuntUserCode = N2Str2Null(rsOldCunter!usercode)
      
      MoveSql = "INSERT INTO Cunter " & _
                "(modul,nextnumber,lastupdate,usercode)" & _
                " values (" & varCuntModul & ", " & varCuntNextNumber & ", " & varCuntLastUpdate & ", " & varCuntUserCode & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldCunter.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldCunter.MoveNext
   Loop
   MoveSql = "INSERT INTO Cunter " & _
                "(modul,nextnumber,lastupdate,usercode)" & _
                " values ('PP',1, '" & LOGDATE & "', 'WIZ')"
   gconPMIOS.Execute MoveSql
   Me.Caption = "Counter Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub MoveSalesman()
Dim MoveSql As String
Dim i As Integer

Dim varSMEmpno As String
Dim varSMLastname As String
Dim varSMFirstname As String
Dim varSMMiddleInt As String
Dim varSMFullname As String
Dim varSMSignname As String
Dim varSMPosition As String
Dim varSMLastupdate As String
Dim varSMusercode As String

Dim rsOldSalesman As ADODB.Recordset
gconPMIOS.Execute "delete from salesman"
Set rsOldSalesman = New ADODB.Recordset
    rsOldSalesman.Open "select * from salesman order by empno asc", gconOldPMIS
If Not rsOldSalesman.EOF And Not rsOldSalesman.BOF Then
   Me.Caption = "Currently Converting Salesman Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldSalesman.EOF
      varSMEmpno = N2Str2Null(rsOldSalesman!empno)
      varSMLastname = N2Str2Null(rsOldSalesman!lastname)
      varSMFirstname = N2Str2Null(rsOldSalesman!firstname)
      varSMMiddleInt = N2Str2Null(rsOldSalesman!middleint)
      varSMFullname = N2Str2Null(rsOldSalesman!FullName)
      varSMSignname = N2Str2Null(rsOldSalesman!signname)
      varSMPosition = N2Str2Null(rsOldSalesman!Position)
      varSMLastupdate = N2Str2Null(rsOldSalesman!lastupdate)
      varSMusercode = N2Str2Null(rsOldSalesman!usercode)

      MoveSql = "INSERT INTO salesman " & _
                "(empno,lastname,firstname,middleint,fullname,signname,positions,lastupdate,usercode)" & _
                " values (" & varSMEmpno & ", " & varSMLastname & ", " & varSMFirstname & ", " & varSMMiddleInt & ", " & varSMFullname & ", " & varSMSignname & ", " & varSMPosition & ", " & varSMLastupdate & ", " & varSMusercode & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldSalesman.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldSalesman.MoveNext
   Loop
   Me.Caption = "Salesman Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub MoveCustCtl()
Dim MoveSql As String
Dim i As Integer

Dim varCustCtlcde As String
Dim varCustCtldsc As String

Dim rsOldCustCtl As ADODB.Recordset
gconPMIOS.Execute "delete from cusctl"
Set rsOldCustCtl = New ADODB.Recordset
    rsOldCustCtl.Open "select * from CusCtl order by ctlcde asc", gconOldPMIS
If Not rsOldCustCtl.EOF And Not rsOldCustCtl.BOF Then
   Me.Caption = "Currently Converting Customer Control Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldCustCtl.EOF
      varCustCtlcde = N2Str2Null(rsOldCustCtl!ctlcde)
      varCustCtldsc = N2Str2Null(rsOldCustCtl!ctldsc)
      
      MoveSql = "INSERT INTO cusctl " & _
                "(ctlcde,ctldsc)" & _
                " values (" & varCustCtlcde & ", " & varCustCtldsc & ")"
      gconPMIOS.Execute MoveSql
      i = i + 1
      progCPB.Value = (i / rsOldCustCtl.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsOldCustCtl.MoveNext
   Loop
   Me.Caption = "Customer Control Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub MoveLocation()
Dim rsPartmas As ADODB.Recordset
Dim PrevLoc As String
Dim i As Integer
gconPMIOS.Execute "delete from location"
Set rsPartmas = New ADODB.Recordset
    rsPartmas.Open "select distinct location from partmas order by location asc", gconPMIOS
If Not rsPartmas.EOF And Not rsPartmas.BOF Then
   rsPartmas.MoveFirst
   Me.Caption = "Currently Converting Location Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   PrevLoc = ""
   Do While Not rsPartmas.EOF
      If Null2String(rsPartmas!location) <> "" Then
         gconPMIOS.Execute "insert into location (location) values (" & N2Str2Null(rsPartmas!location) & ")"
      End If
      i = i + 1
      progCPB.Value = (i / rsPartmas.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsPartmas.MoveNext
   Loop
   Me.Caption = "Location Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
End Sub

Sub UpdateLastM_MAC()
Dim i As Integer
Dim rsPartmas As ADODB.Recordset
Dim rsPartmasBackup As ADODB.Recordset

Set rsPartmasBackup = New ADODB.Recordset
    rsPartmasBackup.Open "select partno,lastm_mac from partmas", gconPMIOSBackUp
If Not rsPartmasBackup.EOF And Not rsPartmasBackup.BOF Then
   rsPartmasBackup.MoveFirst
   i = 0
   Do While Not rsPartmasBackup.EOF
      Set rsPartmas = New ADODB.Recordset
          rsPartmas.Open "select partno from partmas where partno = " & N2Str2Null(rsPartmasBackup!PartNo), gconPMIOS
      If Not rsPartmas.EOF And Not rsPartmas.BOF Then
         gconPMIOS.Execute "update partmas set lastm_mac = " & N2Str2Zero(rsPartmasBackup!lastm_mac) & _
                           " where partno = " & N2Str2Null(rsPartmasBackup!PartNo)
      Else
         txtMissingPartNo.Text = txtMissingPartNo.Text & UCase(Null2String(rsPartmasBackup!PartNo)) & vbCrLf
      End If
      i = i + 1
      progCPB.Value = (i / rsPartmasBackup.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsPartmasBackup.MoveNext
   Loop
End If
End Sub



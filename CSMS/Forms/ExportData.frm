VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmCSMIOSExportData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATA TRANSFER FROM UNFORMATTED TO STANDARD FORMAT"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "ExportData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1470
   ScaleWidth      =   5715
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   765
      Left            =   4740
      Picture         =   "ExportData.frx":0442
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
      Picture         =   "ExportData.frx":074C
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
      Begin wizProgBar.Prg progTransfer_Data 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "ExportData.frx":0A56
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "ExportData.frx":0A72
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
         TabIndex        =   6
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   7
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
            MICON           =   "ExportData.frx":0A8E
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
         TabIndex        =   8
         Top             =   30
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frmCSMIOSExportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gconOLDCSMIOS As ADODB.Connection

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPost_Click()
cmdPost.Enabled = False
cmdExit.Enabled = False
TRANSFER_DATA
On Error Resume Next
gconOLDCSMIOS.Close
cmdExit.Enabled = True
End Sub

Private Sub Form_Load()
CenterMe frmMain, Me, 1
End Sub

Private Function OpenOldDb() As Boolean
Dim OLDCSMIOS_Connection As String
With wizVar
     If .VerifyCryptoFile(App.Path & "\CSMIOS.crp") = True Then
        OLDCSMIOS_Connection = .OpenCryptoFile("OLDCSMIOS", "CONNECT")
     End If
End With
On Error Resume Next
deOLDCSMIOS.deConnOLDCSMIOS.Close
On Error GoTo ConnErr
If OLDCSMIOS_Connection <> "" Then
   deOLDCSMIOS.deConnOLDCSMIOS.ConnectionString = OLDCSMIOS_Connection
   Set gconOLDCSMIOS = New ADODB.Connection
   Set gconOLDCSMIOS = deOLDCSMIOS.deConnOLDCSMIOS
   gconOLDCSMIOS.Open
   OpenOldDb = True
Else
   OpenOldDb = False
End If
Exit Function

ConnErr:
ShowADOErrors gconOLDCSMIOS
End Function

Sub TRANSFER_DATA()
If OpenOldDb Then
   MoveEsti_HD
   MoveEsti_Det
   MoveRepor
   MoveRo_Det
   MoveCusmas
   MoveCusVeh
   MoveROJOBS
   MoveJobMast
   MoveMATMAS
   MoveCLRMAS
   MoveCustCtl
   MoveS_Model
   MoveEmpNo
   MoveInvFlag
Else
   MsgBoxXP "Cannot Find Old Database File"
End If
End Sub

Sub MoveCLRMAS()
Dim MoveSql As String
Dim i As Integer

Dim varVSCODE, varVSCOLOR As String

Dim rsOldCLRMAS As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from clrmas"
Set rsOldCLRMAS = New ADODB.Recordset
    rsOldCLRMAS.Open "select * from CLRMAS order by vscode asc", gconCSMIOS
If Not rsOldCLRMAS.EOF And Not rsOldCLRMAS.BOF Then
   rsOldCLRMAS.MoveFirst
   Me.Caption = "Currently Converting Color Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldCLRMAS.EOF
      varVSCODE = N2Str2Null(rsOldCLRMAS!vscode)
      varVSCOLOR = N2Str2Null(rsOldCLRMAS!vscolor)
      If varVSCODE <> "NULL" Then
         MoveSql = "INSERT INTO CLRMAS " & _
                   "(vscode,vscolor)" & _
                   " values (" & varVSCODE & ", " & varVSCOLOR & ")"
         On Error GoTo ErrorCode
         gconOLDCSMIOS.Execute MoveSql
      End If
      i = i + 1
      progTransfer_Data.Value = (i / rsOldCLRMAS.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldCLRMAS.MoveNext
   Loop
   Me.Caption = "Color Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveCustCtl()
Dim MoveSql As String
Dim i As Integer

Dim varCustCtlcde As String
Dim varCustCtldsc As String

Dim rsOldCustCtl As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from cusctl"
Set rsOldCustCtl = New ADODB.Recordset
    rsOldCustCtl.Open "select * from CusCtl order by ctlcde asc", gconCSMIOS
If Not rsOldCustCtl.EOF And Not rsOldCustCtl.BOF Then
   rsOldCustCtl.MoveFirst
   Me.Caption = "Currently Converting Cusmas Control Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldCustCtl.EOF
      varCustCtlcde = N2Str2Null(rsOldCustCtl!ctlcde)
      varCustCtldsc = N2Str2Null(rsOldCustCtl!ctldsc)
      
      MoveSql = "INSERT INTO cusctl " & _
                "(ctlcde,ctldsc)" & _
                " values (" & varCustCtlcde & ", " & varCustCtldsc & ")"
      On Error GoTo ErrorCode
      gconOLDCSMIOS.Execute MoveSql
      i = i + 1
      progTransfer_Data.Value = (i / rsOldCustCtl.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldCustCtl.MoveNext
   Loop
   Me.Caption = "Cusmas Control Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveCusmas()
Dim MoveSql As String
Dim i As Integer

Dim varCustCuscde, varCustCusComp, varCustCusnam As String
Dim varCustCusnam1, varCustCusnam2, varCustCusnam3 As String
Dim varCustCusadd1, varCustCusadd2, varCustCusadd3 As String
Dim varCustCRAmount As Double
Dim varCustCuszipc, varCustCusphon1, varCustCuscat As String
Dim varCustUserCode, varCustLastUpdate, varCustTimeUpdate As String
Dim varCustUserCode2, varCustEditDate, varCustEditTime As String
Dim varCustOldCode, varCustCusType, varCustPlateNo As String

Dim rsOldCusmas, rsOldCusVeh As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from cusmas"
Set rsOldCusmas = New ADODB.Recordset
    rsOldCusmas.Open "select * from Cusmas order by Cuscde asc", gconCSMIOS
If Not rsOldCusmas.EOF And Not rsOldCusmas.BOF Then
   rsOldCusmas.MoveFirst
   Me.Caption = "Currently Converting Customer Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Dim CusNamLent As Integer
   Do While Not rsOldCusmas.EOF
      varCustCuscde = N2Str2Null(Trim(Left(rsOldCusmas!Cuscde, 8)))
      varCustCusComp = N2Str2Null(Trim(Left(rsOldCusmas!Cuscomp, 40)))
      varCustCusnam = ""
      For CusNamLent = 1 To 30
          If Mid(Null2String(rsOldCusmas!cusnam), CusNamLent, 1) <> "/" Then
             varCustCusnam = varCustCusnam & Mid(Null2String(rsOldCusmas!cusnam), CusNamLent, 1)
          Else
             Exit For
          End If
      Next
      varCustCusnam = N2Str2Null(varCustCusnam)
      varCustCusnam1 = N2Str2Null(UCase(Null2String(rsOldCusmas!Lastname)))
      varCustCusnam2 = N2Str2Null(UCase(Null2String(rsOldCusmas!Firstname)))
      varCustCusnam3 = N2Str2Null(UCase(Null2String(rsOldCusmas!MiddleInitial)))
      varCustCusadd1 = N2Str2Null(Trim(Left(rsOldCusmas!Cusadd, 40)))
      varCustCusadd2 = N2Str2Null(Trim(Mid(rsOldCusmas!Cusadd, 41, 40)))
      varCustCusadd3 = N2Str2Null(Trim(Left(rsOldCusmas!Provadd, 30)))
      varCustCRAmount = N2Str2IntZero(rsOldCusmas!CRamount)
      varCustCuszipc = N2Str2Null(rsOldCusmas!cuszipc)
      varCustCusphon1 = N2Str2Null(rsOldCusmas!cusphon1)
      varCustCuscat = N2Str2Null(rsOldCusmas!cuscat)
      varCustUserCode = N2Str2Null(rsOldCusmas!usercode)
      varCustLastUpdate = N2Date2Null(rsOldCusmas!lastupdate)
      varCustTimeUpdate = N2Str2Null(rsOldCusmas!timeupdate)
      varCustUserCode2 = N2Str2Null(rsOldCusmas!usercode2)
      varCustEditDate = N2Date2Null(rsOldCusmas!editdate)
      varCustEditTime = N2Str2Null(rsOldCusmas!edittime)
      varCustOldCode = N2Str2Null(rsOldCusmas!oldcode)
      varCustCusType = N2Str2Null(rsOldCusmas!Custype)
      
      MoveSql = "INSERT INTO Cusmas " & _
                "(cuscde,Cuscomp,cusnam,cusnam1,cusnam2,cusnam3,cusadd1,cusadd2,cusadd3,cramount,cuszipc,cusphon1,cuscat,usercode,lastupdate,timeupdate,usercode2,editdate,edittime,oldcode,custype)" & _
                " values (" & varCustCuscde & ", " & varCustCusComp & ", " & varCustCusnam & ", " & varCustCusnam1 & ", " & varCustCusnam2 & ", " & varCustCusnam3 & ", " & varCustCusadd1 & ", " & varCustCusadd2 & ", " & varCustCusadd3 & ", " & varCustCRAmount & ", " & varCustCuszipc & ", " & varCustCusphon1 & ", " & varCustCuscat & ", " & varCustUserCode & ", " & varCustLastUpdate & ", " & varCustTimeUpdate & ", " & varCustUserCode2 & ", " & varCustEditDate & ", " & varCustEditTime & ", " & varCustOldCode & ", " & varCustCusType & ")"
      On Error GoTo ErrorCode
      gconOLDCSMIOS.Execute MoveSql
      i = i + 1
      progTransfer_Data.Value = (i / rsOldCusmas.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldCusmas.MoveNext
   Loop
   Me.Caption = "Customer Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveCusVeh()
Dim MoveSql As String
Dim i As Integer

Dim varCustCuscde, varCustNiym, varCustPlateNo As String
Dim varCustVCond_No, varCustClrCde, varCustModel As String
Dim varCustEngine, varCustProdNo, varCustSerial As String
Dim varCustTin_Number, varCustD_Sold, varCustWar_Cert As String
Dim varCustDel_Date As String

Dim rsOldCusVeh As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from cusveh"
Set rsOldCusVeh = New ADODB.Recordset
    rsOldCusVeh.Open "select * from Cusveh order by Cuscde asc", gconCSMIOS
If Not rsOldCusVeh.EOF And Not rsOldCusVeh.BOF Then
   rsOldCusVeh.MoveFirst
   Me.Caption = "Currently Converting Customer Vehicle Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
      Dim CusNiymLent As Integer
      Do While Not rsOldCusVeh.EOF
      varCustCuscde = N2Str2Null(rsOldCusVeh!Cuscde)
      varCustNiym = ""
      For CusNiymLent = 1 To 30
          If Mid(Null2String(rsOldCusVeh!Niym), CusNiymLent, 1) <> "/" Then
             varCustNiym = varCustNiym & Mid(Null2String(rsOldCusVeh!Niym), CusNiymLent, 1)
          Else
             Exit For
          End If
      Next
      varCustNiym = N2Str2Null(varCustNiym)
      varCustPlateNo = N2Str2Null(rsOldCusVeh!plate_no)
      varCustVCond_No = N2Str2Null(rsOldCusVeh!vcond_no)
      varCustClrCde = N2Str2Null(rsOldCusVeh!clrcde)
      varCustModel = N2Str2Null(rsOldCusVeh!model)
      varCustEngine = N2Str2Null(rsOldCusVeh!engine)
      varCustProdNo = N2Str2Null(rsOldCusVeh!prodno)
      varCustSerial = N2Str2Null(rsOldCusVeh!serial)
      varCustTin_Number = N2Str2Null(rsOldCusVeh!tin_number)
      varCustD_Sold = N2Str2Null(rsOldCusVeh!d_sold)
      varCustWar_Cert = N2Str2Null(rsOldCusVeh!war_cert)
      varCustDel_Date = N2Date2Null(rsOldCusVeh!del_date)
      
      If varCustPlateNo <> "NULL" Then
         MoveSql = "INSERT INTO Cusveh " & _
                   "(cuscde,[name],plate_no,vcond_no,clrcde,model,engine,prodno,serial,tin_number,d_sold,war_cert,del_date)" & _
                   " values (" & varCustCuscde & ", " & varCustNiym & ", " & varCustPlateNo & ", " & varCustVCond_No & ", " & varCustClrCde & ", " & varCustModel & ", " & varCustEngine & ", " & varCustProdNo & ", " & varCustSerial & ", " & varCustTin_Number & ", " & varCustD_Sold & ", " & varCustWar_Cert & ", " & varCustDel_Date & ")"
         On Error GoTo ErrorCode
         gconOLDCSMIOS.Execute MoveSql
      End If
      i = i + 1
      progTransfer_Data.Value = (i / rsOldCusVeh.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldCusVeh.MoveNext
   Loop
   Me.Caption = "Customer Vehicle Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveRo_Det()
Dim MoveSql As String
Dim i As Integer

Dim varREP_OR, varLEVEL, varLINE_NO, varDETCDE As String
Dim varDETDSC, varDETUNT As String
Dim varDETVOL, varDETPRC, varDETAMT As Double
Dim varCODE, varWCODE As String
Dim varTAXRATE, varDISCRATE, varTAXVAL, varDISVAL As Double
Dim varPOCODE, varREP_OR2, varDETAIL As String
Dim varDETAIL1, varDETAIL2, varDETAIL3 As String
Dim varDET_AMT, varDIS_VAL, varDISCOUNT_2 As Double
Dim rsRO_DET As ADODB.Recordset
Dim rsOldRo_Det As ADODB.Recordset
Dim rsRemarks As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from ro_det"
gconOLDCSMIOS.Execute "delete * from remarks"
Set rsOldRo_Det = New ADODB.Recordset
    rsOldRo_Det.Open "select * from ro_det WHERE DEALER_TYPE = " & DEALER_TYPE & " order by rep_or desc", gconCSMIOS
If Not rsOldRo_Det.EOF And Not rsOldRo_Det.BOF Then
   rsOldRo_Det.MoveFirst
   Me.Caption = "Currently Converting Repair Order Details File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldRo_Det.EOF
      varREP_OR = N2Str2Null(rsOldRo_Det!rep_or)
      varLEVEL = N2Str2Null(rsOldRo_Det!livil)
      varLINE_NO = N2Str2Null(rsOldRo_Det!line_no)
      varDETCDE = N2Str2Null(rsOldRo_Det!detcde)
      varDETDSC = N2Str2Null(rsOldRo_Det!detdsc)
      varDETUNT = N2Str2Null(rsOldRo_Det!detunt)
      varDETVOL = N2Str2Zero(rsOldRo_Det!detvol)
      varDETPRC = N2Str2Zero(rsOldRo_Det!detprc)
      varDETAMT = N2Str2Zero(rsOldRo_Det!detamt)
      varCODE = N2Str2Null(rsOldRo_Det!Code)
      varWCODE = N2Str2Null(rsOldRo_Det!wCode)
      varTAXRATE = N2Str2Zero(rsOldRo_Det!taxrate)
      varDISCRATE = N2Str2Zero(rsOldRo_Det!discrate)
      varTAXVAL = N2Str2Zero(rsOldRo_Det!taxval)
      varDISVAL = N2Str2Zero(rsOldRo_Det!disval)
      varPOCODE = N2Str2Null(rsOldRo_Det!pocode)
      varREP_OR2 = N2Str2Null(rsOldRo_Det!Rep_Or2)
      varDETAIL = N2Str2Null(Trim(rsOldRo_Det!detail))
      If Null2String(rsOldRo_Det!detail) <> "" Then
         varDETAIL1 = N2Str2Null(Trim(Left(rsOldRo_Det!detail, 79)))
         varDETAIL2 = N2Str2Null(Trim(Mid(rsOldRo_Det!detail, 80, 79)))
         varDETAIL3 = N2Str2Null(Trim(Mid(rsOldRo_Det!detail, 159, 79)))
         gconOLDCSMIOS.Execute "insert into remarks " & _
                               "(REP_OR,[LEVEL],[LINENO],REMARKS1,REMARKS2,REMARKS3) " & _
                               "values (" & varREP_OR & ", " & varLEVEL & ", " & varLINE_NO & _
                               ", " & varDETAIL1 & ", " & varDETAIL2 & ", " & varDETAIL3 & ")"
      End If
      varDET_AMT = N2Str2Zero(rsOldRo_Det!det_amt)
      varDIS_VAL = N2Str2Zero(rsOldRo_Det!dis_val)
      varDISCOUNT_2 = N2Str2Zero(rsOldRo_Det!discount_2)
      MoveSql = "INSERT INTO ro_det " & _
                "(REP_OR,[LEVEL],[LINENO],DETCDE,DETDSC,DETUNT,DETVOL,DETPRC,DETAMT,CODE,WCODE,TAXRATE,DISCRATE,TAXVAL,DISVAL,POCODE,REP_OR2,DET_AMT,DIS_VAL,DISCOUNT_2)" & _
                " values (" & varREP_OR & ", " & varLEVEL & ", " & varLINE_NO & ", " & varDETCDE & ", " & varDETDSC & ", " & varDETUNT & ", " & varDETVOL & ", " & varDETPRC & ", " & varDETAMT & ", " & varCODE & ", " & varWCODE & ", " & varTAXRATE & ", " & varDISCRATE & ", " & varTAXVAL & ", " & varDISVAL & ", " & varPOCODE & ", " & varREP_OR2 & ", " & varDET_AMT & ", " & varDIS_VAL & ", " & varDISCOUNT_2 & ")"
      gconOLDCSMIOS.Execute MoveSql
      i = i + 1
      progTransfer_Data.Value = (i / rsOldRo_Det.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldRo_Det.MoveNext
   Loop
   Me.Caption = "Repair Order Details File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveEmpNo()
Dim MoveSql As String
Dim i As Integer

Dim varCODE, varLASTNAME, varFIRSTNAME As String
Dim varMIDDLEINT, varNAME, varSTATUSCD As String
Dim varCLASSCD, varPOSITION, varSEXCD As String
Dim varCIVILSTAT As String
Dim varBASIC_PAY As Double
Dim varACCTNO, varEMPNO As String

Dim rsOldEmpNo As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from empno"
Set rsOldEmpNo = New ADODB.Recordset
    rsOldEmpNo.Open "select * from EmpNo order by code asc", gconCSMIOS
If Not rsOldEmpNo.EOF And Not rsOldEmpNo.BOF Then
   rsOldEmpNo.MoveFirst
   Me.Caption = "Currently Converting Service Advisor Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldEmpNo.EOF
      varCODE = N2Str2Null(rsOldEmpNo!Code)
      varLASTNAME = N2Str2Null(UCase(Null2String(rsOldEmpNo!Lastname)))
      varFIRSTNAME = N2Str2Null(UCase(Null2String(rsOldEmpNo!Firstname)))
      varMIDDLEINT = N2Str2Null(UCase(Null2String(rsOldEmpNo!middleint)))
      varNAME = N2Str2Null(UCase(Null2String(rsOldEmpNo!naym)))
      varSTATUSCD = N2Str2Null(rsOldEmpNo!statuscd)
      varCLASSCD = N2Str2Null(rsOldEmpNo!classcd)
      varPOSITION = N2Str2Null(rsOldEmpNo!Positions)
      varSEXCD = N2Str2Null(rsOldEmpNo!sexcd)
      varCIVILSTAT = N2Str2Null(rsOldEmpNo!civilstat)
      varBASIC_PAY = N2Str2Zero(rsOldEmpNo!basic_pay)
      varACCTNO = N2Str2Null(rsOldEmpNo!acctno)
      varEMPNO = N2Str2Null(rsOldEmpNo!empno)

      MoveSql = "INSERT INTO empno " & _
                "(code,lastname,firstname,middleint,[name],statuscd,classcd,[position],sexcd,civilstat,basic_pay,acctno,empno)" & _
                " values (" & varCODE & ", " & varLASTNAME & ", " & varFIRSTNAME & ", " & varMIDDLEINT & ", " & varNAME & ", " & varSTATUSCD & ", " & varCLASSCD & ", " & varPOSITION & ", " & varSEXCD & ", " & varCIVILSTAT & ", " & varBASIC_PAY & ", " & varACCTNO & ", " & varEMPNO & ")"
      On Error GoTo ErrorCode
      gconOLDCSMIOS.Execute MoveSql
      i = i + 1
      progTransfer_Data.Value = (i / rsOldEmpNo.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldEmpNo.MoveNext
   Loop
   Me.Caption = "Service Advisor Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveEsti_Det()
Dim MoveSql As String
Dim i As Integer

Dim varESTIMATENO, varLEVEL, varLINE_NO As String
Dim varDETCDE, varDETDSC, varDETUNT As String
Dim varDETVOL, varDETPRC, varDETAMT As Double
Dim varCODE, varWCODE As String
Dim varTAXRATE, varDISCRATE, varTAXVAL As Double
Dim varDISVAL As Double
Dim varPOCODE, varREP_OR2, varDETAIL As String
Dim varDETAIL1, varDETAIL2, varDETAIL3 As String
Dim varDET_AMT, varDIS_VAL, varDISCOUNT_2 As Double

Dim rsOldESTI_Det As ADODB.Recordset
Dim rsRemarks As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from esti_det"
Set rsOldESTI_Det = New ADODB.Recordset
    rsOldESTI_Det.Open "select * from ESTI_Det WHERE DEALER_TYPE = " & DEALER_TYPE, gconCSMIOS
If Not rsOldESTI_Det.EOF And Not rsOldESTI_Det.BOF Then
   rsOldESTI_Det.MoveFirst
   Me.Caption = "Currently Converting Estimate Details File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldESTI_Det.EOF
      varESTIMATENO = N2Str2Null(rsOldESTI_Det!EstimateNo)
      If Len(varESTIMATENO) = 8 Then GoTo 1
      varLEVEL = N2Str2Null(rsOldESTI_Det!livil)
      varLINE_NO = N2Str2Null(rsOldESTI_Det!line_no)
      varDETCDE = N2Str2Null(rsOldESTI_Det!detcde)
      varDETDSC = N2Str2Null(rsOldESTI_Det!detdsc)
      varDETUNT = N2Str2Null(rsOldESTI_Det!detunt)
      varDETVOL = N2Str2Zero(rsOldESTI_Det!detvol)
      varDETPRC = N2Str2Zero(rsOldESTI_Det!detprc)
      varDETAMT = N2Str2Zero(rsOldESTI_Det!detamt)
      varCODE = N2Str2Null(rsOldESTI_Det!Code)
      varWCODE = N2Str2Null(rsOldESTI_Det!wCode)
      varTAXRATE = N2Str2Zero(rsOldESTI_Det!taxrate)
      varDISCRATE = N2Str2Zero(rsOldESTI_Det!discrate)
      varTAXVAL = N2Str2Zero(rsOldESTI_Det!taxval)
      varDISVAL = N2Str2Zero(rsOldESTI_Det!disval)
      varPOCODE = N2Str2Null(rsOldESTI_Det!pocode)
      varREP_OR2 = N2Str2Null(rsOldESTI_Det!Rep_Or2)
      varDETAIL = N2Str2Null(Trim(rsOldESTI_Det!detail))
      If Null2String(rsOldESTI_Det!detail) <> "" Then
         varDETAIL1 = N2Str2Null(Trim(Left(rsOldESTI_Det!detail, 79)))
         varDETAIL2 = N2Str2Null(Trim(Mid(rsOldESTI_Det!detail, 80, 79)))
         varDETAIL3 = N2Str2Null(Trim(Mid(rsOldESTI_Det!detail, 159, 79)))
         gconOLDCSMIOS.Execute "insert into esti_rem " & _
                               "(ESTIMATENO,[LEVEL],[LINENO],REMARKS1,REMARKS2,REMARKS3) " & _
                               "values (" & varESTIMATENO & ", " & varLEVEL & ", " & varLINE_NO & _
                               ", " & varDETAIL1 & ", " & varDETAIL2 & ", " & varDETAIL3 & ")"
      End If
      varDET_AMT = N2Str2Zero(rsOldESTI_Det!det_amt)
      varDIS_VAL = N2Str2Zero(rsOldESTI_Det!dis_val)
      varDISCOUNT_2 = N2Str2Zero(rsOldESTI_Det!discount_2)
   
      MoveSql = "INSERT INTO esti_det " & _
                "(ESTIMATENO,[LEVEL],[LINENO],DETCDE,DETDSC,DETUNT,DETVOL,DETPRC,DETAMT,CODE,WCODE,TAXRATE,DISCRATE,TAXVAL,DISVAL,POCODE,REP_OR2,DETAIL,DET_AMT,DIS_VAL,DISCOUNT_2)" & _
                " values (" & varESTIMATENO & ", " & varLEVEL & ", " & varLINE_NO & ", " & varDETCDE & ", " & varDETDSC & ", " & varDETUNT & ", " & varDETVOL & ", " & varDETPRC & ", " & varDETAMT & ", " & varCODE & ", " & varWCODE & ", " & varTAXRATE & ", " & varDISCRATE & ", " & varTAXVAL & ", " & varDISVAL & ", " & varPOCODE & ", " & varREP_OR2 & ", " & varDETAIL & ", " & varDET_AMT & ", " & varDIS_VAL & ", " & varDISCOUNT_2 & ")"
      On Error GoTo ErrorCode
      gconOLDCSMIOS.Execute MoveSql
1      i = i + 1
      progTransfer_Data.Value = (i / rsOldESTI_Det.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldESTI_Det.MoveNext
   Loop
   Me.Caption = "Estimate Details File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveEsti_HD()
Dim MoveSql As String
Dim i As Integer

Dim varREP_OR, varESTIMATENO, varROTYPE, varSVC_NO As String
Dim varACCT_NO, varINSCDE, varNAME, varPLATE_NO As String
Dim varMODEL, varTERM, varSECTION, varRECD_BY As String
Dim varKM_RDG As Double
Dim varDTE_RECD, varDTE_PRO, varCERTIFIC8 As String
Dim varAMOUNT, varINSAMT, varROVAT As Double
Dim VarLabor, VarParts, varMATERIAL, varL_DISC As Double
Dim varP_DISC, varM_DISC, varWL_AMT, varWP_AMT As Double
Dim varWM_AMT, varL_TAXVAL, varP_TAXVAL, varM_TAXVAL As Double
Dim varENT_DATE, varPRIN_DTE, varDTE_COMP, varDTE_REL As String
Dim varINVBAL As Double
Dim varSERVICE, varDOC_CODE As String
Dim varPARTSPEND As Boolean
Dim varUSERCODE, varSTATUS, varPREVSTAT, varSTATUS2 As String
Dim varSAVEDATE, varSAVETIME, varVERIFIED, varORNUM1 As String
Dim varORNUM2, varBANKCODE1, varBANKCODE2, varTERMS2 As String
Dim varPART_AMT As Double
Dim varPARTICIPAT, varCHECKNO1, varCHECKNO2 As String
Dim varCHECKDATE1, varCHECKDATE2 As String
Dim varCASHAMT1, varCREDITAMT1, varCASHAMT2, varCREDITAMT2 As Double
Dim varREF_NO1, varREF_NO2 As String
Dim varCHECKAMT1, varCHECKAMT2 As Double
Dim varCLCODE1, varCLCODE2 As String
Dim varCLAMT1, varCLAMT2 As Double
Dim varORDATE1, varORDATE2 As String
Dim varPTAG As Boolean
Dim varINVOICE As String
Dim varL_DISC2, varP_DISC2, varM_DISC2 As Double
Dim varL_DISCOUNT, varP_DISCOUNT, varM_DISCOUNT As Double
Dim varL_AMTVALUE, varP_AMTVALUE, varM_AMTVALUE As Double
Dim varRO_AMOUNT As Double

Dim rsOldESTI_HD As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from esti_hd"
Set rsOldESTI_HD = New ADODB.Recordset
    rsOldESTI_HD.Open "select * from ESTI_HD WHERE DEALER_TYPE = " & DEALER_TYPE, gconCSMIOS
If Not rsOldESTI_HD.EOF And Not rsOldESTI_HD.BOF Then
   rsOldESTI_HD.MoveFirst
   Me.Caption = "Currently Converting Estimate Header File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldESTI_HD.EOF
      varREP_OR = N2Str2Null(rsOldESTI_HD!rep_or)
      varESTIMATENO = N2Str2Null(rsOldESTI_HD!EstimateNo)
      If Len(varESTIMATENO) = 8 Then GoTo 1
      varROTYPE = N2Str2Null(rsOldESTI_HD!rotype)
      varSVC_NO = N2Str2Null(rsOldESTI_HD!svc_no)
      varACCT_NO = N2Str2Null(rsOldESTI_HD!acct_no)
      varINSCDE = N2Str2Null(rsOldESTI_HD!inscde)
      varNAME = N2Str2Null(Trim(Left(UCase(Null2String(rsOldESTI_HD!Niym)), 60)))
      varPLATE_NO = N2Str2Null(rsOldESTI_HD!plate_no)
      varMODEL = N2Str2Null(rsOldESTI_HD!model)
      varTERM = N2Str2Null(rsOldESTI_HD!term)
      varSECTION = N2Str2Null(rsOldESTI_HD!sektion)
      varRECD_BY = N2Str2Null(rsOldESTI_HD!recd_by)
      varKM_RDG = N2Str2IntZero(rsOldESTI_HD!km_rdg)
      varDTE_RECD = N2Date2Null(rsOldESTI_HD!dte_recd)
      varDTE_PRO = N2Date2Null(rsOldESTI_HD!dte_pro)
      varCERTIFIC8 = N2Str2Null(rsOldESTI_HD!certific8)
      varAMOUNT = N2Str2Zero(rsOldESTI_HD!amount)
      varINSAMT = N2Str2Zero(rsOldESTI_HD!insamt)
      varROVAT = N2Str2Zero(rsOldESTI_HD!rovat)
      VarLabor = N2Str2Zero(rsOldESTI_HD!labor)
      VarParts = N2Str2Zero(rsOldESTI_HD!parts)
      varMATERIAL = N2Str2Zero(rsOldESTI_HD!material)
      varL_DISC = N2Str2Zero(rsOldESTI_HD!l_disc)
      varP_DISC = N2Str2Zero(rsOldESTI_HD!p_disc)
      varM_DISC = N2Str2Zero(rsOldESTI_HD!m_disc)
      varWL_AMT = N2Str2Zero(rsOldESTI_HD!wl_amt)
      varWP_AMT = N2Str2Zero(rsOldESTI_HD!wp_amt)
      varWM_AMT = N2Str2Zero(rsOldESTI_HD!wm_amt)
      varL_TAXVAL = N2Str2Zero(rsOldESTI_HD!l_taxval)
      varP_TAXVAL = N2Str2Zero(rsOldESTI_HD!p_taxval)
      varM_TAXVAL = N2Str2Zero(rsOldESTI_HD!m_taxval)
      varENT_DATE = N2Date2Null(rsOldESTI_HD!ent_date)
      varPRIN_DTE = N2Date2Null(rsOldESTI_HD!prin_dte)
      varDTE_COMP = N2Date2Null(rsOldESTI_HD!dte_comp)
      varDTE_REL = N2Date2Null(rsOldESTI_HD!dte_rel)
      varINVBAL = N2Str2Zero(rsOldESTI_HD!invbal)
      varSERVICE = N2Str2Null(rsOldESTI_HD!service)
      varDOC_CODE = N2Str2Null(rsOldESTI_HD!doc_code)
      varUSERCODE = N2Str2Null(rsOldESTI_HD!usercde)
      varSTATUS = N2Str2Null(rsOldESTI_HD!Status)
      varPREVSTAT = N2Str2Null(rsOldESTI_HD!prevstat)
      varSTATUS2 = N2Str2Null(rsOldESTI_HD!status2)
      varSAVEDATE = N2Date2Null(LOGDATE)
      varSAVETIME = N2Str2Null(LOGTIME)
      varVERIFIED = N2Str2Null(rsOldESTI_HD!verified)
      varORNUM1 = N2Str2Null(rsOldESTI_HD!ornum1)
      varORNUM2 = N2Str2Null(rsOldESTI_HD!ornum2)
      varBANKCODE1 = N2Str2Null(rsOldESTI_HD!bankcode1)
      varBANKCODE2 = N2Str2Null(rsOldESTI_HD!bankcode2)
      varTERMS2 = N2Str2Null(rsOldESTI_HD!terms2)
      varPART_AMT = N2Str2Zero(rsOldESTI_HD!part_amt)
      varPARTICIPAT = N2Str2Null(rsOldESTI_HD!participat)
      varCHECKNO1 = N2Str2Null(rsOldESTI_HD!checkno1)
      varCHECKNO2 = N2Str2Null(rsOldESTI_HD!checkno2)
      varCHECKDATE1 = N2Date2Null(rsOldESTI_HD!checkdate1)
      varCHECKDATE2 = N2Date2Null(rsOldESTI_HD!checkdate2)
      varCASHAMT1 = N2Str2Zero(rsOldESTI_HD!cashamt1)
      varCREDITAMT1 = N2Str2Zero(rsOldESTI_HD!creditamt1)
      varCASHAMT2 = N2Str2Zero(rsOldESTI_HD!cashamt2)
      varCREDITAMT2 = N2Str2Zero(rsOldESTI_HD!creditamt2)
      varREF_NO1 = N2Str2Null(rsOldESTI_HD!ref_no1)
      varREF_NO2 = N2Str2Null(rsOldESTI_HD!ref_no2)
      varCHECKAMT1 = N2Str2Zero(rsOldESTI_HD!checkamt1)
      varCHECKAMT2 = N2Str2Zero(rsOldESTI_HD!checkamt2)
      varCLCODE1 = N2Str2Null(rsOldESTI_HD!clcode1)
      varCLCODE2 = N2Str2Null(rsOldESTI_HD!clcode2)
      varCLAMT1 = N2Str2Zero(rsOldESTI_HD!clamt1)
      varCLAMT2 = N2Str2Zero(rsOldESTI_HD!clamt2)
      varORDATE1 = N2Str2Null(rsOldESTI_HD!ordate1)
      varORDATE2 = N2Str2Null(rsOldESTI_HD!ordate2)
      varINVOICE = N2Str2Null(rsOldESTI_HD!invoice)
      varL_DISC2 = N2Str2Zero(rsOldESTI_HD!l_disc2)
      varP_DISC2 = N2Str2Zero(rsOldESTI_HD!p_disc2)
      varM_DISC2 = N2Str2Zero(rsOldESTI_HD!m_disc2)
      varL_DISCOUNT = N2Str2Zero(rsOldESTI_HD!l_discount)
      varP_DISCOUNT = N2Str2Zero(rsOldESTI_HD!p_discount)
      varM_DISCOUNT = N2Str2Zero(rsOldESTI_HD!m_discount)
      varL_AMTVALUE = N2Str2Zero(rsOldESTI_HD!l_amtvalue)
      varP_AMTVALUE = N2Str2Zero(rsOldESTI_HD!p_amtvalue)
      varM_AMTVALUE = N2Str2Zero(rsOldESTI_HD!m_amtvalue)
      varRO_AMOUNT = N2Str2Zero(rsOldESTI_HD!ro_amount)
       
      MoveSql = "INSERT INTO esti_hd " & _
                "(REP_OR,ESTIMATENO,ROTYPE,SVC_NO,ACCT_NO,INSCDE,[NAME],PLATE_NO,MODEL,TERM,[SECTION],RECD_BY,KM_RDG,DTE_RECD,DTE_PRO,CERTIFIC8," & _
                "AMOUNT,INSAMT,ROVAT,LABOR,PARTS,MATERIAL,L_DISC,P_DISC,M_DISC,WL_AMT,WP_AMT,WM_AMT,L_TAXVAL,P_TAXVAL,M_TAXVAL,ENT_DATE,PRIN_DTE," & _
                "DTE_COMP,DTE_REL,INVBAL,SERVICE,DOC_CODE,USERCDE,STATUS,PREVSTAT,STATUS2,SAVEDATE,SAVETIME,VERIFIED,ORNUM1,ORNUM2,BANKCODE1,BANKCODE2,TERMS2," & _
                "PART_AMT,PARTICIPAT,CHECKNO1,CHECKNO2,CHECKDATE1,CHECKDATE2,CASHAMT1,CREDITAMT1,CASHAMT2,CREDITAMT2,REF_NO1,REF_NO2,CHECKAMT1,CHECKAMT2,CLCODE1,CLCODE2,CLAMT1,CLAMT2,ORDATE1," & _
                "ORDATE2,INVOICE,L_DISC2,P_DISC2,M_DISC2,L_DISCOUNT,P_DISCOUNT,M_DISCOUNT,L_AMTVALUE,P_AMTVALUE,M_AMTVALUE,RO_AMOUNT)" & _
                " values (" & varREP_OR & "," & varESTIMATENO & "," & varROTYPE & "," & varSVC_NO & "," & varACCT_NO & "," & varINSCDE & "," & varNAME & "," & varPLATE_NO & "," & varMODEL & "," & varTERM & "," & varSECTION & "," & varRECD_BY & "," & varKM_RDG & "," & varDTE_RECD & "," & varDTE_PRO & "," & varCERTIFIC8 & _
                "," & varAMOUNT & "," & varINSAMT & "," & varROVAT & "," & VarLabor & "," & VarParts & "," & varMATERIAL & "," & varL_DISC & "," & varP_DISC & "," & varM_DISC & "," & varWL_AMT & "," & varWP_AMT & "," & varWM_AMT & "," & varL_TAXVAL & "," & varP_TAXVAL & "," & varM_TAXVAL & "," & varENT_DATE & "," & varPRIN_DTE & _
                "," & varDTE_COMP & "," & varDTE_REL & "," & varINVBAL & "," & varSERVICE & "," & varDOC_CODE & "," & varUSERCODE & "," & varSTATUS & "," & varPREVSTAT & "," & varSTATUS2 & "," & varSAVEDATE & "," & varSAVETIME & "," & varVERIFIED & "," & varORNUM1 & "," & varORNUM2 & "," & varBANKCODE1 & "," & varBANKCODE2 & "," & varTERMS2 & _
                "," & varPART_AMT & "," & varPARTICIPAT & "," & varCHECKNO1 & "," & varCHECKNO2 & "," & varCHECKDATE1 & "," & varCHECKDATE2 & "," & varCASHAMT1 & "," & varCREDITAMT1 & "," & varCASHAMT2 & "," & varCREDITAMT2 & "," & varREF_NO1 & "," & varREF_NO2 & "," & varCHECKAMT1 & "," & varCHECKAMT2 & "," & varCLCODE1 & "," & varCLCODE2 & "," & varCLAMT1 & "," & varCLAMT2 & "," & varORDATE1 & _
                "," & varORDATE2 & "," & varINVOICE & "," & varL_DISC2 & "," & varP_DISC2 & "," & varM_DISC2 & "," & varL_DISCOUNT & "," & varP_DISCOUNT & "," & varM_DISCOUNT & "," & varL_AMTVALUE & "," & varP_AMTVALUE & "," & varM_AMTVALUE & "," & varRO_AMOUNT & ")"
      On Error GoTo ErrorCode
      gconOLDCSMIOS.Execute MoveSql
1      i = i + 1
      progTransfer_Data.Value = (i / rsOldESTI_HD.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldESTI_HD.MoveNext
   Loop
   Me.Caption = "Estimate Header File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveRepor()
Dim MoveSql As String
Dim i As Integer

Dim varREP_OR, varESTIMATENO, varROTYPE, varSVC_NO As String
Dim varACCT_NO, varINSCDE, varNAME, varPLATE_NO As String
Dim varMODEL, varTERM, varSECTION, varRECD_BY As String
Dim varKM_RDG As Double
Dim varDTE_RECD, varDTE_PRO, varCERTIFIC8 As String
Dim varAMOUNT, varINSAMT, varROVAT As Double
Dim VarLabor, VarParts, varMATERIAL As Double
Dim varL_DISC, varP_DISC, varM_DISC As Double
Dim varWL_AMT, varWP_AMT, varWM_AMT As Double
Dim varL_TAXVAL, varP_TAXVAL, varM_TAXVAL As Double
Dim varENT_DATE, varPRIN_DTE, varDTE_COMP, varDTE_REL As String
Dim varINVBAL As Double
Dim varSERVICE, varDOC_CODE As String
Dim varPARTSPEND As Boolean
Dim varUSERCODE, varSTATUS, varPREVSTAT As String
Dim varSTATUS2, varSAVEDATE, varSAVETIME As String
Dim varVERIFIED, varORNUM1, varORNUM2 As String
Dim varBANKCODE1, varBANKCODE2, varTERMS2 As String
Dim varPART_AMT As Double
Dim varPARTICIPAT, varCHECKNO1, varCHECKNO2 As String
Dim varCHECKDATE1, varCHECKDATE2 As String
Dim varCASHAMT1, varCREDITAMT1, varCASHAMT2 As Double
Dim varCREDITAMT2 As Double
Dim varREF_NO1, varREF_NO2 As String
Dim varCHECKAMT1, varCHECKAMT2 As Double
Dim varCLCODE1, varCLCODE2 As String
Dim varCLAMT1, varCLAMT2 As Double
Dim varORDATE1, varORDATE2 As String
Dim varPTAG As Boolean
Dim varINVOICE As String
Dim varL_DISC2, varP_DISC2, varM_DISC2 As Double
Dim varL_DISCOUNT, varP_DISCOUNT, varM_DISCOUNT As Double
Dim varL_AMTVALUE, varP_AMTVALUE, varM_AMTVALUE As Double
Dim varRO_AMOUNT, varVAT_EXEMPT, varDEPOSIT As Double
Dim varDEP_ORNUM, varDEP_DATE As String

Dim rsOldRepor As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from repor"
Set rsOldRepor = New ADODB.Recordset
    rsOldRepor.Open "select * from repor WHERE DEALER_TYPE = " & DEALER_TYPE & " order by id asc", gconCSMIOS
If Not rsOldRepor.EOF And Not rsOldRepor.BOF Then
   rsOldRepor.MoveFirst
   Me.Caption = "Currently Converting Repair Order Header File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Dim NiymLen As Integer
   Do While Not rsOldRepor.EOF
      varREP_OR = N2Str2Null(rsOldRepor!rep_or)
      varESTIMATENO = N2Str2Null(rsOldRepor!EstimateNo)
      varROTYPE = N2Str2Null(rsOldRepor!rotype)
      varSVC_NO = N2Str2Null(rsOldRepor!svc_no)
      varACCT_NO = N2Str2Null(rsOldRepor!acct_no)
      varINSCDE = N2Str2Null(rsOldRepor!inscde)
      varNAME = ""
      For NiymLen = 1 To 60
          If N2Str2Null(Mid(rsOldRepor!Niym, NiymLen, 1)) <> "/" Then
             varNAME = varNAME & Mid(Null2String(rsOldRepor!Niym), NiymLen, 1)
          Else
             Exit For
          End If
      Next
      varNAME = N2Str2Null(Trim(UCase(varNAME)))
      varPLATE_NO = N2Str2Null(rsOldRepor!plate_no)
      varMODEL = N2Str2Null(rsOldRepor!model)
      varTERM = N2Str2Null(rsOldRepor!term)
      varSECTION = N2Str2Null(rsOldRepor!sektion)
      varRECD_BY = N2Str2Null(rsOldRepor!recd_by)
      varKM_RDG = N2Str2IntZero(rsOldRepor!km_rdg)
      varDTE_RECD = N2Date2Null(rsOldRepor!dte_recd)
      varDTE_PRO = N2Date2Null(rsOldRepor!dte_pro)
      varCERTIFIC8 = N2Str2Null(rsOldRepor!certific8)
      varAMOUNT = N2Str2Zero(rsOldRepor!amount)
      varINSAMT = N2Str2Zero(rsOldRepor!insamt)
      varROVAT = N2Str2Zero(rsOldRepor!rovat)
      VarLabor = N2Str2Zero(rsOldRepor!labor)
      VarParts = N2Str2Zero(rsOldRepor!parts)
      varMATERIAL = N2Str2Zero(rsOldRepor!material)
      varL_DISC = N2Str2Zero(rsOldRepor!l_disc)
      varP_DISC = N2Str2Zero(rsOldRepor!p_disc)
      varM_DISC = N2Str2Zero(rsOldRepor!m_disc)
      varWL_AMT = N2Str2Zero(rsOldRepor!wl_amt)
      varWP_AMT = N2Str2Zero(rsOldRepor!wp_amt)
      varWM_AMT = N2Str2Zero(rsOldRepor!wm_amt)
      varL_TAXVAL = N2Str2Zero(rsOldRepor!l_taxval)
      varP_TAXVAL = N2Str2Zero(rsOldRepor!p_taxval)
      varM_TAXVAL = N2Str2Zero(rsOldRepor!m_taxval)
      varENT_DATE = N2Date2Null(rsOldRepor!ent_date)
      varPRIN_DTE = N2Date2Null(rsOldRepor!prin_dte)
      varDTE_COMP = N2Date2Null(rsOldRepor!dte_comp)
      varDTE_REL = N2Date2Null(rsOldRepor!dte_rel)
      varINVBAL = N2Str2Zero(rsOldRepor!invbal)
      varSERVICE = N2Str2Null(rsOldRepor!service)
      varDOC_CODE = N2Str2Null(rsOldRepor!doc_code)
      varUSERCODE = N2Str2Null(rsOldRepor!usercde)
      varSTATUS = N2Str2Null(rsOldRepor!Status)
      varPREVSTAT = N2Str2Null(rsOldRepor!prevstat)
      varSTATUS2 = N2Str2Null(rsOldRepor!status2)
      varSAVEDATE = N2Date2Null(LOGDATE)
      varSAVETIME = N2Str2Null(LOGTIME)
      varVERIFIED = N2Str2Null(rsOldRepor!verified)
      varORNUM1 = N2Str2Null(rsOldRepor!ornum1)
      varORNUM2 = N2Str2Null(rsOldRepor!ornum2)
      varBANKCODE1 = N2Str2Null(rsOldRepor!bankcode1)
      varBANKCODE2 = N2Str2Null(rsOldRepor!bankcode2)
      varTERMS2 = N2Str2Null(rsOldRepor!terms2)
      varPART_AMT = N2Str2Zero(rsOldRepor!part_amt)
      varPARTICIPAT = N2Str2Null(rsOldRepor!participat)
      varCHECKNO1 = N2Str2Null(rsOldRepor!checkno1)
      varCHECKNO2 = N2Str2Null(rsOldRepor!checkno2)
      varCHECKDATE1 = N2Date2Null(rsOldRepor!checkdate1)
      varCHECKDATE2 = N2Date2Null(rsOldRepor!checkdate2)
      varCASHAMT1 = N2Str2Zero(rsOldRepor!cashamt1)
      varCREDITAMT1 = N2Str2Zero(rsOldRepor!creditamt1)
      varCASHAMT2 = N2Str2Zero(rsOldRepor!cashamt2)
      varCREDITAMT2 = N2Str2Zero(rsOldRepor!creditamt2)
      varREF_NO1 = N2Str2Null(rsOldRepor!ref_no1)
      varREF_NO2 = N2Str2Null(rsOldRepor!ref_no2)
      varCHECKAMT1 = N2Str2Zero(rsOldRepor!checkamt1)
      varCHECKAMT2 = N2Str2Zero(rsOldRepor!checkamt2)
      varCLCODE1 = N2Str2Null(rsOldRepor!clcode1)
      varCLCODE2 = N2Str2Null(rsOldRepor!clcode2)
      varCLAMT1 = N2Str2Zero(rsOldRepor!clamt1)
      varCLAMT2 = N2Str2Zero(rsOldRepor!clamt2)
      varORDATE1 = N2Date2Null(rsOldRepor!ordate1)
      varORDATE2 = N2Date2Null(rsOldRepor!ordate2)
      varINVOICE = N2Str2Null(rsOldRepor!invoice)
      varL_DISC2 = N2Str2Zero(rsOldRepor!l_disc2)
      varP_DISC2 = N2Str2Zero(rsOldRepor!p_disc2)
      varM_DISC2 = N2Str2Zero(rsOldRepor!m_disc2)
      varL_DISCOUNT = N2Str2Zero(rsOldRepor!l_discount)
      varP_DISCOUNT = N2Str2Zero(rsOldRepor!p_discount)
      varM_DISCOUNT = N2Str2Zero(rsOldRepor!m_discount)
      varL_AMTVALUE = N2Str2Zero(rsOldRepor!l_amtvalue)
      varP_AMTVALUE = N2Str2Zero(rsOldRepor!p_amtvalue)
      varM_AMTVALUE = N2Str2Zero(rsOldRepor!m_amtvalue)
      varRO_AMOUNT = N2Str2Zero(rsOldRepor!ro_amount)
      varVAT_EXEMPT = N2Str2Zero(rsOldRepor!vat_exempt)
      varDEPOSIT = N2Str2Zero(rsOldRepor!deposit)
      varDEP_ORNUM = N2Str2Null(rsOldRepor!dep_ornum)
      varDEP_DATE = N2Date2Null(rsOldRepor!dep_date)
   
      
      MoveSql = "INSERT INTO repor " & _
                "(REP_OR,ESTIMATENO,ROTYPE,SVC_NO,ACCT_NO,INSCDE,[NAME],PLATE_NO,MODEL,TERM,[SECTION],RECD_BY,KM_RDG,DTE_RECD,DTE_PRO,CERTIFIC8," & _
                "AMOUNT,INSAMT,ROVAT,LABOR,PARTS,MATERIAL,L_DISC,P_DISC,M_DISC,WL_AMT,WP_AMT,WM_AMT,L_TAXVAL,P_TAXVAL,M_TAXVAL,ENT_DATE,PRIN_DTE," & _
                "DTE_COMP,DTE_REL,INVBAL,SERVICE,DOC_CODE,USERCDE,STATUS,PREVSTAT,STATUS2,SAVEDATE,SAVETIME,VERIFIED,ORNUM1,ORNUM2,BANKCODE1,BANKCODE2,TERMS2," & _
                "PART_AMT,CHECKNO1,CHECKNO2,CHECKDATE1,CHECKDATE2,CASHAMT1,CREDITAMT1,CASHAMT2,CREDITAMT2,REF_NO1,REF_NO2,CHECKAMT1,CHECKAMT2,CLCODE1,CLCODE2,CLAMT1,CLAMT2,ORDATE1," & _
                "ORDATE2,INVOICE,L_DISC2,P_DISC2,M_DISC2,L_DISCOUNT,P_DISCOUNT,M_DISCOUNT,L_AMTVALUE,P_AMTVALUE,M_AMTVALUE,RO_AMOUNT,VAT_EXEMPT,DEPOSIT,DEP_ORNUM,DEP_DATE)" & _
                " values (" & varREP_OR & "," & varESTIMATENO & "," & varROTYPE & "," & varSVC_NO & "," & varACCT_NO & "," & varINSCDE & "," & varNAME & "," & varPLATE_NO & "," & varMODEL & "," & varTERM & "," & varSECTION & "," & varRECD_BY & "," & varKM_RDG & "," & varDTE_RECD & "," & varDTE_PRO & "," & varCERTIFIC8 & _
                "," & varAMOUNT & "," & varINSAMT & "," & varROVAT & "," & VarLabor & "," & VarParts & "," & varMATERIAL & "," & varL_DISC & "," & varP_DISC & "," & varM_DISC & "," & varWL_AMT & "," & varWP_AMT & "," & varWM_AMT & "," & varL_TAXVAL & "," & varP_TAXVAL & "," & varM_TAXVAL & "," & varENT_DATE & "," & varPRIN_DTE & _
                "," & varDTE_COMP & "," & varDTE_REL & "," & varINVBAL & "," & varSERVICE & "," & varDOC_CODE & "," & varUSERCODE & "," & varSTATUS & "," & varPREVSTAT & "," & varSTATUS2 & "," & varSAVEDATE & "," & varSAVETIME & "," & varVERIFIED & "," & varORNUM1 & "," & varORNUM2 & "," & varBANKCODE1 & "," & varBANKCODE2 & "," & varTERMS2 & _
                "," & varPART_AMT & "," & varCHECKNO1 & "," & varCHECKNO2 & "," & varCHECKDATE1 & "," & varCHECKDATE2 & "," & varCASHAMT1 & "," & varCREDITAMT1 & "," & varCASHAMT2 & "," & varCREDITAMT2 & "," & varREF_NO1 & "," & varREF_NO2 & "," & varCHECKAMT1 & "," & varCHECKAMT2 & "," & varCLCODE1 & "," & varCLCODE2 & "," & varCLAMT1 & "," & varCLAMT2 & "," & varORDATE1 & _
                "," & varORDATE2 & "," & varINVOICE & "," & varL_DISC2 & "," & varP_DISC2 & "," & varM_DISC2 & "," & varL_DISCOUNT & "," & varP_DISCOUNT & "," & varM_DISCOUNT & "," & varL_AMTVALUE & "," & varP_AMTVALUE & "," & varM_AMTVALUE & "," & varRO_AMOUNT & "," & varVAT_EXEMPT & "," & varDEPOSIT & "," & varDEP_ORNUM & "," & varDEP_DATE & ")"
      On Error GoTo ErrorCode
      gconOLDCSMIOS.Execute MoveSql
      i = i + 1
      progTransfer_Data.Value = (i / rsOldRepor.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldRepor.MoveNext
   Loop
   Me.Caption = "Repair Order Header File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveROJOBS()
Dim MoveSql As String
Dim i As Integer

Dim varJCODE, varDESC1, varDETAIL As String
Dim varSTD_MHRS, varFLAT_RATE, varFIELD_1A As Double
Dim varFIELD_2A, varFIELD_1B, varFIELD_2B As Double
Dim varFIELD_1C, varFIELD_2C, varFIELD_1D As Double
Dim varFIELD_2D, varFIELD_1E, varFIELD_2E As Double
Dim varFIELD_1F, varFIELD_2F, varFIELD_1G As Double
Dim varFIELD_2G, varFIELD_1H, varFIELD_2H As Double
Dim varFIELD_1I, varFIELD_2I, varFIELD_1J As Double
Dim varFIELD_2J, varFIELD_1K, varFIELD_2K As Double
Dim varFIELD_1L, varFIELD_2L, varFIELD_1M As Double
Dim varFIELD_2M, varFIELD_1O, varFIELD_2O As Double
Dim varFIELD_1N, varFIELD_2N, varFIELD_1P As Double
Dim varFIELD_2P, varFIELD_1Q, varFIELD_2Q As Double
Dim varFIELD_1R, varFIELD_2R As Double
Dim varPOCODE, varVALIDATE As String

Dim rsOldROjobs As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from rojobs"
Set rsOldROjobs = New ADODB.Recordset
    rsOldROjobs.Open "select * from ROjobs", gconCSMIOS
If Not rsOldROjobs.EOF And Not rsOldROjobs.BOF Then
   rsOldROjobs.MoveFirst
   Me.Caption = "Currently Converting Repair Order Jobs File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldROjobs.EOF
      varJCODE = N2Str2Null(rsOldROjobs!JCode)
      varDESC1 = N2Str2Null(rsOldROjobs!desc1)
      varDETAIL = N2Str2Null(rsOldROjobs!detail)
      varSTD_MHRS = N2Str2Zero(rsOldROjobs!std_mhrs)
      varFLAT_RATE = N2Str2Zero(rsOldROjobs!flatrate)
      varFIELD_1A = N2Str2Zero(rsOldROjobs!field_1a)
      varFIELD_2A = N2Str2Zero(rsOldROjobs!field_2a)
      varFIELD_1B = N2Str2Zero(rsOldROjobs!field_1b)
      varFIELD_2B = N2Str2Zero(rsOldROjobs!FIELD_2B)
      varFIELD_1C = N2Str2Zero(rsOldROjobs!field_1c)
      varFIELD_2C = N2Str2Zero(rsOldROjobs!field_2c)
      varFIELD_1D = N2Str2Zero(rsOldROjobs!field_1d)
      varFIELD_2D = N2Str2Zero(rsOldROjobs!field_2d)
      varFIELD_1E = N2Str2Zero(rsOldROjobs!field_1e)
      varFIELD_2E = N2Str2Zero(rsOldROjobs!FIELD_2E)
      varFIELD_1F = N2Str2Zero(rsOldROjobs!field_1f)
      varFIELD_2F = N2Str2Zero(rsOldROjobs!field_2f)
      varFIELD_1G = N2Str2Zero(rsOldROjobs!field_1g)
      varFIELD_2G = N2Str2Zero(rsOldROjobs!field_2g)
      varFIELD_1H = N2Str2Zero(rsOldROjobs!field_1h)
      varFIELD_2H = N2Str2Zero(rsOldROjobs!FIELD_2H)
      varFIELD_1I = N2Str2Zero(rsOldROjobs!field_1i)
      varFIELD_2I = N2Str2Zero(rsOldROjobs!field_2i)
      varFIELD_1J = N2Str2Zero(rsOldROjobs!field_1j)
      varFIELD_2J = N2Str2Zero(rsOldROjobs!field_2j)
      varFIELD_1K = N2Str2Zero(rsOldROjobs!field_1k)
      varFIELD_2K = N2Str2Zero(rsOldROjobs!FIELD_2K)
      varFIELD_1L = N2Str2Zero(rsOldROjobs!field_1l)
      varFIELD_2L = N2Str2Zero(rsOldROjobs!field_2l)
      varFIELD_1M = N2Str2Zero(rsOldROjobs!field_1m)
      varFIELD_2M = N2Str2Zero(rsOldROjobs!field_2m)
      varFIELD_1O = N2Str2Zero(rsOldROjobs!field_1o)
      varFIELD_2O = N2Str2Zero(rsOldROjobs!FIELD_2O)
      varFIELD_1N = N2Str2Zero(rsOldROjobs!field_1n)
      varFIELD_2N = N2Str2Zero(rsOldROjobs!field_2n)
      varFIELD_1P = N2Str2Zero(rsOldROjobs!field_1p)
      varFIELD_2P = N2Str2Zero(rsOldROjobs!field_2p)
      varFIELD_1Q = N2Str2Zero(rsOldROjobs!field_1q)
      varFIELD_2Q = N2Str2Zero(rsOldROjobs!FIELD_2Q)
      varFIELD_1R = N2Str2Zero(rsOldROjobs!field_1r)
      varFIELD_2R = N2Str2Zero(rsOldROjobs!field_2r)
      varPOCODE = N2Str2Null(rsOldROjobs!pocode)
      varVALIDATE = N2Str2Null(rsOldROjobs!Validate)
      
      MoveSql = "INSERT INTO rojobs " & _
                "(JCODE,DESC1,DETAIL,STD_MHRS,FLATRATE,FIELD_1A,FIELD_2A,FIELD_1B," & _
                "FIELD_2B,FIELD_1C,FIELD_2C,FIELD_1D,FIELD_2D,FIELD_1E," & _
                "FIELD_2E,FIELD_1F,FIELD_2F,FIELD_1G,FIELD_2G,FIELD_1H," & _
                "FIELD_2H,FIELD_1I,FIELD_2I,FIELD_1J,FIELD_2J,FIELD_1K," & _
                "FIELD_2K,FIELD_1L,FIELD_2L,FIELD_1M,FIELD_2M,FIELD_1O," & _
                "FIELD_2O,FIELD_1N,FIELD_2N,FIELD_1P,FIELD_2P,FIELD_1Q," & _
                "FIELD_2Q,FIELD_1R,FIELD_2R,POCODE,VALIDATE)" & _
                " values (" & varJCODE & "," & varDESC1 & "," & varDETAIL & "," & varSTD_MHRS & "," & varFLAT_RATE & "," & varFIELD_1A & "," & varFIELD_2A & "," & varFIELD_1B & _
                "," & varFIELD_2B & "," & varFIELD_1C & "," & varFIELD_2C & "," & varFIELD_1D & "," & varFIELD_2D & "," & varFIELD_1E & _
                "," & varFIELD_2E & "," & varFIELD_1F & "," & varFIELD_2F & "," & varFIELD_1G & "," & varFIELD_2G & "," & varFIELD_1H & _
                "," & varFIELD_2H & "," & varFIELD_1I & "," & varFIELD_2I & "," & varFIELD_1J & "," & varFIELD_2J & "," & varFIELD_1K & _
                "," & varFIELD_2K & "," & varFIELD_1L & "," & varFIELD_2L & "," & varFIELD_1M & "," & varFIELD_2M & "," & varFIELD_1O & _
                "," & varFIELD_2O & "," & varFIELD_1N & "," & varFIELD_2N & "," & varFIELD_1P & "," & varFIELD_2P & "," & varFIELD_1Q & _
                "," & varFIELD_2Q & "," & varFIELD_1R & "," & varFIELD_2R & "," & varPOCODE & "," & varVALIDATE & ")"
      On Error GoTo ErrorCode
      gconOLDCSMIOS.Execute MoveSql
      i = i + 1
      progTransfer_Data.Value = (i / rsOldROjobs.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldROjobs.MoveNext
   Loop
   Me.Caption = "Repair Order Jobs File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveMATMAS()
Dim MoveSql As String
Dim i As Integer

Dim varMATCDE, varMATDSC As String
Dim varS_PRICE, VarCOST As Double
Dim varPOCODE As String

Dim rsOldMATMAS As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from matmas"
Set rsOldMATMAS = New ADODB.Recordset
    rsOldMATMAS.Open "select * from MATMAS order by matcde asc", gconCSMIOS
If Not rsOldMATMAS.EOF And Not rsOldMATMAS.BOF Then
   rsOldMATMAS.MoveFirst
   Me.Caption = "Currently Converting Materials Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldMATMAS.EOF
      varMATCDE = N2Str2Null(rsOldMATMAS!MATCDE)
      varMATDSC = N2Str2Null(rsOldMATMAS!MatDsc)
      varS_PRICE = N2Str2Zero(rsOldMATMAS!s_price)
      VarCOST = N2Str2Zero(rsOldMATMAS!COST)
      varPOCODE = N2Str2Null(rsOldMATMAS!pocode)
      
      If varMATCDE <> "NULL" Then
         MoveSql = "INSERT INTO MATMAS " & _
                   "(MATCDE,MATDSC,S_PRICE,COST,POCODE)" & _
                   " values (" & varMATCDE & "," & varMATDSC & "," & varS_PRICE & "," & VarCOST & "," & varPOCODE & ")"
         On Error GoTo ErrorCode
         gconOLDCSMIOS.Execute MoveSql
      End If
      i = i + 1
      progTransfer_Data.Value = (i / rsOldMATMAS.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldMATMAS.MoveNext
   Loop
   Me.Caption = "Materials Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveS_Model()
Dim MoveSql As String
Dim i As Integer

Dim varMODEL, varMAKE, varYEAR, varJOBVEH As String

Dim rsOldS_Model As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from s_model"
Set rsOldS_Model = New ADODB.Recordset
    rsOldS_Model.Open "select * from S_Model order by model asc", gconCSMIOS
If Not rsOldS_Model.EOF And Not rsOldS_Model.BOF Then
   rsOldS_Model.MoveFirst
   Me.Caption = "Currently Converting Car Model Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldS_Model.EOF
      varMODEL = N2Str2Null(rsOldS_Model!model)
      varMAKE = N2Str2Null(rsOldS_Model!make)
      varYEAR = N2Str2Null(rsOldS_Model!Yeer)
      varJOBVEH = N2Str2Null(rsOldS_Model!jobveh)
      
      MoveSql = "INSERT INTO s_model " & _
                "(model,make,[year],jobveh)" & _
                " values (" & varMODEL & ", " & varMAKE & ", " & varYEAR & ", " & varJOBVEH & ")"
      On Error GoTo ErrorCode
      gconOLDCSMIOS.Execute MoveSql
      i = i + 1
      progTransfer_Data.Value = (i / rsOldS_Model.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldS_Model.MoveNext
   Loop
   Me.Caption = "Car Model Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveJobMast()
Dim MoveSql As String
Dim i As Integer

Dim varJCODE, varMAIN_CAT, varDESC1 As String
Dim varDESC2, varDETAIL As String

Dim rsOldJobmast As ADODB.Recordset
gconOLDCSMIOS.Execute "delete * from jobmast"
Set rsOldJobmast = New ADODB.Recordset
    rsOldJobmast.Open "select * from Jobmast order by Jcode asc", gconCSMIOS
If Not rsOldJobmast.EOF And Not rsOldJobmast.BOF Then
   rsOldJobmast.MoveFirst
   Me.Caption = "Currently Converting Jobs Model Master File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Do While Not rsOldJobmast.EOF
      varJCODE = N2Str2Null(rsOldJobmast!JCode)
      varMAIN_CAT = N2Str2Null(rsOldJobmast!main_cat)
      varDESC1 = N2Str2Null(Trim(Left(rsOldJobmast!Description, 35)))
      varDESC2 = N2Str2Null(Trim(Mid(rsOldJobmast!Description, 36, 35)))
      varDETAIL = N2Str2Null(rsOldJobmast!detail)
      
      If varJCODE <> "NULL" Then
         MoveSql = "INSERT INTO jobmast " & _
                   "(jcode,main_cat,desc1,desc2,detail)" & _
                   " values (" & varJCODE & ", " & varMAIN_CAT & ", " & varDESC1 & ", " & varDESC2 & ", " & varDETAIL & ")"
         On Error GoTo ErrorCode
         gconOLDCSMIOS.Execute MoveSql
      End If
      i = i + 1
      progTransfer_Data.Value = (i / rsOldJobmast.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldJobmast.MoveNext
   Loop
   Me.Caption = "Jobs Master File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

Sub MoveInvFlag()
Dim MoveSql As String
Dim i As Integer

Dim varREP_ORNUM As String
Dim varINV_NUMBER As String
Dim varSTATUS As String

Dim rsOldRepor, rsOldInvFlag As ADODB.Recordset
Set rsOldRepor = New ADODB.Recordset
    rsOldRepor.Open "select id,invoice,rep_or from repor order by id asc", gconCSMIOS
If Not rsOldRepor.EOF And Not rsOldRepor.BOF Then
   rsOldRepor.MoveFirst
   Me.Caption = "Currently Converting Invoice Flag File"
   Screen.MousePointer = 11
   DoEvents
   i = 0
   Dim NiymLen As Integer
   Do While Not rsOldRepor.EOF
      Me.Caption = "Reading RO " & Null2String(rsOldRepor!rep_or) & " invoice no. " & Null2String(rsOldRepor!invoice)
      If Null2String(rsOldRepor!invoice) <> "" Then
         varREP_ORNUM = N2Str2Null(rsOldRepor!rep_or)
         varINV_NUMBER = N2Str2Null(rsOldRepor!invoice)
         Set rsOldInvFlag = New ADODB.Recordset
             rsOldInvFlag.Open "select * from invflag where rep_ornum = " & varREP_ORNUM, gconOLDCSMIOS, adOpenForwardOnly, adLockReadOnly
         If rsOldInvFlag.EOF And rsOldInvFlag.BOF Then
            On Error GoTo ErrorCode
            gconOLDCSMIOS.Execute "insert into invflag " & _
                                  "(rep_ornum,inv_number,status)" & _
                                  " values (" & varREP_ORNUM & ", " & varINV_NUMBER & ", 'U')"
            Me.Caption = "Inserting RO " & Null2String(rsOldRepor!rep_or) & " invoice no. " & Null2String(rsOldRepor!invoice)
         End If
      End If
      i = i + 1
      progTransfer_Data.Value = (i / rsOldRepor.RecordCount) * 100
      labCPB.Caption = Int(progTransfer_Data.Value) & "% Completed"
      DoEvents
      rsOldRepor.MoveNext
   Loop
   Me.Caption = "Repair Order Header File Successfully Converted"
   Screen.MousePointer = 0
   DoEvents
End If
Exit Sub

ErrorCode:
ShowVBError
Resume Next
End Sub

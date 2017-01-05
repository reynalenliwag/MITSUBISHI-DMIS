VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmCMISUpdateCUTOFFMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Cut-Off Master File"
   ClientHeight    =   2370
   ClientLeft      =   435
   ClientTop       =   780
   ClientWidth     =   5865
   Icon            =   "UpdateCutOffMaster.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
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
      Left            =   4935
      MouseIcon       =   "UpdateCutOffMaster.frx":27A2
      MousePointer    =   99  'Custom
      Picture         =   "UpdateCutOffMaster.frx":28F4
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Exit Window"
      Top             =   1410
      Width           =   705
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Update"
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
      Left            =   4245
      MaskColor       =   &H0000FFFF&
      MouseIcon       =   "UpdateCutOffMaster.frx":2C5A
      MousePointer    =   99  'Custom
      Picture         =   "UpdateCutOffMaster.frx":2DAC
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Process New Cut-Off Entry"
      Top             =   1410
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   330
      TabIndex        =   10
      Top             =   2400
      Width           =   1305
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   60
      ScaleHeight     =   585
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   60
      Width           =   5715
      Begin VB.TextBox txtCUTDATE 
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   2760
         TabIndex        =   1
         Top             =   90
         Width           =   1755
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Cut-Off Date"
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
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2925
      End
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   60
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   3
      Top             =   720
      Width           =   5715
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   4
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
            TabIndex        =   5
            Top             =   -30
            Width           =   3525
         End
      End
      Begin wizProgBar.Prg prg1 
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "UpdateCutOffMaster.frx":30D1
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "UpdateCutOffMaster.frx":30ED
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
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   0
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   6
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   60
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
            MICON           =   "UpdateCutOffMaster.frx":3109
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
         TabIndex        =   9
         Top             =   30
         Width           =   5595
      End
   End
   Begin Crystal.CrystalReport rptCMISReportRange 
      Left            =   90
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmCMISUpdateCUTOFFMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOFF_HD                                                        As ADODB.Recordset
Dim rsOFF_DT                                                        As ADODB.Recordset
Dim rsPETTY                                                         As ADODB.Recordset
Dim rsINCASH                                                        As ADODB.Recordset
Dim rsBANKDEPO                                                      As ADODB.Recordset
Dim rsCash_Pos                                                      As ADODB.Recordset

Function SetAccountName(XXX As String)
    Dim rsChartAccounts                                             As ADODB.Recordset
    Set rsChartAccounts = New ADODB.Recordset
    Set rsChartAccounts = gconDMIS.Execute("Select * from ChartAccount where AcctCode = " & N2Str2Null(XXX))
    If Not rsChartAccounts.EOF And Not rsChartAccounts.BOF Then
        SetAccountName = Null2String(rsChartAccounts!Description)
    Else
    End If
End Function

Sub PostOREntries()
    Dim vITEMNO                                                     As String
    Dim vOR_NUM                                                     As String
    Dim vOR_DATE                                                    As String
    Dim vACCT_CODE                                                  As String
    Dim vACCT_NAME                                                  As String
    Dim vAPPLICATION                                                As String
    Dim vDEBIT                                                      As Double
    Dim vCREDIT                                                     As Double
    Dim i                                                           As Integer
    Dim ItemNoCnt                                                   As Long
    
    Set rsOFF_HD = New ADODB.Recordset
    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd Where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
        rsOFF_HD.MoveFirst: i = 0
        Do While Not rsOFF_HD.EOF
            Set rsOFF_DT = New ADODB.Recordset
            Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_Off_Dt Where OR_NUM = " & N2Str2Null(rsOFF_HD!OR_NUM) & " order by id asc")
            If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                rsOFF_DT.MoveFirst: ItemNoCnt = 0
                Do While Not rsOFF_DT.EOF
                    ItemNoCnt = ItemNoCnt + 1
                    vITEMNO = "'" & Format(ItemNoCnt, "0000") & "'"
                    vOR_NUM = N2Str2Null(rsOFF_HD!OR_NUM)
                    vOR_DATE = N2Str2Null(rsOFF_HD!OR_DATE)
                    If Null2String(rsOFF_HD!TOF) = "1" Or Null2String(rsOFF_HD!TOF) = "2" Then
                        vACCT_CODE = COA_CASH_ON_HAND
                        vDEBIT = N2Str2Zero(rsOFF_DT!Payment)
                    Else
                        vACCT_CODE = COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD
                        vDEBIT = N2Str2Zero(rsOFF_DT!Payment)
                    End If
                    vACCT_NAME = N2Str2Null(SetAccountName(vACCT_CODE))
                    vACCT_CODE = N2Str2Null(vACCT_CODE)
                    vAPPLICATION = "'" & Null2String(rsOFF_DT!TranType) & "-" & Null2String(rsOFF_DT!REFERENCE) & "'"
                    vCREDIT = 0
                    
                    gconDMIS.Execute ("Insert into CMIS_Journal_Det " & _
                                      "(ITEMNO,OR_NUM,OR_DATE,ACCT_CODE,ACCT_NAME,APPLICATION,DEBIT,CREDIT)" & _
                                      " VALUES (" & vITEMNO & "," & vOR_NUM & "," & vOR_DATE & "," & vACCT_CODE & "," & vACCT_NAME & "," & vAPPLICATION & "," & vDEBIT & "," & vCREDIT & ")")
                    
                    vDEBIT = N2Str2Zero(rsOFF_DT!Payment)
                    If Null2String(rsOFF_DT!TranType) = "RO" Then
                        vACCT_CODE = COA_AR_TRADE_SERVICE
                        vCREDIT = N2Str2Zero(rsOFF_DT!Payment)
                    ElseIf Null2String(rsOFF_DT!TranType) = "CSH" Then
                        vACCT_CODE = COA_AR_TRADE_PARTS
                        vCREDIT = N2Str2Zero(rsOFF_DT!Payment)
                    ElseIf Null2String(rsOFF_DT!TranType) = "CHG" Then
                        vACCT_CODE = COA_AR_TRADE_PARTS
                        vCREDIT = N2Str2Zero(rsOFF_DT!Payment)
                    ElseIf Null2String(rsOFF_DT!TranType) = "VI" Then
                        vACCT_CODE = COA_AR_TRADE_UNITS
                        vCREDIT = N2Str2Zero(rsOFF_DT!Payment)
                    ElseIf Null2String(rsOFF_DT!TranType) = "EST" Then
                        vACCT_CODE = COA_CUSTOMER_DEPOSIT
                        vCREDIT = N2Str2Zero(rsOFF_DT!Payment)
                    ElseIf Null2String(rsOFF_DT!TranType) = "IOC" Then
                        vACCT_CODE = COA_BRANCH_LEGASPI
                        vCREDIT = N2Str2Zero(rsOFF_DT!Payment)
                    ElseIf Null2String(rsOFF_DT!TranType) = "BRA" Then
                        vACCT_CODE = COA_BRANCH_LEGASPI
                        vCREDIT = N2Str2Zero(rsOFF_DT!Payment)
                    End If
                    'If Null2String(rsOFF_DT!TRANTYPE) = "WAR" Then
                    '   vACCT_CODE = COA_BRANCH_LEGASPI
                    '   vCREDIT = N2Str2Zero(rsOFF_DT!PAYMENT)
                    'End If
                    'If Null2String(rsOFF_DT!TRANTYPE) = "INV" Then
                    '   vCREDIT = 0
                    'End If
                    'If Null2String(rsOFF_DT!TRANTYPE) = "CRD" Then
                    '   vCREDIT = 0
                    'End If
                    If Null2String(rsOFF_DT!TranType) = "OTH" Then
                        If Null2String(rsOFF_DT!PAIDFOR) = "412" Then
                            vACCT_CODE = COA_CUSTOMER_DEPOSIT
                            vCREDIT = N2Str2Zero(rsOFF_DT!Payment)
                        ElseIf Null2String(rsOFF_DT!PAIDFOR) = "413" Then
                            vACCT_CODE = COA_INSURANCE_PREMIUM_PAYABLE
                            vCREDIT = N2Str2Zero(rsOFF_DT!Payment)
                        ElseIf Null2String(rsOFF_DT!PAIDFOR) = "414" Then
                            vACCT_CODE = COA_INSURANCE_PREMIUM_PAYABLE
                            vCREDIT = N2Str2Zero(rsOFF_DT!Payment)
                        End If
                        'If Null2String(rsOFF_DT!PAIDFOR) = "415" Then
                        '   vACCT_CODE = COA_INSURANCE_PREMIUM_PAYABLE
                        '   vCREDIT = N2Str2Zero(rsOFF_DT!PAYMENT)
                        'End If
                        'If Null2String(rsOFF_DT!PAIDFOR) = "416" Then
                        '   vCREDIT = 0
                        'End If
                        'If Null2String(rsOFF_DT!PAIDFOR) = "417" Then
                        '   vCREDIT = 0
                        'End If
                        'If Null2String(rsOFF_DT!PAIDFOR) = "418" Then
                        '   vCREDIT = 0
                        'End If
                        'If Null2String(rsOFF_DT!PAIDFOR) = "419" Then
                        '   vCREDIT = 0
                        'End If
                        'If Null2String(rsOFF_DT!PAIDFOR) = "420" Then
                        '   vCREDIT = 0
                        'End If
                        'If Null2String(rsOFF_DT!PAIDFOR) = "421" Then
                        '   vCREDIT = 0
                        'End If
                        'If Null2String(rsOFF_DT!PAIDFOR) = "422" Then
                        '   vCREDIT = 0
                        'End If
                    End If
                    ItemNoCnt = ItemNoCnt + 1
                    vITEMNO = "'" & Format(ItemNoCnt, "0000") & "'"
                    vACCT_NAME = N2Str2Null(SetAccountName(vACCT_CODE))
                    vACCT_CODE = N2Str2Null(vACCT_CODE)
                    vDEBIT = 0
                    vAPPLICATION = "'" & Null2String(rsOFF_DT!TranType) & "-" & Null2String(rsOFF_DT!REFERENCE) & "'"
                    
                    gconDMIS.Execute ("Insert into CMIS_Journal_Det " & _
                                      "(ITEMNO,OR_NUM,OR_DATE,ACCT_CODE,ACCT_NAME,APPLICATION,DEBIT,CREDIT)" & _
                                      " VALUES (" & vITEMNO & "," & vOR_NUM & "," & vOR_DATE & "," & vACCT_CODE & "," & vACCT_NAME & "," & vAPPLICATION & "," & vDEBIT & "," & vCREDIT & ")")
                    rsOFF_DT.MoveNext
                Loop
            End If
            i = i + 1
            prg1.Value = (i / rsOFF_HD.RecordCount) * 100
            prg1.Text = Int(prg1.Value) & "% Completed"
            DoEvents
            rsOFF_HD.MoveNext
        Loop
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPost_Click()
    Dim PROCESS_CUT_OFF_DATE                                        As Date
    
    If Function_Access(LOGID, "Acess_Process", "OFFICIAL RECEIPT CUT-OFF ENTRY") = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:
    
    Dim OLD_CUTOFF_DATE                                             As String
    Dim vCASH                                                       As Double
    Dim vCHECK                                                      As Double
    Dim vCARD                                                       As Double
    Dim vREPLENISH                                                  As Double
    Dim vEXPENSE                                                    As Double
    Dim vADVANCES                                                   As Double

    Dim vCASHDEPO                                                   As Double
    Dim vCHECKDEPO                                                  As Double
    Dim vCARDDEPO                                                   As Double
    Dim vFUND                                                       As Double
    Dim vLTO_REPL                                                   As Double
    Dim vLTO_EXP                                                    As Double
    Dim vLTO_ADV                                                    As Double
    Dim vLTO_AR                                                     As Double
    Dim vLTO                                                        As Double
    Dim vPETPAYMENT                                                 As Double
    Dim vLTOPAYMENT                                                 As Double
    Dim vPETTYCASH                                                  As Double
    Dim vLTOCASH                                                    As Double
    Dim vINCASHMENT                                                 As Double
    Dim vLTOINCASH                                                  As Double
    Dim vCHKPAYMENT                                                 As Double
    Dim vCHKINCASH                                                  As Double
    Dim vCSHPAYMENT                                                 As Double
    Dim vCSHINCASH                                                  As Double
    Dim vBEGIN                                                      As Double
    Dim vEND                                                        As Double

    Dim HARI                                                        As Integer
    Dim rsCUTOFF                                                    As ADODB.Recordset
    Dim rsPREV_CUTOFF                                               As ADODB.Recordset

    OLD_CUTOFF_DATE = DateSerial(Year(txtCUTDATE), Month(txtCUTDATE), Day(txtCUTDATE))
    HARI = DateSerial(Year(LOGDATE), Month(LOGDATE), Day(LOGDATE)) - DateSerial(Year(txtCUTDATE), Month(txtCUTDATE), Day(txtCUTDATE))

    Set rsPREV_CUTOFF = New ADODB.Recordset
    Set rsPREV_CUTOFF = gconDMIS.Execute("Select * from CMIS_Cash_Pos where CUTDATE = '" & OLD_CUTOFF_DATE & "'")
    If Not rsPREV_CUTOFF.EOF And Not rsPREV_CUTOFF.BOF Then
        vCASH = N2Str2Zero(rsPREV_CUTOFF!CASH)
        vCHECK = 0    'N2Str2Zero(rsPREV_CUTOFF!CHECK)
        vCARD = 0    'N2Str2Zero(rsPREV_CUTOFF!CARD)
        'vREPLENISH = N2Str2Zero(rsPREV_CUTOFF!REPLENISH)
        'vEXPENSE = N2Str2Zero(rsPREV_CUTOFF!EXPENSE)
        'vADVANCES = N2Str2Zero(rsPREV_CUTOFF!ADVANCES)
        vREPLENISH = 0
        vEXPENSE = 0
        vADVANCES = 0
        vCASHDEPO = 0
        vCHECKDEPO = 0
        vCARDDEPO = 0
        vFUND = N2Str2Zero(rsPREV_CUTOFF!FUND)
        vLTO_REPL = 0    'N2Str2Zero(rsPREV_CUTOFF!LTO_REPL)
        vLTO_EXP = 0    'N2Str2Zero(rsPREV_CUTOFF!LTO_EXP)
        vLTO_ADV = 0    'N2Str2Zero(rsPREV_CUTOFF!LTO_ADV)
        vLTO_AR = N2Str2Zero(rsPREV_CUTOFF!LTO_AR)
        vLTO = N2Str2Zero(rsPREV_CUTOFF!LTO)
        vPETPAYMENT = N2Str2Zero(rsPREV_CUTOFF!PETPAYMENT)
        vLTOPAYMENT = N2Str2Zero(rsPREV_CUTOFF!LTOPAYMENT)
        vPETTYCASH = N2Str2Zero(rsPREV_CUTOFF!PETTYCASH)
        vLTOCASH = N2Str2Zero(rsPREV_CUTOFF!LTOCASH)
        vINCASHMENT = N2Str2Zero(rsPREV_CUTOFF!INCASHMENT)
        vLTOINCASH = N2Str2Zero(rsPREV_CUTOFF!LTOINCASH)
        vCHKPAYMENT = N2Str2Zero(rsPREV_CUTOFF!CHKPAYMENT)
        vCHKINCASH = N2Str2Zero(rsPREV_CUTOFF!CHKINCASH)
        vCSHPAYMENT = N2Str2Zero(rsPREV_CUTOFF!CSHPAYMENT)
        vCSHINCASH = N2Str2Zero(rsPREV_CUTOFF!CSHINCASH)
        vBEGIN = vCASH + vCHECK + vCARD
        vEND = 0
        
        gconDMIS.Execute ("Update CMIS_Cash_Pos Set TAG = 1, [END] = " & vBEGIN & " where CUTDATE = '" & OLD_CUTOFF_DATE & "'")
    Else
        vCASH = 0
        vCHECK = 0
        vCARD = 0
        vREPLENISH = 0
        vEXPENSE = 0
        vADVANCES = 0
        vCASHDEPO = 0
        vCHECKDEPO = 0
        vCARDDEPO = 0
        vFUND = 0
        vLTO_REPL = 0
        vLTO_EXP = 0
        vLTO_ADV = 0
        vLTO_AR = 0
        vLTO = 0
        vPETPAYMENT = 0
        vLTOPAYMENT = 0
        vPETTYCASH = 0
        vLTOCASH = 0
        vINCASHMENT = 0
        vLTOINCASH = 0
        vCHKPAYMENT = 0
        vCHKINCASH = 0
        vCSHPAYMENT = 0
        vCSHINCASH = 0
        vBEGIN = 0
        vEND = 0
        
        gconDMIS.Execute ("Insert into CMIS_Cash_Pos " & _
                          "(CUTDATE,CASH,[CHECK],CARD,REPLENISH,EXPENSE,ADVANCES,CASHDEPO,CHECKDEPO,CARDDEPO,FUND,LTO_REPL,LTO_EXP,LTO_ADV,LTO_AR,LTO,PETPAYMENT,LTOPAYMENT,PETTYCASH,LTOCASH,INCASHMENT,LTOINCASH,CHKPAYMENT,CHKINCASH,CSHPAYMENT,CSHINCASH,[BEGIN],[END])" & _
                          " VALUES ('" & OLD_CUTOFF_DATE & "'," & vCASH & "," & vCHECK & "," & vCARD & "," & vREPLENISH & "," & vEXPENSE & "," & vADVANCES & "," & vCASHDEPO & "," & vCHECKDEPO & "," & vCARDDEPO & "," & vFUND & "," & vLTO_REPL & "," & vLTO_EXP & "," & vLTO_ADV & "," & vLTO_AR & "," & vLTO & "," & vPETPAYMENT & "," & vLTOPAYMENT & "," & vPETTYCASH & "," & vLTOCASH & "," & vINCASHMENT & "," & vLTOINCASH & "," & vCHKPAYMENT & "," & vCHKINCASH & "," & vCSHPAYMENT & "," & vCSHINCASH & "," & vBEGIN & "," & vEND & ")")
    End If
    
    Dim KIM                                                         As Integer
    Dim rsOFF_HD                                                    As ADODB.Recordset
    Dim rsINCASH                                                    As ADODB.Recordset
    Dim rsBANKDEPO                                                  As ADODB.Recordset
    Dim rsPETTY                                                     As ADODB.Recordset
    
    Screen.MousePointer = 11
    For KIM = 0 To HARI
        OLD_CUTOFF_DATE = DateSerial(Year(txtCUTDATE), Month(txtCUTDATE), Day(txtCUTDATE)) + KIM - 1
        PROCESS_CUT_OFF_DATE = DateSerial(Year(txtCUTDATE), Month(txtCUTDATE), Day(txtCUTDATE)) + KIM
        
        Set rsCUTOFF = New ADODB.Recordset
        Set rsCUTOFF = gconDMIS.Execute("Select * from CMIS_Cash_Pos where CUTDATE = '" & PROCESS_CUT_OFF_DATE & "'")
        If Not rsCUTOFF.EOF And Not rsCUTOFF.BOF Then
            vBEGIN = rsCUTOFF![Begin]
        End If
        
        Set rsOFF_HD = New ADODB.Recordset
        Set rsOFF_HD = gconDMIS.Execute("Select SUM(CASHAMOUNT) AS CASH,SUM(CHKAMOUNT) AS [CHECK],SUM(CARDAMOUNT) AS CARD from CMIS_OFF_HD WHERE CANCEL = 0 AND OR_DATE = '" & PROCESS_CUT_OFF_DATE & "' AND OR_DATE <> '" & OLD_CUTOFF_DATE & "'")
        If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
            vCASH = vCASH + N2Str2Zero(rsOFF_HD!CASH)
            vCHECK = vCHECK + N2Str2Zero(rsOFF_HD![CHECK])
            vCARD = vCARD + N2Str2Zero(rsOFF_HD!CARD)
            gconDMIS.Execute ("UPDATE CMIS_OFF_HD Set CUTDATE = '" & PROCESS_CUT_OFF_DATE & "' WHERE CANCEL = 0 AND OR_DATE = '" & PROCESS_CUT_OFF_DATE & "' AND OR_DATE <> '" & OLD_CUTOFF_DATE & "'")
        End If
        
        'If PROCESS_CUT_OFF_DATE = CDate("8/6/2007") Then Stop
        Set rsINCASH = New ADODB.Recordset
        Set rsINCASH = gconDMIS.Execute("Select SUM(CHKAMOUNT) AS [CHECK] from CMIS_INCASH WHERE INCASHDATE = '" & PROCESS_CUT_OFF_DATE & "'")
        If Not rsINCASH.EOF And Not rsINCASH.BOF Then
            'If PROCESS_CUT_OFF_DATE <> CDate(txtCUTDATE) Then vCASH = vCASH - N2Str2Zero(rsINCASH![CHECK])
            vCHECK = vCHECK + N2Str2Zero(rsINCASH![CHECK])
            vINCASHMENT = N2Str2Zero(rsINCASH![CHECK])
            gconDMIS.Execute ("UPDATE CMIS_INCASH Set CUTDATE = '" & PROCESS_CUT_OFF_DATE & "' WHERE INCASHDATE = '" & PROCESS_CUT_OFF_DATE & "'")
        End If
        
        Set rsBANKDEPO = New ADODB.Recordset
        Set rsBANKDEPO = gconDMIS.Execute("Select SUM(DEPOSIT) AS CASH from CMIS_BANKDEPO WHERE TYPE = '1' AND DATDEPOSIT = '" & PROCESS_CUT_OFF_DATE & "'")
        If Not rsBANKDEPO.EOF And Not rsBANKDEPO.BOF Then
            vCASH = vCASH - N2Str2Zero(rsBANKDEPO!CASH)
            vCASHDEPO = N2Str2Zero(rsBANKDEPO!CASH)
        End If
        
        Set rsBANKDEPO = New ADODB.Recordset
        Set rsBANKDEPO = gconDMIS.Execute("Select SUM(DEPOSIT) AS [CHECK] from CMIS_BANKDEPO WHERE TYPE = '2' AND DATDEPOSIT = '" & PROCESS_CUT_OFF_DATE & "'")
        If Not rsBANKDEPO.EOF And Not rsBANKDEPO.BOF Then
            vCHECK = vCHECK - N2Str2Zero(rsBANKDEPO![CHECK])
            vCHECKDEPO = N2Str2Zero(rsBANKDEPO![CHECK])
        End If
        
        Set rsBANKDEPO = New ADODB.Recordset
        Set rsBANKDEPO = gconDMIS.Execute("Select SUM(DEPOSIT) AS CARD from CMIS_BANKDEPO WHERE TYPE = '3' AND DATDEPOSIT = '" & PROCESS_CUT_OFF_DATE & "'")
        If Not rsBANKDEPO.EOF And Not rsBANKDEPO.BOF Then
            vCARD = vCARD - N2Str2Zero(rsBANKDEPO!CARD)
            vCARDDEPO = N2Str2Zero(rsBANKDEPO!CARD)
        End If
        gconDMIS.Execute ("UPDATE CMIS_BANKDEPO Set CUTDATE = '" & PROCESS_CUT_OFF_DATE & "' WHERE DATDEPOSIT = '" & PROCESS_CUT_OFF_DATE & "'")
        
        Set rsPETTY = New ADODB.Recordset
        Set rsPETTY = gconDMIS.Execute("Select SUM(PETTY_CASH) AS PETTY_CASH from CMIS_Petty where (PETTY_DATE <= '" & PROCESS_CUT_OFF_DATE & "' AND PETTY_CODE = '002') AND PETTY_CASH > 0")
        If Not rsPETTY.EOF And Not rsPETTY.BOF Then
            vADVANCES = N2Str2Zero(rsPETTY!PETTY_CASH)
        End If
        gconDMIS.Execute ("UPDATE CMIS_BANKDEPO Set CUTDATE = '" & PROCESS_CUT_OFF_DATE & "' WHERE DATDEPOSIT = '" & PROCESS_CUT_OFF_DATE & "'")
        
        Set rsPETTY = New ADODB.Recordset
        Set rsPETTY = gconDMIS.Execute("Select SUM(PETTY_CASH) AS PETTY_CASH from CMIS_Petty where (PETTY_DATE = '" & PROCESS_CUT_OFF_DATE & "' AND PETTY_CODE = '001' AND REPLENISH <> '1') OR (PETTY_DATE < '" & PROCESS_CUT_OFF_DATE & "' AND PETTY_CODE = '001' AND REPLENISH <> '1')")
        If Not rsPETTY.EOF And Not rsPETTY.BOF Then
            vEXPENSE = N2Str2Zero(rsPETTY!PETTY_CASH)
        End If
        
        Set rsPETTY = New ADODB.Recordset
        Set rsPETTY = gconDMIS.Execute("Select SUM(PETTY_CASH) AS PETTY_CASH from CMIS_Petty where PETTY_CODE = '001' AND REPLENISH = '1' AND PETTY_DATE = '" & PROCESS_CUT_OFF_DATE & "'")
        If Not rsPETTY.EOF And Not rsPETTY.BOF Then
            vREPLENISH = vREPLENISH + N2Str2Zero(rsPETTY!PETTY_CASH)
        End If

        Set rsPETTY = New ADODB.Recordset
        Set rsPETTY = gconDMIS.Execute("Select SUM(CHKAMOUNT) AS CHKAMOUNT from CMIS_PettyPay where CUTDATE = '" & PROCESS_CUT_OFF_DATE & "'")
        If Not rsPETTY.EOF And Not rsPETTY.BOF Then
            vREPLENISH = vREPLENISH - N2Str2Zero(rsPETTY!CHKAMOUNT)
        End If

        Set rsPETTY = New ADODB.Recordset
        Set rsPETTY = gconDMIS.Execute("Select SUM(PETTY_CASH) AS PETTY_CASH from CMIS_LTOPondo where (PETTY_DATE <= '" & PROCESS_CUT_OFF_DATE & "' AND PETTY_CODE = '002') AND PETTY_CASH > 0")
        If Not rsPETTY.EOF And Not rsPETTY.BOF Then
            vLTO_ADV = N2Str2Zero(rsPETTY!PETTY_CASH)
        End If
        
        Set rsPETTY = New ADODB.Recordset
        Set rsPETTY = gconDMIS.Execute("Select SUM(PETTY_CASH) AS PETTY_CASH from CMIS_LTOPondo where (PETTY_DATE = '" & PROCESS_CUT_OFF_DATE & "' AND PETTY_CODE = '001' AND REPLENISH <> '1') OR (PETTY_DATE < '" & PROCESS_CUT_OFF_DATE & "' AND PETTY_CODE = '001' AND REPLENISH <> '1')")
        If Not rsPETTY.EOF And Not rsPETTY.BOF Then
            vLTO_EXP = N2Str2Zero(rsPETTY!PETTY_CASH)
        End If
        
        Set rsPETTY = New ADODB.Recordset
        Set rsPETTY = gconDMIS.Execute("Select SUM(PETTY_CASH) AS PETTY_CASH from CMIS_LTOPondo where PETTY_CODE = '001' AND REPLENISH = '1' AND PETTY_DATE = '" & PROCESS_CUT_OFF_DATE & "'")
        If Not rsPETTY.EOF And Not rsPETTY.BOF Then
            vLTO_REPL = vLTO_REPL + N2Str2Zero(rsPETTY!PETTY_CASH)
        End If
        
        Set rsPETTY = Nothing
        Set rsPETTY = New ADODB.Recordset
        Set rsPETTY = gconDMIS.Execute("Select SUM(CHKAMOUNT) AS CHKAMOUNT from CMIS_LTOPay where CUTDATE = '" & PROCESS_CUT_OFF_DATE & "'")
        If Not rsPETTY.EOF And Not rsPETTY.BOF Then
            vLTO_REPL = vLTO_REPL - N2Str2Zero(rsPETTY!CHKAMOUNT)
        End If
        vEND = vCASH + vCHECK + vCARD

        'gconDMIS.Execute ("update CMIS_Off_Hd Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_Off_Dt Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_LTOPondo Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_Petty Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_PettyPay Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_InCash Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_BankDepo Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_TranList Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")

        'gconDMIS.Execute ("update CMIS_LTOPondo Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_Petty Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_PettyPay Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_InCash Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_BankDepo Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_TranList Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")

        Set rsCUTOFF = New ADODB.Recordset
        Set rsCUTOFF = gconDMIS.Execute("Select * from CMIS_Cash_Pos where CUTDATE = '" & PROCESS_CUT_OFF_DATE & "'")
        If Not rsCUTOFF.EOF And Not rsCUTOFF.BOF Then
            gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                              " CASH = " & vCASH & ", [CHECK] = " & vCHECK & ", CARD = " & vCARD & "," & _
                              " REPLENISH = " & vREPLENISH & ", EXPENSE = " & vEXPENSE & ", ADVANCES = " & vADVANCES & "," & _
                              " CASHDEPO = " & vCASHDEPO & ", CHECKDEPO = " & vCHECKDEPO & ", CARDDEPO = " & vCARDDEPO & "," & _
                              " FUND = " & vFUND & "," & _
                              " LTO_REPL = " & vLTO_REPL & ", LTO_EXP = " & vLTO_EXP & ", LTO_ADV = " & vLTO_ADV & ", LTO_AR = " & vLTO_AR & "," & _
                              " LTO = " & vLTO & "," & _
                              " PETPAYMENT = " & vPETPAYMENT & ", LTOPAYMENT = " & vLTOPAYMENT & "," & _
                              " PETTYCASH = " & vPETTYCASH & "," & _
                              " INCASHMENT = " & vINCASHMENT & "," & _
                              " LTOINCASH = " & vLTOINCASH & "," & _
                              " CHKPAYMENT = " & vCHKPAYMENT & "," & _
                              " CSHPAYMENT = " & vCSHPAYMENT & "," & _
                              " CSHINCASH = " & vCSHINCASH & "," & _
                              " [BEGIN] = " & vBEGIN & "," & _
                              " [END] = " & vEND & "," & _
                              " TAG = 1 WHERE CUTDATE = '" & PROCESS_CUT_OFF_DATE & "'")
        Else
            gconDMIS.Execute ("Insert into CMIS_Cash_Pos " & _
                              "(CUTDATE,CASH,[CHECK],CARD,REPLENISH,EXPENSE,ADVANCES,CASHDEPO,CHECKDEPO,CARDDEPO,FUND,LTO_REPL,LTO_EXP,LTO_ADV,LTO_AR,LTO,PETPAYMENT,LTOPAYMENT,PETTYCASH,LTOCASH,INCASHMENT,LTOINCASH,CHKPAYMENT,CHKINCASH,CSHPAYMENT,CSHINCASH,[BEGIN],[END])" & _
                              " VALUES ('" & PROCESS_CUT_OFF_DATE & "'," & vCASH & "," & vCHECK & "," & vCARD & "," & vREPLENISH & "," & vEXPENSE & "," & vADVANCES & "," & vCASHDEPO & "," & vCHECKDEPO & "," & vCARDDEPO & "," & vFUND & "," & vLTO_REPL & "," & vLTO_EXP & "," & vLTO_ADV & "," & vLTO_AR & "," & vLTO & "," & vPETPAYMENT & "," & vLTOPAYMENT & "," & vPETTYCASH & "," & vLTOCASH & "," & vINCASHMENT & "," & vLTOINCASH & "," & vCHKPAYMENT & "," & vCHKINCASH & "," & vCSHPAYMENT & "," & vCSHINCASH & "," & vBEGIN & "," & vEND & ")")
            
            gconDMIS.Execute ("Insert into CMIS_Cash " & _
                              "(CUTDATE)" & _
                              " VALUES ('" & PROCESS_CUT_OFF_DATE & "')")
        End If
        prg1.Value = (KIM + 1 / HARI) * 100
    Next

    Set rsBANKDEPO = New ADODB.Recordset
    Set rsBANKDEPO = gconDMIS.Execute("Select * from CMIS_BankDepo order by datdeposit asc")
    If Not rsBANKDEPO.EOF And Not rsBANKDEPO.BOF Then
        rsBANKDEPO.MoveFirst
        Do While Not rsBANKDEPO.EOF
            gconDMIS.Execute ("Update CMIS_OFF_HD set DEPOSIT = 1 where OR_NUM = " & N2Str2Null(rsBANKDEPO!OR_NUM))
            rsBANKDEPO.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
    LogAudit "R", "UPDATE CUT-OFF MASTER FILE", txtCUTDATE
    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    PostOREntries
End Sub

Private Sub Form_Load()
     CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    Dim rsCash_Pos                                                  As ADODB.Recordset
    Set rsCash_Pos = New ADODB.Recordset
    Set rsCash_Pos = gconDMIS.Execute("Select * from CMIS_Cash_Pos order by CUTDATE ASC")
    If Not rsCash_Pos.EOF And Not rsCash_Pos.BOF Then
        txtCUTDATE.Text = Null2Date(rsCash_Pos!CUTDATE)
    Else
        txtCUTDATE.Text = firstDay(LOGDATE) - 1   'CURRENT_CUTOFF_DATE
    End If
End Sub


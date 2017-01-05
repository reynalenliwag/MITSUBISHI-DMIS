VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmCMISProcessCUTOFF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process New Cut-Off Entry"
   ClientHeight    =   2160
   ClientLeft      =   435
   ClientTop       =   780
   ClientWidth     =   5865
   Icon            =   "ProcessCutOffEntry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2160
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
      Left            =   4965
      MouseIcon       =   "ProcessCutOffEntry.frx":27A2
      MousePointer    =   99  'Custom
      Picture         =   "ProcessCutOffEntry.frx":28F4
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Exit Window"
      Top             =   1260
      Width           =   765
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
      Left            =   4215
      MaskColor       =   &H0000FFFF&
      MouseIcon       =   "ProcessCutOffEntry.frx":2C5A
      MousePointer    =   99  'Custom
      Picture         =   "ProcessCutOffEntry.frx":2DAC
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Process New Cut-Off Entry"
      Top             =   1260
      Width           =   765
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
      Top             =   30
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
         Left            =   2190
         TabIndex        =   1
         Top             =   90
         Width           =   1845
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "New Cut-Off Date"
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
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   2175
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
      Top             =   570
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
            Top             =   60
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
         Picture         =   "ProcessCutOffEntry.frx":30D1
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "ProcessCutOffEntry.frx":30ED
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
         ScaleWidth      =   4545
         TabIndex        =   6
         Top             =   660
         Width           =   4545
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
            MICON           =   "ProcessCutOffEntry.frx":3109
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
         Top             =   60
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
Attribute VB_Name = "frmCMISProcessCUTOFF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOFF_HD                                                          As ADODB.Recordset
Dim rsOFF_DT                                                          As ADODB.Recordset
Dim rsCash_Pos                                                        As ADODB.Recordset
Dim REPORT_CUTDATE                                                    As String

Function SetAccountName(XXX As String)
    Dim rsChartAccounts                                               As ADODB.Recordset
    Set rsChartAccounts = New ADODB.Recordset
    Set rsChartAccounts = gconDMIS.Execute("Select * from ChartAccount Where AcctCode = " & N2Str2Null(XXX))
    If Not rsChartAccounts.EOF And Not rsChartAccounts.BOF Then
        SetAccountName = Null2String(rsChartAccounts!Description)
    Else
    End If
End Function

Sub PostOREntries()
    Dim vITEMNO                                                       As String
    Dim vOR_NUM                                                       As String
    Dim vOR_DATE                                                      As String
    Dim vACCT_CODE                                                    As String
    Dim vACCT_NAME                                                    As String
    Dim vAPPLICATION                                                  As String
    Dim vDEBIT                                                        As Double
    Dim vCREDIT                                                       As Double
    Dim I                                                             As Integer
    Dim ItemNoCnt                                                     As Long
    Set rsOFF_HD = New ADODB.Recordset
    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd Where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
        rsOFF_HD.MoveFirst: I = 0
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
                        vDEBIT = N2Str2Zero(rsOFF_DT!payment)
                    Else
                        vACCT_CODE = COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD
                        vDEBIT = N2Str2Zero(rsOFF_DT!payment)
                    End If
                    vACCT_NAME = N2Str2Null(SetAccountName(vACCT_CODE))
                    vACCT_CODE = N2Str2Null(vACCT_CODE)
                    vAPPLICATION = "'" & Null2String(rsOFF_DT!TRANTYPE) & "-" & Null2String(rsOFF_DT!REFERENCE) & "'"
                    vCREDIT = 0
                    gconDMIS.Execute ("Insert into CMIS_Journal_Det " & _
                                      "(ITEMNO,OR_NUM,OR_DATE,ACCT_CODE,ACCT_NAME,APPLICATION,DEBIT,CREDIT)" & _
                                    " values (" & vITEMNO & "," & vOR_NUM & "," & vOR_DATE & "," & vACCT_CODE & "," & vACCT_NAME & "," & vAPPLICATION & "," & vDEBIT & "," & vCREDIT & ")")
                    vDEBIT = N2Str2Zero(rsOFF_DT!payment)
                    If Null2String(rsOFF_DT!TRANTYPE) = "RO" Then
                        vACCT_CODE = COA_AR_TRADE_SERVICE
                        vCREDIT = N2Str2Zero(rsOFF_DT!payment)
                    End If
                    If Null2String(rsOFF_DT!TRANTYPE) = "CSH" Then
                        vACCT_CODE = COA_AR_TRADE_PARTS
                        vCREDIT = N2Str2Zero(rsOFF_DT!payment)
                    End If
                    If Null2String(rsOFF_DT!TRANTYPE) = "CHG" Then
                        vACCT_CODE = COA_AR_TRADE_PARTS
                        vCREDIT = N2Str2Zero(rsOFF_DT!payment)
                    End If
                    If Null2String(rsOFF_DT!TRANTYPE) = "VI" Then
                        vACCT_CODE = COA_AR_TRADE_UNITS
                        vCREDIT = N2Str2Zero(rsOFF_DT!payment)
                    End If
                    If Null2String(rsOFF_DT!TRANTYPE) = "EST" Then
                        vACCT_CODE = COA_CUSTOMER_DEPOSIT
                        vCREDIT = N2Str2Zero(rsOFF_DT!payment)
                    End If
                    If Null2String(rsOFF_DT!TRANTYPE) = "IOC" Then
                        vACCT_CODE = COA_BRANCH_LEGASPI
                        vCREDIT = N2Str2Zero(rsOFF_DT!payment)
                    End If
                    If Null2String(rsOFF_DT!TRANTYPE) = "BRA" Then
                        vACCT_CODE = COA_BRANCH_LEGASPI
                        vCREDIT = N2Str2Zero(rsOFF_DT!payment)
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
                    If Null2String(rsOFF_DT!TRANTYPE) = "OTH" Then
                        If Null2String(rsOFF_DT!PAIDFOR) = "412" Then
                            vACCT_CODE = COA_CUSTOMER_DEPOSIT
                            vCREDIT = N2Str2Zero(rsOFF_DT!payment)
                        End If
                        If Null2String(rsOFF_DT!PAIDFOR) = "413" Then
                            vACCT_CODE = COA_INSURANCE_PREMIUM_PAYABLE
                            vCREDIT = N2Str2Zero(rsOFF_DT!payment)
                        End If
                        If Null2String(rsOFF_DT!PAIDFOR) = "414" Then
                            vACCT_CODE = COA_INSURANCE_PREMIUM_PAYABLE
                            vCREDIT = N2Str2Zero(rsOFF_DT!payment)
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
                    vAPPLICATION = "'" & Null2String(rsOFF_DT!TRANTYPE) & "-" & Null2String(rsOFF_DT!REFERENCE) & "'"
                    gconDMIS.Execute ("Insert into CMIS_Journal_Det " & _
                                      "(ITEMNO,OR_NUM,OR_DATE,ACCT_CODE,ACCT_NAME,APPLICATION,DEBIT,CREDIT)" & _
                                    " values (" & vITEMNO & "," & vOR_NUM & "," & vOR_DATE & "," & vACCT_CODE & "," & vACCT_NAME & "," & vAPPLICATION & "," & vDEBIT & "," & vCREDIT & ")")
                    rsOFF_DT.MoveNext
                Loop
            End If
            I = I + 1
            prg1.Value = (I / rsOFF_HD.RecordCount) * 100
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
    If Function_Access(LOGID, "Acess_Process", "OFFICIAL RECEIPT CUT-OFF ENTRY") = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:
    Dim rsCashCount                                                   As ADODB.Recordset
    Dim TotalCashCounted                                              As Double
    Set rsCashCount = New ADODB.Recordset
    Set rsCashCount = gconDMIS.Execute("Select * from CMIS_Cash Where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    If Not rsCashCount.EOF And Not rsCashCount.BOF Then
        TotalCashCounted = (NumericVal(rsCashCount!ISANGLIBO) * 1000) + (NumericVal(rsCashCount!LIMANGDAAN) * 500) + (NumericVal(rsCashCount!DALAWANGDAAN) * 200) + (NumericVal(rsCashCount!ISANGDAAN) * 100) + (NumericVal(rsCashCount!SINGKWENTA) * 50) + (NumericVal(rsCashCount!BENTE) * 20) + (NumericVal(rsCashCount!SAMPU) * 10) + (NumericVal(rsCashCount!LIMANGPISO) * 5) + (NumericVal(rsCashCount!BENTESINKO) * 0.25) + (NumericVal(rsCashCount!DYES) * 0.1) + (NumericVal(rsCashCount!SINKO) * 0.05) + (NumericVal(rsCashCount!SENTIMO) * 0.01)
    End If


    Dim OLD_CUTOFF_DATE                                               As String

    If MsgBox("Proceed CUT OFF process?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        REPORT_CUTDATE = CDate(CURRENT_CUTOFF_DATE)
        OLD_CUTOFF_DATE = CURRENT_CUTOFF_DATE
        If CDate(CURRENT_CUTOFF_DATE) = CDate(txtCutDate.Text) Then
            MsgBox "Invalid New Cut Off Date!", vbExclamation, "Error"
            Exit Sub
        End If
        
        If COMPANY_CODE = "HGC" Then
        
        Else
            If CheckUnPostedOR = True Then
                MsgBox "Cut Off processing halted. Please check Unposted OR.", vbExclamation, "Unposted OR"
                PrintUnpostedOR
                Exit Sub
            End If
        End If

        Set rsCash_Pos = New ADODB.Recordset
        Set rsCash_Pos = gconDMIS.Execute("Select * from CMIS_Cash_Pos Where CUTDATE = '" & txtCutDate.Text & "'")
        If Not rsCash_Pos.EOF And Not rsCash_Pos.BOF Then
            MsgBox "Invalid New Cut Off Date!", vbExclamation, "Error"
            Exit Sub
        End If


        'gconDMIS.Execute ("update CMIS_Off_Hd Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_Off_Dt Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_LTOPondo Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_Petty Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_PettyPay Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_InCash Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_BankDepo Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
        'gconDMIS.Execute ("update CMIS_TranList Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")

        'REMARKED BY AXP DUE TO SECURITY REASON
        Set rsCash_Pos = New ADODB.Recordset
        Set rsCash_Pos = gconDMIS.Execute("Select * from CMIS_Cash_Pos Where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        '
        If Not rsCash_Pos.EOF And Not rsCash_Pos.BOF Then
            '           If NumericVal(rsCash_Pos!CASH) <> TotalCashCounted Then
            '              MsgBox "Ending Balance for Cash On Hand is not Equal" & vbCrLf & "in Cash Counted by Denomination." & vbCrLf & "Pls Check your Cash Count...", vbInformation, "Can not Proceed Posting"
            '              Exit Sub
            '           End If
            '

            gconDMIS.Execute ("update CMIS_Off_Hd Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
            gconDMIS.Execute ("update CMIS_Off_Dt Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
            gconDMIS.Execute ("update CMIS_LTOPondo Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
            gconDMIS.Execute ("update CMIS_Petty Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
            gconDMIS.Execute ("update CMIS_PettyPay Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
            gconDMIS.Execute ("update CMIS_InCash Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
            gconDMIS.Execute ("update CMIS_BankDepo Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")
            gconDMIS.Execute ("update CMIS_TranList Set CUTDATE = '" & CURRENT_CUTOFF_DATE & "' Where CUTDATE IS NULL")

            Dim vCASH                                                 As Double
            Dim vCHECK                                                As Double
            Dim vCARD                                                 As Double
            Dim vREPLENISH                                            As Double
            Dim vEXPENSE                                              As Double
            Dim vADVANCES                                             As Double

            Dim vCASHDEPO                                             As Double
            Dim vCHECKDEPO                                            As Double
            Dim vCARDDEPO                                             As Double
            Dim vFUND                                                 As Double
            Dim vLTO_REPL                                             As Double
            Dim vLTO_EXP                                              As Double
            Dim vLTO_ADV                                              As Double
            Dim vLTO_AR                                               As Double
            Dim vLTO                                                  As Double
            Dim vPETPAYMENT                                           As Double
            Dim vLTOPAYMENT                                           As Double
            Dim vPETTYCASH                                            As Double
            Dim vLTOCASH                                              As Double
            Dim vINCASHMENT                                           As Double
            Dim vLTOINCASH                                            As Double
            Dim vCHKPAYMENT                                           As Double
            Dim vCHKINCASH                                            As Double
            Dim vCSHPAYMENT                                           As Double
            Dim vCSHINCASH                                            As Double
            Dim vBEGIN                                                As Double
            Dim vEND                                                  As Double

            vCASH = N2Str2Zero(rsCash_Pos!CASH)
            vCHECK = N2Str2Zero(rsCash_Pos!CHECK)
            vCARD = N2Str2Zero(rsCash_Pos!CARD)
            vREPLENISH = N2Str2Zero(rsCash_Pos!REPLENISH)
            vEXPENSE = N2Str2Zero(rsCash_Pos!EXPENSE)
            vADVANCES = N2Str2Zero(rsCash_Pos!ADVANCES)
            vCASHDEPO = 0
            vCHECKDEPO = 0
            vCARDDEPO = 0
            vFUND = N2Str2Zero(rsCash_Pos!FUND)
            vLTO_REPL = N2Str2Zero(rsCash_Pos!LTO_REPL)
            vLTO_EXP = N2Str2Zero(rsCash_Pos!LTO_EXP)
            vLTO_ADV = N2Str2Zero(rsCash_Pos!LTO_ADV)
            vLTO_AR = N2Str2Zero(rsCash_Pos!LTO_AR)
            vLTO = N2Str2Zero(rsCash_Pos!LTO)
            vPETPAYMENT = N2Str2Zero(rsCash_Pos!PETPAYMENT)
            vLTOPAYMENT = N2Str2Zero(rsCash_Pos!LTOPAYMENT)
            vPETTYCASH = N2Str2Zero(rsCash_Pos!PETTYCASH)
            vLTOCASH = N2Str2Zero(rsCash_Pos!LTOCASH)
            vINCASHMENT = N2Str2Zero(rsCash_Pos!INCASHMENT)
            vLTOINCASH = N2Str2Zero(rsCash_Pos!LTOINCASH)
            vCHKPAYMENT = N2Str2Zero(rsCash_Pos!CHKPAYMENT)
            vCHKINCASH = N2Str2Zero(rsCash_Pos!CHKINCASH)
            vCSHPAYMENT = N2Str2Zero(rsCash_Pos!CSHPAYMENT)
            vCSHINCASH = N2Str2Zero(rsCash_Pos!CSHINCASH)
            vBEGIN = vCASH + vCHECK + vCARD
            vEND = 0
            gconDMIS.Execute ("update CMIS_Cash_Pos Set TAG = 1, [END] = " & vBEGIN & " Where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            CURRENT_CUTOFF_DATE = txtCutDate.Text
            gconDMIS.Execute ("Insert into CMIS_Cash_Pos " & _
                              "(CUTDATE,CASH,[CHECK],CARD,REPLENISH,EXPENSE,ADVANCES,CASHDEPO,CHECKDEPO,CARDDEPO,FUND,LTO_REPL,LTO_EXP,LTO_ADV,LTO_AR,LTO,PETPAYMENT,LTOPAYMENT,PETTYCASH,LTOCASH,INCASHMENT,LTOINCASH,CHKPAYMENT,CHKINCASH,CSHPAYMENT,CSHINCASH,[BEGIN],[END])" & _
                            " values ('" & CURRENT_CUTOFF_DATE & "'," & vCASH & "," & vCHECK & "," & vCARD & "," & vREPLENISH & "," & vEXPENSE & "," & vADVANCES & "," & vCASHDEPO & "," & vCHECKDEPO & "," & vCARDDEPO & "," & vFUND & "," & vLTO_REPL & "," & vLTO_EXP & "," & vLTO_ADV & "," & vLTO_AR & "," & vLTO & "," & vPETPAYMENT & "," & vLTOPAYMENT & "," & vPETTYCASH & "," & vLTOCASH & "," & vINCASHMENT & "," & vLTOINCASH & "," & vCHKPAYMENT & "," & vCHKINCASH & "," & vCSHPAYMENT & "," & vCSHINCASH & "," & vBEGIN & "," & vEND & ")")
            gconDMIS.Execute ("Insert into CMIS_Cash " & _
                              "(CUTDATE)" & _
                            " values ('" & CURRENT_CUTOFF_DATE & "')")
            cmdPOST.Enabled = False
            MsgBox "PROCESS COMPLETED.", vbInformation, "Message"
            Screen.MousePointer = 11

            PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Cash_Tally_Sheet_Report.rpt", "{CASH_POS.CUTDATE} = Date(" & Year(REPORT_CUTDATE) & "," & Month(REPORT_CUTDATE) & "," & Day(REPORT_CUTDATE) & ")", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        End If
        'NEW LOG AUDIT-------------------------------------------------
            Call NEW_LogAudit("R", "OFFICIAL RECEIPT CUT-OFF ENTRY", "", "", "", "CUT OFF DATE: " & txtCutDate, "", "")
        'NEW LOG AUDIT-------------------------------------------------
    End If

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub Command1_Click()
    PostOREntries
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtCutDate.Text = CURRENT_CUTOFF_DATE
End Sub

Function CheckUnPostedOR() As Boolean
    Dim rsCheckPostedOR As ADODB.Recordset
    Set rsCheckPostedOR = New ADODB.Recordset
    rsCheckPostedOR.Open "SELECT * FROM CMIS_OFF_HD WHERE PAIDNA=0 AND CANCEL=0", gconDMIS, adOpenKeyset
    If Not rsCheckPostedOR.EOF And Not rsCheckPostedOR.BOF Then
        Do While Not rsCheckPostedOR.EOF
            CheckUnPostedOR = True
            rsCheckPostedOR.MoveNext
        Loop
    End If
    Set rsCheckPostedOR = Nothing
End Function

Sub PrintUnpostedOR()
    Dim xlApplication       As Excel.Application
    Dim xlWorkbook          As Excel.Workbook
    Dim xlWorksheet         As Excel.Worksheet
    Dim xlRange             As Excel.Range
    Dim xCounter            As Integer
    Dim xOR_Amt             As Double
    Dim rsUnpostedOR        As ADODB.Recordset
    Set xlApplication = CreateObject("Excel.Application")
    Set xlWorkbook = xlApplication.Workbooks.Open(CMIS_REPORT_PATH & "UnpostedOR.xlt")
    Set xlWorksheet = xlWorkbook.Worksheets(1)
    xCounter = 7
    xlWorksheet.Cells(1, "A") = COMPANY_NAME
    xlWorksheet.Cells(1, "A").Font.Bold = True
    xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
    xlWorksheet.Cells(2, "A").Font.Bold = True
    xlWorksheet.Cells(3, "A") = "UNPOSTED OR REPORT"
    xlWorksheet.Cells(3, "A").Font.Bold = True
    xlWorksheet.Cells(5, "C") = Format(CDate(txtCutDate.Text), "mmmm dd, yyyy")
    xlWorksheet.Cells(5, "A").Font.Bold = True
    Set rsUnpostedOR = New ADODB.Recordset
    rsUnpostedOR.Open "SELECT * FROM CMIS_OFF_HD WHERE PAIDNA=0 AND CANCEL=0 ORDER BY OR_DATE DESC", gconDMIS, adOpenKeyset
    If Not rsUnpostedOR.EOF And Not rsUnpostedOR.BOF Then
        Do While Not rsUnpostedOR.EOF
            xlWorksheet.Cells(xCounter, "A") = Format(Null2String(rsUnpostedOR!OR_DATE), "mm/dd/yyyy")
            xlWorksheet.Cells(xCounter, "B") = Format(Null2String(rsUnpostedOR!OR_NUM), "000000")
            xlWorksheet.Cells(xCounter, "C") = Null2String(rsUnpostedOR!CUSCDE)
            xlWorksheet.Cells(xCounter, "D") = Null2String(UCase(rsUnpostedOR!CUSNAME))
            xlWorksheet.Cells(xCounter, "E") = NumericVal(rsUnpostedOR!OR_AMT)
            xOR_Amt = xOR_Amt + NumericVal(rsUnpostedOR!OR_AMT)
            xCounter = xCounter + 1
            rsUnpostedOR.MoveNext
        Loop
    End If
    xlWorksheet.Cells(xCounter, "E") = xOR_Amt
    xlApplication.Visible = True
    Set xlApplication = Nothing
    Set rsUnpostedOR = Nothing
End Sub

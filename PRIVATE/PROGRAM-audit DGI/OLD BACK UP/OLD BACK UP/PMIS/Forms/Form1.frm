VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AMIS Checker"
   ClientHeight    =   4950
   ClientLeft      =   885
   ClientTop       =   1245
   ClientWidth     =   9030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command21 
      Caption         =   "Adj"
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   3690
      Width           =   555
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Update SRP"
      Height          =   345
      Left            =   90
      TabIndex        =   20
      Top             =   4530
      Width           =   4515
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Re-Update Master"
      Height          =   375
      Left            =   90
      TabIndex        =   19
      Top             =   4110
      Width           =   4515
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Command18"
      Height          =   435
      Left            =   4080
      TabIndex        =   18
      Top             =   2550
      Width           =   555
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   2235
      Left            =   4050
      TabIndex        =   17
      Top             =   60
      Width           =   585
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Month End"
      Height          =   345
      Left            =   2070
      TabIndex        =   16
      Top             =   3690
      Width           =   1905
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Process Beginning"
      Height          =   615
      Left            =   2070
      TabIndex        =   15
      Top             =   3030
      Width           =   1905
   End
   Begin VB.CommandButton Command14 
      Caption         =   "total qty"
      Height          =   345
      Left            =   90
      TabIndex        =   14
      Top             =   3690
      Width           =   1845
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Update Partmas MAC"
      Height          =   435
      Left            =   2070
      TabIndex        =   13
      Top             =   2550
      Width           =   1905
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   615
      Left            =   90
      TabIndex        =   12
      Top             =   3030
      Width           =   1845
   End
   Begin VB.CommandButton Command11 
      Caption         =   "CHECK RECON TRANS"
      Height          =   495
      Left            =   2070
      TabIndex        =   11
      Top             =   2010
      Width           =   1905
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Check Recon RECEIVING"
      Height          =   465
      Left            =   2070
      TabIndex        =   10
      Top             =   1500
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Check Recon Master"
      Height          =   465
      Left            =   2100
      TabIndex        =   9
      Top             =   990
      Width           =   1905
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Reconcile Master"
      Height          =   405
      Left            =   2100
      TabIndex        =   8
      Top             =   540
      Width           =   1905
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Process Parts Cost"
      Height          =   435
      Left            =   2100
      TabIndex        =   7
      Top             =   60
      Width           =   1905
   End
   Begin VB.CommandButton cmdUpdateREPOR 
      Caption         =   "Update REPOR"
      Height          =   465
      Left            =   60
      TabIndex        =   6
      Top             =   2520
      Width           =   1905
   End
   Begin VB.CommandButton Command6 
      Caption         =   "NBI"
      Height          =   465
      Left            =   60
      TabIndex        =   5
      Top             =   2010
      Width           =   1905
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   465
      Left            =   60
      TabIndex        =   4
      Top             =   1500
      Width           =   1905
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Invalid Customer Code"
      Height          =   465
      Left            =   60
      TabIndex        =   3
      Top             =   990
      Width           =   1905
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   3030
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UnBalance Journal Entry"
      Height          =   405
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Invalid Account Code"
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1905
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   315
      Left            =   4890
      TabIndex        =   22
      Top             =   270
      Width           =   1365
      VariousPropertyBits=   746605595
      DisplayStyle    =   3
      Size            =   "2408;556"
      ColumnCount     =   3
      cColumnInfo     =   3
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "352776;105833;141111"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUpdateREPOR_Click()
    Dim rsOrd_Hd         As ADODB.Recordset
    Set rsOrd_Hd = New ADODB.Recordset
    Set rsOrd_Hd = gconDMIS.Execute("Select * from PMIOS_Ord_Hist Where RONO IS NOT NULL AND Trantype = 'RIV' Order by RONO asc")
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        rsOrd_Hd.MoveFirst
        MsgBox "Press Ok To Begin!"
        Do While Not rsOrd_Hd.EOF
            'If Null2String(rsORD_HD!RONO) <> "" Then
            gconDMIS.Execute "update PMIOS_Ord_Hist set REP_OR = '" & Left(Null2String(rsOrd_Hd!rono), 1) & "-" & Right(Null2String(rsOrd_Hd!rono), 6) & _
                             "' Where Trantype = 'RIV' and Tranno = " & N2Str2Null(rsOrd_Hd!Tranno)
            'End If
            Me.Caption = "Processing Tran. # " & Null2String(rsOrd_Hd!Tranno)
            DoEvents
            rsOrd_Hd.MoveNext
        Loop
        Me.Caption = "Operation Completed!"
        MsgBox "Tapos"
    End If
End Sub

Private Sub Command10_Click()
    frmPMIOSRECON_ReceivingHist.Show
End Sub

Private Sub Command11_Click()
    frmPMIOSRECONCheckDupTrans.Show
End Sub

Private Sub Command12_Click()
    Set rsOrd_Hd = New ADODB.Recordset
    rsOrd_Hd.Open "select id,trantype,tranno,status from PMIOS_Ord_Hist where MONTH(trandate) = 3 AND YEAR(TRANDATE) = 2005 order by trantype,tranno asc", gconDMIS
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        rsOrd_Hd.MoveFirst
        i = 0
        MsgSpeech "Computing Issuances Netcost and Netprice..."
        Me.Caption = "Computing Order Netcost and Netprice..."
        Screen.MousePointer = 11
        DoEvents
        Do While Not rsOrd_Hd.EOF
            vOrdHDRecNo = rsOrd_Hd!ID
            'labProcessing.Caption = "Processing: " & Null2String(rsOrd_Hd!trantype) & " #" & Null2String(rsOrd_Hd!tranno)
            DoEvents
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select id,trantype,tranno,netprice,netcost,status,itemno from PMIOS_DayTran where trantype = " & N2Str2Null(rsOrd_Hd!TRANTYPE) & " and tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " order by itemno asc", gconDMIS
            If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
                rsTDAYTRAN.MoveFirst
                vNetPrice = 0: vNetCost = 0
                Do While Not rsTDAYTRAN.EOF
                    If N2Str2Zero(rsTDAYTRAN!NETprice) > 0 Then
                        vTDNetPrice = rsTDAYTRAN!NETprice
                    Else
                        vTDNetPrice = N2Str2Zero(rsTDAYTRAN!NETprice)
                    End If
                    If N2Str2Zero(rsTDAYTRAN!netcost) > 0 Then
                        vTDNetCost = rsTDAYTRAN!netcost
                    Else
                        vTDNetCost = N2Str2Zero(rsTDAYTRAN!netcost)
                    End If
                    vTDStatus = Null2String(rsTDAYTRAN!Status)
                    If vTDStatus <> "C" Then
                        vNetPrice = vNetPrice + vTDNetPrice
                        vNetCost = vNetCost + vTDNetCost
                    End If
                    rsTDAYTRAN.MoveNext
                Loop
                If Null2String(rsOrd_Hd!Status) <> "C" Then
                    gconDMIS.Execute "update PMIOS_Ord_Hist set netcost = " & vNetCost & ", netinvamt2 = " & vNetPrice & ", status = 'P' where id = " & vOrdHDRecNo
                End If
            End If
            i = i + 1
            'progCPB.Value = (i / rsOrd_Hd.RecordCount) * 100
            'labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsOrd_Hd.MoveNext
        Loop
        'labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If
    MsgBox "tapos"
End Sub

Private Sub Command13_Click()
    Dim rsNEW_DAYTRAN    As ADODB.Recordset
    Dim rsPARTMAS        As ADODB.Recordset
    Dim kim              As Long
    Set rsNEW_DAYTRAN = New ADODB.Recordset
    Set rsNEW_DAYTRAN = gconDMIS.Execute("Select LASTM_OH,ONHAND,STOCKNO,LASTM_MAC,MAC,LASTM_MAD,MAD,INVCLASS,SUBINVCLAS,RESSERVICE,SSTOCK,DATE_ENTERED,RECEIPTS,ISSUANCES from NEW_PARTMAS order by STOCKNO asc")
    If Not rsNEW_DAYTRAN.EOF And Not rsNEW_DAYTRAN.EOF Then
        rsNEW_DAYTRAN.MoveFirst
        kim = 0
        MsgBox "begin"
        Do While Not rsNEW_DAYTRAN.EOF
            kim = kim + 1
            Set rsPARTMAS = New ADODB.Recordset
            Set rsPARTMAS = gconDMIS.Execute("Select STOCKNO from PMIOS_STOCKMAS where STOCKNO = " & N2Str2Null(rsNEW_DAYTRAN!STOCKNO))
            If Not rsPARTMAS.EOF And Not rsPARTMAS.BOF Then
                gconDMIS.Execute "update PMIOS_STOCKMAS Set" & _
                               " ONHAND = " & N2Str2Zero(rsNEW_DAYTRAN!ONHAND) & "," & _
                               " LASTM_OH = " & N2Str2Zero(rsNEW_DAYTRAN!lastm_oh) & "," & _
                               " LASTM_MAC = " & N2Str2Zero(rsNEW_DAYTRAN!lastm_mac) & "," & _
                               " MAC = " & N2Str2Zero(rsNEW_DAYTRAN!Mac) & "," & _
                               " LASTM_MAD = " & N2Str2Zero(rsNEW_DAYTRAN!lastm_mad) & "," & _
                               " MAD = " & N2Str2Zero(rsNEW_DAYTRAN!mad) & "," & _
                               " RECEIPTS = " & N2Str2Zero(rsNEW_DAYTRAN!receipts) & "," & _
                               " ISSUANCES = " & N2Str2Zero(rsNEW_DAYTRAN!issuances) & "," & _
                               " SSTOCK = " & N2Str2Zero(rsNEW_DAYTRAN!SSTOCK) & "," & _
                               " RESSERVICE = " & N2Str2Zero(rsNEW_DAYTRAN!RESSERVICE) & "," & _
                               " DATE_ENTERED = " & N2Str2Null(rsNEW_DAYTRAN!DATE_ENTERED) & "," & _
                               " INVCLASS = " & N2Str2Null(rsNEW_DAYTRAN!InvClass) & "," & _
                               " SUBINVCLAS = " & N2Str2Null(rsNEW_DAYTRAN!SubInvClas) & _
                               " WHERE STOCKNO = " & N2Str2Null(rsNEW_DAYTRAN!STOCKNO)
            Else
Stop
                gconDMIS.Execute "INSERT into PMIOS_STOCKMAS " & _
                                 "(STOCKNO,STOCKDESC,ONHAND,MAC,LASTM_MAC,MAD,LASTM_MAD)" & _
                               " VALUES (" & N2Str2Null(rsNEW_DAYTRAN!STOCKNO) & "," & N2Str2Null(rsNEW_DAYTRAN!STOCKDESC) & "," & N2Str2Null(rsNEW_DAYTRAN!ONHAND) & "," & N2Str2Null(rsNEW_DAYTRAN!Mac) & "," & N2Str2Null(rsNEW_DAYTRAN!lastm_mac) & "," & N2Str2Null(rsNEW_DAYTRAN!mad) & "," & N2Str2Null(rsNEW_DAYTRAN!lastm_mad) & ")"
Stop
            End If
            rsNEW_DAYTRAN.MoveNext
            'Me.Caption = Null2String(rsNEW_DAYTRAN!STOCKNO)
            DoEvents
        Loop
        MsgBox "tapos!"
    End If
End Sub

Private Sub Command14_Click()
    Dim rsOrd_Hd         As ADODB.Recordset
    Dim rdTDAYTRAN       As ADODB.Recordset
    Dim i                As Integer
    Dim total_value      As Long
    Dim vOrdHDRecNo, vTotalQty As Long
    Set rsOrd_Hd = New ADODB.Recordset
    rsOrd_Hd.Open "select id,trantype,tranno,status from PMIOS_Ord_Hist  order by trantype,tranno asc", gconDMIS
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        rsOrd_Hd.MoveFirst
        i = 0
        MsgSpeech "Computing Total Quantity of Issuances..."
        Me.Caption = "Computing Total Quantity of Issuances..."
        Screen.MousePointer = 11
        DoEvents
        Do While Not rsOrd_Hd.EOF
            vOrdHDRecNo = rsOrd_Hd!ID
            'labProcessing.Caption = "Processing: " & Null2String(rsORD_HD!trantype) & " #" & Null2String(rsORD_HD!tranno)
            DoEvents
            Set rsTDAYTRAN = New ADODB.Recordset
            rsTDAYTRAN.Open "select id,trantype,tranno,tranqty,status,itemno from PMIOS_DayTran where trantype = " & N2Str2Null(rsOrd_Hd!TRANTYPE) & " and tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " order by itemno asc", gconDMIS
            If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
                rsTDAYTRAN.MoveFirst
                vTotalQty = 0
                Do While Not rsTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(rsTDAYTRAN!tranqty)
                    rsTDAYTRAN.MoveNext
                Loop
                If Null2String(rsOrd_Hd!Status) <> "C" Then
                    gconDMIS.Execute "update PMIOS_Ord_Hist set TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
                End If
            End If
            i = i + 1
            total_value = (i / rsOrd_Hd.RecordCount) * 100
            Me.Caption = Int(total_value) & "% Completed"
            DoEvents
            rsOrd_Hd.MoveNext
        Loop
        'labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Command15_Click()

    Dim vCATEGORY        As String
    Dim vSTOCKNO         As String
    Dim vSTOCKDESC       As String
    Dim vMODELCODE       As String
    Dim vLOCATION        As String
    Dim vMAC             As Double
    Dim vMAD             As Double
    Dim vSRP             As Double
    Dim vNOSHIP          As Double
    Dim vLASTM_MAC       As Double
    Dim vLASTM_MAD       As Double
    Dim vLASTM_SELL      As Double
    Dim vLASTM_OH        As Double
    Dim vLASTM_OO        As Double
    Dim vONHAND          As Double
    Dim vTRECQTY         As Double
    Dim vTISSQTY         As Double
    Dim vONORDER         As Double
    Dim vPOQTY           As Double
    Dim vTPOQTY          As Double
    Dim vTPPQTY          As Double
    Dim vPRQTY           As Double
    Dim vTPRQTY          As Double
    Dim vLAST_RECQ       As String
    Dim vLAST_RECD       As String
    Dim vLASTY_OH        As Double
    Dim vLASTY_MAC       As Double
    Dim vLASTY_OO        As Double
    Dim vLASTY_ADJ       As Double
    Dim vSupCode         As String
    Dim vRECEIPTS        As Double
    Dim vISSUANCES       As Double
    Dim vDNP             As Double

    vCATEGORY = ""
    vSTOCKNO = ""
    vSTOCKDESC = ""
    vMODELCODE = ""
    vLOCATION = ""
    vMAC = 0
    vMAD = 0
    vSRP = 0
    vNOSHIP = 0
    vLASTM_MAC = 0
    vLASTM_MAD = 0
    vLASTM_SELL = 0
    vLASTM_OH = 0
    vLASTM_OO = 0
    vONHAND = 0
    vTRECQTY = 0
    vTISSQTY = 0
    vONORDER = 0
    vPOQTY = 0
    vTPOQTY = 0
    vTPPQTY = 0
    vPRQTY = 0
    vTPRQTY = 0
    vLAST_RECQ = 0
    vLAST_RECD = ""
    vLASTY_OH = 0
    vLASTY_MAC = 0
    vLASTY_OO = 0
    vLASTY_ADJ = 0
    vSupCode = ""
    vRECEIPTS = 0
    vISSUANCES = 0
    vDNP = 0

    Dim rsBEGPART        As ADODB.Recordset
    'Set RSBEGPART = New ADODB.Recordset
    'Set RSBEGPART = gconDMIS.Execute("SELECT * from BEGPART ORDER BY ID ASC")
    'If Not RSBEGPART.EOF And Not RSBEGPART.BOF Then
    '   RSBEGPART.MoveFirst
    '   Do While Not RSBEGPART.EOF
    '      vCATEGORY = N2Str2Null(RSBEGPART!Category)
    '      vSTOCKNO = N2Str2Null(RSBEGPART!STOCKNO)
    '      vSTOCKDESC = N2Str2Null(RSBEGPART!STOCKDESC)
    '      vMODELCODE = N2Str2Null(RSBEGPART!modelcode)
    '      vLOCATION = N2Str2Null(RSBEGPART!location)
    '      vMAC = N2Str2Zero(RSBEGPART!Mac)
    '      'vMAD = N2Str2Zero(RSBEGPART!mad)
    '      vSRP = N2Str2Zero(RSBEGPART!SRP)
    '      'vNOSHIP = N2Str2Zero(rsBEGPART!noship)
    '      'vLASTM_MAC = N2Str2Zero(rsBEGPART!lastm_Mac)
    '      'vLASTM_MAD = N2Str2Zero(rsBEGPART!lastm_mad)
    '      'vLASTM_SELL = N2Str2Zero(rsBEGPART!lastm_sell)
    '     'vLASTM_OH = N2Str2Zero(rsBEGPART!lastm_oh)
    '     'vLASTM_OO = N2Str2Zero(rsBEGPART!lastm_oo)
    '     vONHAND = N2Str2Zero(RSBEGPART!onhand)
    '      vTRECQTY = N2Str2Zero(RSBEGPART!JAN_R)
    '      vTISSQTY = N2Str2Zero(RSBEGPART!JAN_I)
    '      'vONORDER = N2Str2Zero(rsBEGPART!onorder)
    '     'vPOQTY = N2Str2Zero(rsBEGPART!poqty)
    '     'vTPOQTY = N2Str2Zero(rsBEGPART!tpoqty)
    '     'vTPPQTY = N2Str2Zero(rsBEGPART!tppqty)
    '     'vPRQTY = N2Str2Zero(rsBEGPART!prqty)
    ''     'vTPRQTY = N2Str2Zero(rsBEGPART!tprqty)
    '     vLAST_RECQ = N2Str2Zero(RSBEGPART!JAN_R)
    '     vLAST_RECD = N2Str2Null(RSBEGPART!last_recd)
    '    'vLASTY_OH = N2Str2Zero(rsBEGPART!lasty_oh)
    '    'vLASTY_MAC = N2Str2Zero(rsBEGPART!lasty_mac)
    '    'vLASTY_OO = N2Str2Zero(rsBEGPART!lasty_oo)
    '    'vLASTY_ADJ = N2Str2Zero(rsBEGPART!lasty_adj)
    '    'vSupCode = N2Str2Null(RSBEGPART!supcode)
    '    vRECEIPTS = N2Str2Zero(RSBEGPART!JAN_R)
    ''    vISSUANCES = N2Str2Zero(RSBEGPART!JAN_I)
    '    vDNP = N2Str2Zero(RSBEGPART!DNP)
    '    On Error GoTo kim
    '      gconDMIS.Execute ("Insert into PMIOS_STOCKMAS (DEALER_TYPE,CATEGORY,STOCKNO,STOCKDESC,MODELCODE,LOCATION,MAC,SRP,ONHAND,TRECQTY,TISSQTY,LAST_RECQ,LAST_RECD,LASTY_OH,LASTY_MAC,RECEIPTS,ISSUANCES,DNP)" & _
           '                         " values ('1'," & vCATEGORY & "," & vSTOCKNO & "," & vSTOCKDESC & "," & vMODELCODE & "," & vLOCATION & "," & vMAC & "," & vSRP & "," & vONHAND & "," & vTRECQTY & "," & vTISSQTY & "," & vLAST_RECQ & "," & vLAST_RECD & "," & vLASTY_OH & "," & vLASTY_MAC & "," & vRECEIPTS & "," & vISSUANCES & "," & vDNP & ")")
    '      RSBEGPART.MoveNext
    '   Loop
    '   Exit Sub
    Dim vTrandate        As String
    Dim vTRANTYPE        As String
    Dim vTRANNO          As String
    Dim vITEMNO          As String
    Dim vSTOCK_ORD       As String
    Dim vSTOCK_SUP       As String
    Dim vTranQty         As String
    Dim vTRANUCOST       As Double
    Dim vTRANUPRICE      As Double
    Dim VStatus          As String
    Dim vIN_OUT          As String
    Dim vTRANINVAMT      As Double
    Dim rsPARTMAS        As ADODB.Recordset
    'Set rsBEGPART = New ADODB.Recordset
    'Set rsBEGPART = gconDMIS.Execute("Select * from [JULbeg] ORDER BY STOCKNO ASC")
    'rsBEGPART.MoveFirst
    'Do While Not rsBEGPART.EOF
    '   vTRANDATE = "'7/31/2006'"
    '   vTRANTYPE = "'BEG'"
    '   vTRANNO = "'000000'"
    '   vITEMNO = "'0000'"
    '   vSTOCK_ORD = N2Str2Null(rsBEGPART!STOCKNO)
    '   vSTOCK_SUP = N2Str2Null(rsBEGPART!STOCKNO)
    '   vTRANQTY = N2Str2Zero(rsBEGPART!ONHAND)
    '   vTRANUCOST = N2Str2Zero(rsBEGPART!DNP)
    '   vSTATUS = "'P'"
    '   vIN_OUT = "'I'"
    '   vTRANINVAMT = N2Str2Zero(rsBEGPART!DNP)
    '   If vTRANQTY > 0 Then
    '   Set RSPARTMAS = New ADODB.Recordset
    '   Set RSPARTMAS = gconDMIS.Execute("SELECT * from PMIOS_STOCKMAS where STOCKNO = " & vSTOCK_ORD)
    '   If RSPARTMAS.EOF And RSPARTMAS.BOF Then
    '      Me.Caption = "ADDING STOCKNO": DoEvents
    '      gconDMIS.Execute ("INSERT into PMIOS_STOCKMAS " & _
           '                        "(DEALER_TYPE,STOCKNO,STOCKDESC,MODELCODE,LOCATION,DNP,MAC,ONHAND,LASTM_OH,NON_HARI)" & _
           '                        " VALUES ('1'," & vSTOCK_ORD & "," & N2Str2Null(rsBEGPART!STOCKDESC) & "," & N2Str2Null(rsBEGPART!MODELCODE) & "," & N2Str2Null(rsBEGPART!LOCATION) & "," & N2Str2Zero(rsBEGPART!DNP) & "," & N2Str2Zero(rsBEGPART!DNP) & "," & vTRANQTY & "," & vTRANQTY & ",'Y')")
    '   End If
    '   Me.Caption = "INSERTING TRANSACTION": DoEvents
    '   gconDMIS.Execute ("Insert into PMIOS_DayTran (DEALER_TYPE,TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,STATUS,IN_OUT,TRANINVAMT,NON_HARI)" & _
        '                      " VALUES ('1'," & vTRANDATE & "," & vTRANTYPE & "," & vTRANNO & "," & vITEMNO & "," & vSTOCK_ORD & "," & vSTOCK_SUP & "," & vTRANQTY & "," & vTRANUCOST & "," & vSTATUS & "," & vIN_OUT & "," & vTRANINVAMT & ",'N')")
    '   End If
    '   rsBEGPART.MoveNext
    'Loop
    'MsgBox "tapos"
    'Exit Sub

    Set rsBEGPART = New ADODB.Recordset
    Set rsBEGPART = gconDMIS.Execute("Select * from Aug23_26_3")
    'rsBEGPART.MoveFirst
    'Do While Not rsBEGPART.EOF
    '   vTRANDATE = "'8/1/2006'"
    '   vTRANTYPE = "'IN'"
    '   vTRANNO = "'000000'"
    '   vITEMNO = "'0000'"
    '   vSTOCK_ORD = N2Str2Null(rsBEGPART!n2)
    '  vSTOCK_SUP = N2Str2Null(rsBEGPART!n2)
    '   vTRANQTY = N2Str2Zero(rsBEGPART!n10)
    '   vTRANUCOST = N2Str2Zero(rsBEGPART!n6)
    '   vSTATUS = "'P'"
    '   vIN_OUT = "'I'"
    '   vTRANINVAMT = N2Str2Zero(rsBEGPART!n6)
    '  If vTRANQTY > 0 Then
    '  Set RSPARTMAS = New ADODB.Recordset
    '   Set RSPARTMAS = gconDMIS.Execute("SELECT * from PMIOS_STOCKMAS where STOCKNO = " & vSTOCK_ORD)
    '   If RSPARTMAS.EOF And RSPARTMAS.BOF Then
    '     Me.Caption = "ADDING STOCKNO": DoEvents
    '     gconDMIS.Execute ("INSERT into PMIOS_STOCKMAS " & _
          '                       "(DEALER_TYPE,STOCKNO,STOCKDESC,MODELCODE,LOCATION,DNP,MAC,SRP,NON_HARI)" & _
          ''                       " VALUES ('1'," & vSTOCK_ORD & "," & N2Str2Null(rsBEGPART!n3) & "," & N2Str2Null(rsBEGPART!n4) & "," & N2Str2Null(rsBEGPART!n5) & "," & N2Str2Zero(rsBEGPART!n6) & "," & N2Str2Zero(rsBEGPART!n6) & "," & N2Str2Zero(rsBEGPART!n7) & ",'N')")
    '  End If
    '   Me.Caption = "INSERTING TRANSACTION": DoEvents
    '   gconDMIS.Execute ("Insert into PMIOS_TdayTran (DEALER_TYPE,TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,STATUS,IN_OUT,TRANINVAMT)" & _
        '                      " VALUES ('1'," & vTRANDATE & "," & vTRANTYPE & "," & vTRANNO & "," & vITEMNO & "," & vSTOCK_ORD & "," & vSTOCK_SUP & "," & vTRANQTY & "," & vTRANUCOST & "," & vSTATUS & "," & vIN_OUT & "," & vTRANINVAMT & ")")
    '   End If
    '   rsBEGPART.MoveNext
    'Loop
    '
    rsBEGPART.MoveFirst
    Do While Not rsBEGPART.EOF
        vTrandate = "'" & Null2String(rsBEGPART!n12) & "'"
        'vTRANDATE = "'" & firstDay(LOGDATE) & "'"
        vTRANTYPE = "'OUT'"
        vTRANNO = "'000000'"
        vITEMNO = "'0000'"
        vSTOCK_ORD = N2Str2Null(rsBEGPART!n2)
        vSTOCK_SUP = N2Str2Null(rsBEGPART!n2)
        vTranQty = N2Str2Zero(rsBEGPART!n11)
        VStatus = "'P'"
        vIN_OUT = "'O'"
        If vTranQty > 0 Then
            Set rsPARTMAS = New ADODB.Recordset
            Set rsPARTMAS = gconDMIS.Execute("SELECT * from PMIOS_STOCKMAS where STOCKNO = " & vSTOCK_ORD)
            If rsPARTMAS.EOF And rsPARTMAS.BOF Then
            Else
                vTRANUPRICE = N2Str2Zero(rsPARTMAS!SRP)
                vTRANINVAMT = N2Str2Zero(rsPARTMAS!SRP)
            End If
            Me.Caption = "INSERTING TRANSACTION": DoEvents
            gconDMIS.Execute ("Insert into PMIOS_TdayTran (TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUPRICE,STATUS,IN_OUT)" & _
                            " VALUES (" & vTrandate & "," & vTRANTYPE & "," & vTRANNO & "," & vITEMNO & "," & vSTOCK_ORD & "," & vSTOCK_SUP & "," & vTranQty & "," & vTRANUPRICE & "," & VStatus & "," & vIN_OUT & ")")
        End If
        'Stop
        rsBEGPART.MoveNext
    Loop
    'End If
    MsgBox "tapos"
    Exit Sub
kim:
    MsgBox Err.Description
    If MsgBox("COntinue?", vbQuestion + vbYesNo, "Decide") = vbNo Then
        End
    Else
        Resume Next
    End If
End Sub

Private Sub Command16_Click()
    Dim i                As Integer
    Dim rsTDAYTRAN       As ADODB.Recordset

    Dim rsPARTMAS, rsCURPartmas, rsRR_HD As ADODB.Recordset
    Dim vTotTranCost, vTotTranInvAmt, vTotTranQTY, vTDTranQTY As Double
    Dim vTDTranType, vTDTranno, vSupCode As String
    Dim vVatAmt, vMAC    As Double
    Dim vPMOnhand        As Integer
    Dim vSTOCKDESC       As String
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select id,ItemNo,trantype,tranno,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIOS_TdayTran where trantype <> 'ADB' and status <> 'C' and status <> 'N' order by id asc", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        rsTDAYTRAN.MoveFirst
        Screen.MousePointer = 11
        DoEvents
        Me.Caption = "Updating Part Master File"
        gconDMIS.Execute "update PMIOS_STOCKMAS set" & _
                       " lastm_mac = PMIOS_STOCKMAS.mac " & _
                       " where lastm_mac = 0 and onhand > 0"
        If Month(LOGDATE) = 1 Then
            gconDMIS.Execute "update PMIOS_STOCKMAS set" & _
                           " onhand = PMIOS_STOCKMAS.lastm_oh," & _
                           " mac = PMIOS_STOCKMAS.lastm_mac," & _
                           " onorder = PMIOS_STOCKMAS.lastm_oo," & _
                           " tissqty = 0," & _
                           " trecqty = 0," & _
                           " receipts = 0," & _
                           " issuances = 0"
        Else
            gconDMIS.Execute "update PMIOS_STOCKMAS set" & _
                           " onhand = PMIOS_STOCKMAS.lastm_oh," & _
                           " mac = PMIOS_STOCKMAS.lastm_mac," & _
                           " onorder = PMIOS_STOCKMAS.lastm_oo," & _
                           " tissqty = 0," & _
                           " trecqty = 0"
        End If
        DoEvents
        Me.Caption = "Updating Transactions to Part Master File"
        DoEvents
        i = 0
        Do While Not rsTDAYTRAN.EOF
            'gconDMIS.Execute "update PMIOS_TdayTran set ItemNo = '" & Format(Null2String(rsTDAYTRAN!itemno), "0000") & "' where ID = " & rsTDAYTRAN!ID
            vTDTranDate = N2Date2Null(rsTDAYTRAN!trandate)
            vTDTranType = Null2String(rsTDAYTRAN!TRANTYPE)
            vTDTranno = Null2String(rsTDAYTRAN!Tranno)
            vTDTranQTY = N2Str2IntZero(rsTDAYTRAN!tranqty)
            If N2Str2Zero(rsTDAYTRAN!TRANUCOST) > 0 Then
                vTotTranCost = rsTDAYTRAN!TRANUCOST * vTDTranQTY
            Else
                vTotTranCost = 0
            End If
            'vTotTranCost = N2Str2Zero(rsTDAYTRAN!tranucost) * vTDTranQTY
            If N2Str2Zero(rsTDAYTRAN!TRANINVAMT) > 0 Then
                vTotTranInvAmt = rsTDAYTRAN!TRANINVAMT * vTDTranQTY
            Else
                vTotTranInvAmt = 0
            End If
            'vTotTranInvAmt = N2Str2Zero(rsTDAYTRAN!traninvamt * vTDTranQTY)
            Me.Caption = "Processing: " & Null2String(rsTDAYTRAN!TRANTYPE) & " #" & Null2String(rsTDAYTRAN!Tranno)
            DoEvents
            Set rsPARTMAS = New ADODB.Recordset
            Set rsPARTMAS = gconDMIS.Execute("select STOCKNO from PMIOS_STOCKMAS where STOCKNO = " & N2Str2Null(rsTDAYTRAN!STOCK_ORD))
            If rsPARTMAS.EOF And rsPARTMAS.BOF Then
                Set rsCURPartmas = New ADODB.Recordset
                Set rsCURPartmas = gconDMIS.Execute("Select STOCKNO,STOCKDESC from PMIOS_STOCKMAS where STOCKNO = " & N2Str2Null(rsTDAYTRAN!STOCK_ORD))
                If Not rsCURPartmas.EOF And Not rsCURPartmas.BOF Then
                    vSTOCKDESC = N2Str2Null(rsCURPartmas!STOCKDESC)
                Else
                    vSTOCKDESC = "'NO DESCRIPTION'"
                End If
                gconDMIS.Execute ("Insert into PMIOS_STOCKMAS (STOCKNO,STOCKDESC,date_entered) values (" & N2Str2Null(rsTDAYTRAN!STOCK_ORD) & "," & vSTOCKDESC & "," & N2Str2Null(rsTDAYTRAN!trandate) & ")")
            End If
            Set rsPARTMAS = New ADODB.Recordset
            rsPARTMAS.Open "select id,STOCKNO,mac,dnp,tissqty,trecqty,poqty,tprqty,prqty,issuances,receipts,onhand from PMIOS_STOCKMAS where STOCKNO = " & N2Str2Null(rsTDAYTRAN!STOCK_ORD), gconDMIS
            If Not rsPARTMAS.EOF And Not rsPARTMAS.BOF Then
                If N2Str2Zero(rsPARTMAS!Mac) > 0 Then vMAC = rsPARTMAS!Mac Else vMAC = rsPARTMAS!DNP
                'vMAC = N2Str2Zero(rsPartmas!MAC)
                vPMOnhand = N2Str2IntZero(rsPARTMAS!ONHAND)
                If Null2String(rsTDAYTRAN!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    gconDMIS.Execute "update PMIOS_STOCKMAS set " & _
                                     "onhand = " & vPMOnhand - vTDTranQTY & ", " & _
                                     "tissqty = " & N2Str2IntZero(rsPARTMAS!TISSQTY) + vTDTranQTY & ", " & _
                                     "issuances = " & N2Str2IntZero(rsPARTMAS!issuances) + vTDTranQTY & _
                                   " where id = " & rsPARTMAS!ID
                    If vMAC = 0 Then
                        'vMAC = 131
                        'Stop
                        MsgBox "Error Encountered on Part Number (" & Null2String(rsTDAYTRAN!STOCK_ORD) & ")" & vbCrLf & " Contact the wizweirdo immediately.", vbCritical, "Warning"
                        'Exit Sub
                        'MsgBox Null2String(rsPARTMAS!STOCKNO)
                        'vMAC = 131
                        'Stop
                        'Set rsSupplier = New ADODB.Recordset
                        'Set rsSupplier = gconDMIS.Execute("Select mac from PMIOS_TdayTran where month(trandate) = " & Month(txtFrom.Text) & " and year(trandate) = " & Year(txtFrom.Text) & " and trantype = 'RR' and STOCK_ORD = " & N2Str2Null(rsPARTMAS!STOCKNO) & " order by trandate asc")
                        'If Not rsSupplier.EOF And Not rsSupplier.BOF Then
                        '  If rsSupplier!MAC <> 0 Then
                        '      vMAC = rsSupplier!MAC
                        '      Stop
                        '   Else
                        '      Stop
                        '   End If
                        '  'Stop
                        'Else
                        '   Stop
                        'End If
                        'Stop
                        'vMAC = 2272.73
                    End If
                    gconDMIS.Execute "update PMIOS_TdayTran set tranucost = " & vMAC & " where ID = " & rsTDAYTRAN!ID
                End If

                If Null2String(rsTDAYTRAN!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    If vPMOnhand <= 0 Then
                        vMAC = vTotTranCost / vTDTranQTY
                        gconDMIS.Execute "update PMIOS_STOCKMAS set " & _
                                         "mac = " & vMAC & ", " & _
                                         "Onhand = " & vPMOnhand + vTDTranQTY & ", " & _
                                         "last_recd = " & vTDTranDate & ", " & _
                                         "trecqty = " & N2Str2IntZero(rsPARTMAS!trecqty) + vTDTranQTY & ", " & _
                                         "receipts = " & N2Str2IntZero(rsPARTMAS!receipts) + vTDTranQTY & _
                                       " where id =" & rsPARTMAS!ID
                    Else
                        vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                        gconDMIS.Execute "update PMIOS_STOCKMAS set " & _
                                         "mac = " & vMAC & ", " & _
                                         "Onhand = " & vPMOnhand + vTDTranQTY & ", " & _
                                         "last_recd = " & vTDTranDate & ", " & _
                                         "trecqty = " & N2Str2IntZero(rsPARTMAS!trecqty) + vTDTranQTY & ", " & _
                                         "receipts = " & N2Str2IntZero(rsPARTMAS!receipts) + vTDTranQTY & _
                                       " where id =" & rsPARTMAS!ID
                    End If
                    gconDMIS.Execute "update PMIOS_TdayTran set mac = " & vMAC & ", status = 'P' where id = " & rsTDAYTRAN!ID
                End If

            End If
            DoEvents
            i = i + 1
            'progCPB.Value = (i / rsTDAYTRAN.RecordCount) * 100
            'labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTDAYTRAN.MoveNext
        Loop
        'labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "Error Opening TTDAYTRAN File"
        Exit Sub
    End If

    'end update master
    Set rsTDAYTRAN = New ADODB.Recordset
    'rsTDAYTRAN.Open "Select id,in_out,trantype,tranno,STOCK_ORD,status,tranqty,netcost,tranucost,trandate,tranuprice,traninvamt from NEW_daytran where status <> 'C'  AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by trandate asc,trantype desc,tranno asc,itemno asc", gconDMIS
    rsTDAYTRAN.Open "Select id,in_out,trantype,tranno,STOCK_ORD,status,tranqty,netcost,tranucost,trandate,tranuprice,traninvamt from PMIOS_TdayTran where status <> 'C' and status <> 'N' order by id asc", gconDMIS
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        rsTDAYTRAN.MoveFirst
        i = 0
        Screen.MousePointer = 11
        MsgSpeech "Posting Transactions from Daily Transactions File..."
        Me.Caption = "Posting Transactions from PMIOS_TdayTran File..."
        DoEvents
        Do While Not rsTDAYTRAN.EOF
            vTDRecNo = rsTDAYTRAN!ID
            vTDInOut = Null2String(rsTDAYTRAN!IN_OUT)
            vTDTranType = Null2String(rsTDAYTRAN!TRANTYPE)
            vTDTranno = Null2String(rsTDAYTRAN!Tranno)
            vTDPartOrd = Null2String(rsTDAYTRAN!STOCK_ORD)
            vTDStatus = Null2String(rsTDAYTRAN!Status)
            vTDTranQTY = N2Str2IntZero(rsTDAYTRAN!tranqty)
            If N2Str2Zero(rsTDAYTRAN!netcost) > 0 Then
                vTDNetCost = rsTDAYTRAN!netcost
            Else
                vTDNetCost = N2Str2Zero(rsTDAYTRAN!netcost)
            End If
            If N2Str2Zero(rsTDAYTRAN!TRANUCOST) > 0 Then
                vTDTranucost = rsTDAYTRAN!TRANUCOST
            Else
                vTDTranucost = N2Str2Zero(rsTDAYTRAN!TRANUCOST)
            End If
            If N2Str2Zero(rsTDAYTRAN!TRANINVAMT) > 0 Then
                vTDTranInvAmt = rsTDAYTRAN!TRANINVAMT
            Else
                vTDTranInvAmt = N2Str2Zero(rsTDAYTRAN!TRANINVAMT)
            End If
            vTDTranDate = Null2Date(rsTDAYTRAN!trandate)
            vTotTranCost = vTDTranucost * vTDTranQTY
            vTDTranuprice = N2Str2Zero(rsTDAYTRAN!TRANUPRICE)
            'labProcessing.Caption = "Processing: " & vTDTranType & " #" & vTDTranno
            DoEvents
            Set rsPARTMAS = New ADODB.Recordset
            rsPARTMAS.Open "Select id,onhand,trecqty,last_recd,receipts,tissqty,issuances,lastm_MAC,MAC from PMIOS_STOCKMAS where STOCKNO = '" & vTDPartOrd & "'", gconDMIS
            If Not rsPARTMAS.EOF And Not rsPARTMAS.BOF Then
                If vTDTranType = "IN" Then
                    vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(vVatAmt)
                    vPMRecNo = rsPARTMAS!ID
                    'vPMOnhand = N2Str2IntZero(rsPARTMAS!Onhand)
                    vPMTrecqty = N2Str2IntZero(rsPARTMAS!trecqty)
                    vPMLast_Recd = Null2Date(rsPARTMAS!last_recd)
                    'vPMReceipts = N2Str2IntZero(rsPARTMAS!RECEIPTS)
                    'vMAC = N2Str2Zero(rsPARTMAS!MAC)
                    gconDMIS.Execute "update PMIOS_STOCKMAS set " & _
                                     "trecqty = " & vPMTrecqty - vTDTranQTY & ", " & _
                                     "last_recd = " & N2Str2Null(vTDTranDate) & _
                                   " where id =" & vPMRecNo
                    gconDMIS.Execute "update PMIOS_TdayTran set status = 'P' where id = " & vTDRecNo
                    'gconDMIS.Execute "update PMIOS_TdayTran set mac = " & vMAC & ", status = 'P', netcost = " & vTotTranCost & " where id = " & vTDRecNo
                End If
                If vTDTranType = "OUT" Then
                    vORDTotPrice = (vTDTranuprice * vTDTranQTY) / ConvertToBIRDecimalFormat(VAT_RATE)
                    vPMRecNo = rsPARTMAS!ID
                    vPMTissqty = N2Str2IntZero(rsPARTMAS!TISSQTY)
                    'vPMIssuances = N2Str2IntZero(rsPARTMAS!ISSUANCES)
                    'vMAC = N2Str2Zero(rsPARTMAS!MAC)
                    vTotTranCost = vTDTranucost * vTDTranQTY
                    gconDMIS.Execute "update PMIOS_STOCKMAS set " & _
                                     "tissqty = " & vPMTissqty - vTDTranQTY & _
                                   " where id =" & vPMRecNo
                    gconDMIS.Execute "update PMIOS_TdayTran set netcost = " & vTotTranCost & ", netprice = " & vORDTotPrice & ", status = 'P' where id = " & vTDRecNo
                    'gconDMIS.Execute "update PMIOS_TdayTran set tranucost = " & vMAC & ", netcost = " & vTotTranCost & ", netprice = " & vORDTotPrice & ", status = 'P' where id = " & vTDRecNo
                    Set rsShipping = New ADODB.Recordset
                    rsShipping.Open "select * from PMIOS_Shipping where STOCKNO = '" & vTDPartOrd & "'", gconDMIS
                    If Not rsShipping.EOF And Not rsShipping.BOF Then
                        vShRecNo = rsShipping!ID
                        vShCurrMonth = N2Str2IntZero(rsShipping!curr_month)
                        gconDMIS.Execute "update PMIOS_Shipping set curr_month = " & vShCurrMonth + vTDTranQTY & ", " & _
                                         "freq_curr = 1 where id = " & vShRecNo
                    Else
                        gconDMIS.Execute "insert into PMIOS_Shipping (STOCKNO,curr_month,freq_curr)" & _
                                       " values ('" & vTDPartOrd & "', " & vTDTranQTY & ", 1)"
                    End If
                End If



            Else
                If vTDTranType <> "ADB" Then
                    gconDMIS.Execute "insert into PMIOS_No_Mstr " & _
                                     "(trantype,tranno,recno)" & _
                                   " values ('" & vTDInOut & "', '" & vTDTranno & "', " & vTDRecNo & ")"
                    MsgSpeechBox "Error in " & vTDTranType & "-" & vTDTranno & " Part Number: " & vTDPartOrd & " is not in Master File"
                End If
            End If
            i = i + 1
            'progCPB.Value = (i / rsTDAYTRAN.RecordCount) * 100
            'labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTDAYTRAN.MoveNext
        Loop
        'labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    End If

    gconDMIS.Execute "update PMIOS_STOCKMAS set" & _
                   " PMIOS_STOCKMAS.lastm_oh = PMIOS_STOCKMAS.onhand," & _
                   " PMIOS_STOCKMAS.lastm_mac = PMIOS_STOCKMAS.Mac," & _
                   " PMIOS_STOCKMAS.lastm_mad = PMIOS_STOCKMAS.Mad," & _
                   " PMIOS_STOCKMAS.lastm_oo = PMIOS_STOCKMAS.onorder," & _
                   " PMIOS_STOCKMAS.noship = PMIOS_STOCKMAS.noship + 1," & _
                   " PMIOS_STOCKMAS.mad = (M_shipping.Curr_Month + M_shipping.Prev_Month + M_shipping.Months_2 + M_shipping.Months_3 + M_shipping.Months_4 + M_shipping.Months_5) / 6 from M_shipping" & _
                   " where M_shipping.Curr_Month <= 0 and PMIOS_STOCKMAS.STOCKNO = M_shipping.STOCKNO"

    gconDMIS.Execute "update PMIOS_STOCKMAS set" & _
                   " PMIOS_STOCKMAS.lastm_oh = PMIOS_STOCKMAS.onhand," & _
                   " PMIOS_STOCKMAS.lastm_mac = PMIOS_STOCKMAS.Mac," & _
                   " PMIOS_STOCKMAS.lastm_mad = PMIOS_STOCKMAS.Mad," & _
                   " PMIOS_STOCKMAS.lastm_oo = PMIOS_STOCKMAS.onorder," & _
                   " PMIOS_STOCKMAS.noship = 0," & _
                   " PMIOS_STOCKMAS.mad = (M_shipping.Curr_Month + M_shipping.Prev_Month + M_shipping.Months_2 + M_shipping.Months_3 + M_shipping.Months_4 + M_shipping.Months_5) / 6 from M_shipping" & _
                   " where M_shipping.Curr_Month > 0"

    gconDMIS.Execute "update M_shipping set" & _
                   " months_60 = Months_59, months_59 = Months_58, months_58 = Months_57, months_57 = Months_56," & _
                   " months_56 = Months_55, months_55 = Months_54, months_54 = Months_53, months_53 = Months_52," & _
                   " months_52 = Months_51, months_51 = Months_50, months_50 = Months_49, months_49 = Months_48," & _
                   " months_48 = Months_47, months_47 = Months_46, months_46 = Months_45, months_45 = Months_44," & _
                   " months_44 = Months_43, months_43 = Months_42, months_42 = Months_41, months_41 = Months_40," & _
                   " months_40 = Months_39, months_39 = Months_38, months_38 = Months_37, months_37 = Months_36," & _
                   " months_36 = Months_35, months_35 = Months_34, months_34 = Months_33, months_33 = Months_32," & _
                   " months_32 = Months_31, months_31 = Months_30, months_30 = Months_29, months_29 = Months_28," & _
                   " months_28 = Months_27, months_27 = Months_26, months_26 = Months_25, months_25 = Months_24," & _
                   " months_24 = Months_23, months_23 = Months_22, months_22 = Months_21, months_21 = Months_20," & _
                   " months_20 = Months_19, months_19 = Months_18, months_18 = Months_17, months_17 = Months_16," & _
                   " months_16 = Months_15, months_15 = Months_14, months_14 = Months_13, months_13 = Months_12," & _
                   " months_12 = Months_11, months_11 = Months_10, months_10 = Months_9, months_9 = Months_8," & _
                   " months_8 = Months_7, months_7 = Months_6, months_6 = Months_5, months_5 = Months_4," & _
                   " months_4 = Months_3, months_3 = Months_2, months_2 = Prev_Month, prev_month = Curr_Month," & _
                   " curr_month = 0"
    MsgBox "tapos na"
End Sub

Private Sub Command17_Click()
    Dim rsBEGPART        As ADODB.Recordset
    Set rsBEGPART = New ADODB.Recordset
    Set rsBEGPART = gconDMIS.Execute("SELECT * FROM BEGPART ORDER BY ID ASC")
    If Not rsBEGPART.EOF And Not rsBEGPART.BOF Then
        rsBEGPART.MoveFirst
        Do While Not rsBEGPART.EOF
            gconDMIS.Execute "UPDATE BEGPART SET ID = '" & Format(rsBEGPART!ID, "00000") & "' WHERE ID = '" & rsBEGPART!ID & "'"
            rsBEGPART.MoveNext
        Loop
    End If
End Sub

Private Sub Command18_Click()
    Dim rsDAYTRAN_BEG    As ADODB.Recordset
    Set rsDAYTRAN_BEG = New ADODB.Recordset
    Set rsDAYTRAN_BEG = gconDMIS.Execute("Select * from PMIOS_DayTran Where trantype = 'BEG' order by id asc")
    If Not rsDAYTRAN_BEG.EOF And Not rsDAYTRAN_BEG.BOF Then
        rsDAYTRAN_BEG.MoveFirst
        Do While Not rsDAYTRAN_BEG.EOF
            Me.Caption = "PROC. IN - " & rsDAYTRAN_BEG!STOCK_ORD: DoEvents
            gconDMIS.Execute ("update PMIOS_STOCKMAS SEt onhand = " & N2Str2Zero(rsDAYTRAN_BEG!tranqty) & ", lastm_oh = " & N2Str2Zero(rsDAYTRAN_BEG!tranqty) & " where STOCKNO = " & N2Str2Null(rsDAYTRAN_BEG!STOCK_ORD))
            rsDAYTRAN_BEG.MoveNext
        Loop
    End If
    MsgBox "tapos"
    Exit Sub
    Set rsDAYTRAN_BEG = New ADODB.Recordset
    Set rsDAYTRAN_BEG = gconDMIS.Execute("Select * from PMIOS_DayTran Where trantype = 'OUT' order by id asc")
    If Not rsDAYTRAN_BEG.EOF And Not rsDAYTRAN_BEG.BOF Then
        rsDAYTRAN_BEG.MoveFirst
        Do While Not rsDAYTRAN_BEG.EOF
            Me.Caption = "PROC. OUT - " & rsDAYTRAN_BEG!STOCK_ORD: DoEvents
            gconDMIS.Execute ("update PMIOS_STOCKMAS SEt ISSUANCES = ISSUANCES + " & N2Str2Zero(rsDAYTRAN_BEG!tranqty) & " where STOCKNO = " & N2Str2Null(rsDAYTRAN_BEG!STOCK_ORD))
            rsDAYTRAN_BEG.MoveNext
        Loop
    End If
    MsgBox "tapos"
End Sub

Private Sub Command19_Click()
    Dim rsPARTMAS, rsCURPartmas, rsTDAYTRAN, rsRR_HD As ADODB.Recordset
    Dim i                As Integer
    Dim vTotTranCost, vTotTranInvAmt, vTotTranQTY, vTDTranQTY As Double
    Dim vTDTranType, vTDTranno, vSupCode As String
    Dim vVatAmt, vMAC    As Double
    Dim vPMOnhand        As Integer
    Dim vSTOCKDESC       As String
    'If chkUpdateAdjustment.Value = 1 Then
    Set rsTDAYTRAN = New ADODB.Recordset
    'rsTDAYTRAN.Open "select id,ItemNo,trantype,tranno,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIOS_TdayTran where trantype <> 'ADB' and status <> 'C' AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by trandate asc,trantype desc,tranno asc,itemno asc", gconDMIS
    rsTDAYTRAN.Open "select * from PMIOS_DayTran where trantype <> 'ADB' and status <> 'C' and status <> 'N' order by TRANDATE, TRANTYPE, ID asc", gconDMIS
    'Else
    '   Set rsTDAYTRAN = New ADODB.Recordset
    '       rsTDAYTRAN.Open "select id,ItemNo,trantype,tranno,STOCK_ORD,tranqty,status,in_out,tranucost from tTDAYTRAN where trantype <> 'ADJ' and trantype <> 'ADB' and status <> 'C' AND trandate >= '" & txtFrom.Text & "' AND trandate <= '" & txtTo.Text & "' order by trantype desc,trandate asc,tranno asc,itemno asc", gconDMIS
    'End If
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        rsTDAYTRAN.MoveFirst
        Screen.MousePointer = 11
        DoEvents
        Me.Caption = "Updating Part Master File"
        gconDMIS.Execute "update PMIOS_STOCKMAS set" & _
                       " onhand = 0," & _
                       " mac = 0," & _
                       " onorder = 0," & _
                       " tissqty = 0," & _
                       " trecqty = 0," & _
                       " receipts = 0," & _
                       " issuances = 0"
        DoEvents
        Me.Caption = "Updating Transactions to Part Master File"
        DoEvents
        i = 0
        Do While Not rsTDAYTRAN.EOF
            vTDTranDate = N2Date2Null(rsTDAYTRAN!trandate)
            vTDTranType = Null2String(rsTDAYTRAN!TRANTYPE)
            vTDTranno = Null2String(rsTDAYTRAN!Tranno)
            vTDTranQTY = N2Str2IntZero(rsTDAYTRAN!tranqty)
            If N2Str2Zero(rsTDAYTRAN!TRANUCOST) > 0 Then
                vTotTranCost = rsTDAYTRAN!TRANUCOST * vTDTranQTY
            Else
                vTotTranCost = 0
            End If
            'vTotTranCost = N2Str2Zero(rsTDAYTRAN!tranucost) * vTDTranQTY
            If N2Str2Zero(rsTDAYTRAN!TRANINVAMT) > 0 Then
                vTotTranInvAmt = rsTDAYTRAN!TRANINVAMT * vTDTranQTY
            Else
                vTotTranInvAmt = 0
            End If
            'vTotTranInvAmt = N2Str2Zero(rsTDAYTRAN!traninvamt * vTDTranQTY)
            Me.Caption = "Processing: " & Null2String(rsTDAYTRAN!TRANTYPE) & " #" & Null2String(rsTDAYTRAN!Tranno)
            DoEvents
            Set rsPARTMAS = New ADODB.Recordset
            Set rsPARTMAS = gconDMIS.Execute("select STOCKNO from PMIOS_STOCKMAS where STOCKNO = " & N2Str2Null(rsTDAYTRAN!STOCK_ORD))
            If rsPARTMAS.EOF And rsPARTMAS.BOF Then
                Set rsCURPartmas = New ADODB.Recordset
                Set rsCURPartmas = gconDMIS.Execute("Select STOCKNO,STOCKDESC from PMIOS_STOCKMAS where STOCKNO = " & N2Str2Null(rsTDAYTRAN!STOCK_ORD))
                If Not rsCURPartmas.EOF And Not rsCURPartmas.BOF Then
                    vSTOCKDESC = N2Str2Null(rsCURPartmas!STOCKDESC)
                Else
                    vSTOCKDESC = "'NO DESCRIPTION'"
                End If
                gconDMIS.Execute ("Insert into PMIOS_STOCKMAS (STOCKNO,STOCKDESC,date_entered) values (" & N2Str2Null(rsTDAYTRAN!STOCK_ORD) & "," & vSTOCKDESC & "," & N2Str2Null(rsTDAYTRAN!trandate) & ")")
            End If
            Set rsPARTMAS = New ADODB.Recordset
            rsPARTMAS.Open "select id,STOCKNO,mac,tissqty,trecqty,poqty,tprqty,prqty,issuances,receipts,onhand from PMIOS_STOCKMAS where STOCKNO = " & N2Str2Null(rsTDAYTRAN!STOCK_ORD), gconDMIS
            If Not rsPARTMAS.EOF And Not rsPARTMAS.BOF Then
                If N2Str2Zero(rsPARTMAS!Mac) > 0 Then vMAC = rsPARTMAS!Mac Else vMAC = 0
                'vMAC = N2Str2Zero(rsPartmas!MAC)
                vPMOnhand = N2Str2IntZero(rsPARTMAS!ONHAND)
                If Null2String(rsTDAYTRAN!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    gconDMIS.Execute "update PMIOS_STOCKMAS set " & _
                                     "onhand = " & vPMOnhand - vTDTranQTY & ", " & _
                                     "tissqty = " & N2Str2IntZero(rsPARTMAS!TISSQTY) + vTDTranQTY & ", " & _
                                     "issuances = " & N2Str2IntZero(rsPARTMAS!issuances) + vTDTranQTY & _
                                   " where id = " & rsPARTMAS!ID
                    If vMAC = 0 Then
                        'vMAC = 131
                        'Stop
                        'MsgBox "Error Encountered on Part Number (" & Null2String(rsTDAYTRAN!STOCK_ORD) & ")" & vbCrLf & " Contact the wizweirdo immediately.", vbCritical, "Warning"
                        'Exit Sub
                        'MsgBox Null2String(rsPARTMAS!STOCKNO)
                        'vMAC = 131
                        'Stop
                        'Set rsSupplier = New ADODB.Recordset
                        'Set rsSupplier = gconDMIS.Execute("Select mac from PMIOS_TdayTran where month(trandate) = " & Month(txtFrom.Text) & " and year(trandate) = " & Year(txtFrom.Text) & " and trantype = 'RR' and STOCK_ORD = " & N2Str2Null(rsPARTMAS!STOCKNO) & " order by trandate asc")
                        'If Not rsSupplier.EOF And Not rsSupplier.BOF Then
                        '  If rsSupplier!MAC <> 0 Then
                        '      vMAC = rsSupplier!MAC
                        '      Stop
                        '   Else
                        '      Stop
                        '   End If
                        '  'Stop
                        'Else
                        '   Stop
                        'End If
                        'Stop
                        'vMAC = 2272.73
                    End If
                    gconDMIS.Execute "update PMIOS_DayTran set tranucost = " & vMAC & " where ID = " & rsTDAYTRAN!ID
                End If

                If Null2String(rsTDAYTRAN!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    If vPMOnhand <= 0 Then
                        vMAC = vTotTranCost / vTDTranQTY
                        gconDMIS.Execute "update PMIOS_STOCKMAS set " & _
                                         "mac = " & vMAC & ", " & _
                                         "Onhand = " & vPMOnhand + vTDTranQTY & ", " & _
                                         "last_recd = " & vTDTranDate & ", " & _
                                         "trecqty = " & N2Str2IntZero(rsPARTMAS!trecqty) + vTDTranQTY & ", " & _
                                         "receipts = " & N2Str2IntZero(rsPARTMAS!receipts) + vTDTranQTY & _
                                       " where id =" & rsPARTMAS!ID
                    Else
                        vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                        gconDMIS.Execute "update PMIOS_STOCKMAS set " & _
                                         "mac = " & vMAC & ", " & _
                                         "Onhand = " & vPMOnhand + vTDTranQTY & ", " & _
                                         "last_recd = " & vTDTranDate & ", " & _
                                         "trecqty = " & N2Str2IntZero(rsPARTMAS!trecqty) + vTDTranQTY & ", " & _
                                         "receipts = " & N2Str2IntZero(rsPARTMAS!receipts) + vTDTranQTY & _
                                       " where id =" & rsPARTMAS!ID
                    End If
                    gconDMIS.Execute "update PMIOS_DayTran set mac = " & vMAC & ", status = 'P' where id = " & rsTDAYTRAN!ID
                End If

            End If
            DoEvents
            i = i + 1
            'progCPB.Value = (i / rsTDAYTRAN.RecordCount) * 100
            'Me.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTDAYTRAN.MoveNext
        Loop
        'labProcessing.Caption = ""
        DoEvents
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "Error Opening TDAYTRAN File"
        Exit Sub
    End If
    MsgSpeechBox "Updating of Master File Completed..."

End Sub

Private Sub Command20_Click()
    Dim RSNEWLOCATION    As ADODB.Recordset
    Set RSNEWLOCATION = New ADODB.Recordset
    Set RSNEWLOCATION = gconDMIS.Execute("sELECT * from PMIOS_NewLocation")
    If Not RSNEWLOCATION.EOF And Not RSNEWLOCATION.BOF Then
        RSNEWLOCATION.MoveFirst
        Do While Not RSNEWLOCATION.EOF
            gconDMIS.Execute ("update PMIOS_STOCKMAS SET LOCATION = " & N2Str2Null(RSNEWLOCATION!Location) & " WHERE STOCKNO = " & N2Str2Null(RSNEWLOCATION!STOCKNO))
            RSNEWLOCATION.MoveNext
        Loop
    End If
    Exit Sub
    Dim rsBEGPART        As ADODB.Recordset
    'Set rsBEGPART = New ADODB.Recordset
    'Set rsBEGPART = gconDMIS.Execute("Select * from JULBEG")
    'If Not rsBEGPART.EOF Then
    '   rsBEGPART.MoveFirst
    '   Do While Not rsBEGPART.EOF
    '      Me.Caption = Null2String(rsBEGPART!STOCKNO)
    '      gconDMIS.Execute ("update PMIOS_STOCKMAS Set SRP = " & N2Str2Zero(rsBEGPART!SRP) & " where STOCKNO = " & Trim(N2Str2Null(rsBEGPART!STOCKNO)))
    '      rsBEGPART.MoveNext
    '   Loop
    'End If
    Set rsBEGPART = New ADODB.Recordset
    Set rsBEGPART = gconDMIS.Execute("Select * from PMIOS_TdayTran where in_out = 'O' ORDER BY TRANDATE DESC")
    If Not rsBEGPART.EOF Then
        rsBEGPART.MoveFirst
        Do While Not rsBEGPART.EOF
            Me.Caption = Null2String(rsBEGPART!STOCK_ORD)
            gconDMIS.Execute ("update PMIOS_STOCKMAS Set SRP = " & N2Str2Zero(rsBEGPART!TRANINVAMT) & " where (SRP IS NULL OR SRP=0) AND STOCKNO = " & Trim(N2Str2Null(rsBEGPART!STOCK_ORD)))
            rsBEGPART.MoveNext
        Loop
    End If
    MsgBox "tapos"
End Sub

Private Sub Command21_Click()
    Dim vTrandate        As String
    Dim vTRANTYPE        As String
    Dim vTRANNO          As String
    Dim vITEMNO          As String
    Dim vSTOCK_ORD       As String
    Dim vSTOCK_SUP       As String
    Dim vTranQty         As String
    Dim vTRANUCOST       As Double
    Dim vTRANUPRICE      As Double
    Dim VStatus          As String
    Dim vIN_OUT          As String
    Dim vTRANINVAMT      As Double

    Dim TAMA_QTY         As Double
    Dim rsPARTMAS        As ADODB.Recordset
    Set rsBEGPART = New ADODB.Recordset
    Set rsBEGPART = gconDMIS.Execute("Select * from PMIOS_STOCKMAS ORDER BY STOCKNO ASC")
    rsBEGPART.MoveFirst
    Do While Not rsBEGPART.EOF
        vTrandate = "'8/1/2006'"
        vTRANTYPE = "'ADJ'"
        vSTOCK_ORD = N2Str2Null(rsBEGPART!STOCKNO)
        vSTOCK_SUP = N2Str2Null(rsBEGPART!STOCKNO)
        vTranQty = N2Str2Zero(rsBEGPART!ONHAND)
        vTRANUCOST = N2Str2Zero(rsBEGPART!DNP)
        VStatus = "'P'"
        vTRANINVAMT = N2Str2Zero(rsBEGPART!DNP)

        Set rsPARTMAS = New ADODB.Recordset
        Set rsPARTMAS = gconDMIS.Execute("SELECT * FROM JUL_OH where N2 = " & vSTOCK_ORD)
        If Not rsPARTMAS.EOF And Not rsPARTMAS.BOF Then
            TAMA_QTY = N2Str2Zero(rsPARTMAS!n12)
            If vTranQty <> TAMA_QTY Then
                Me.Caption = "ADJUSTING STOCKNO : " & Null2String(rsPARTMAS!n2): DoEvents
                If vTranQty < 0 Then
                    vTRANNO = "'111111'"
                    vITEMNO = "'1111'"
                    vTranQty = Abs(vTranQty) + TAMA_QTY
                    vIN_OUT = "'I'"
                ElseIf vTranQty > TAMA_QTY Then
                    vTRANNO = "'000000'"
                    vITEMNO = "'0000'"
                    vTranQty = vTranQty - TAMA_QTY
                    vIN_OUT = "'O'"
                ElseIf vTranQty < TAMA_QTY Then
                    vTranQty = TAMA_QTY - vTranQty
                    vIN_OUT = "'I'"
                    vTRANNO = "'111111'"
                    vITEMNO = "'1111'"
                ElseIf vTranQty = 0 Then
                    vTranQty = TAMA_QTY
                    vIN_OUT = "'I'"
                    vTRANNO = "'111111'"
                    vITEMNO = "'1111'"
                End If
                gconDMIS.Execute ("Insert into PMIOS_DayTran (TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,STATUS,IN_OUT,TRANINVAMT,NON_HARI)" & _
                                " VALUES (" & vTrandate & "," & vTRANTYPE & "," & vTRANNO & "," & vITEMNO & "," & vSTOCK_ORD & "," & vSTOCK_SUP & "," & vTranQty & "," & vTRANUCOST & "," & VStatus & "," & vIN_OUT & "," & vTRANINVAMT & ",'N')")
            End If
        Else
            If vTranQty > 0 Then
                vTRANNO = "'000000'"
                vITEMNO = "'0000'"
                vTranQty = vTranQty
                vIN_OUT = "'O'"
                gconDMIS.Execute ("Insert into PMIOS_DayTran (TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,STATUS,IN_OUT,TRANINVAMT,NON_HARI)" & _
                                " VALUES (" & vTrandate & "," & vTRANTYPE & "," & vTRANNO & "," & vITEMNO & "," & vSTOCK_ORD & "," & vSTOCK_SUP & "," & vTranQty & "," & vTRANUCOST & "," & VStatus & "," & vIN_OUT & "," & vTRANINVAMT & ",'N')")
            End If
            If vTranQty < 0 Then
                vTRANNO = "'111111'"
                vITEMNO = "'1111'"
                vTranQty = Abs(vTranQty)
                vIN_OUT = "'I'"
                gconDMIS.Execute ("Insert into PMIOS_DayTran (TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,STATUS,IN_OUT,TRANINVAMT,NON_HARI)" & _
                                " VALUES (" & vTrandate & "," & vTRANTYPE & "," & vTRANNO & "," & vITEMNO & "," & vSTOCK_ORD & "," & vSTOCK_SUP & "," & vTranQty & "," & vTRANUCOST & "," & VStatus & "," & vIN_OUT & "," & vTRANINVAMT & ",'N')")

            End If
        End If
Ryan:   rsBEGPART.MoveNext
    Loop
    MsgBox "tapos"
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command6_Click()
    Dim rsSTKSTAT        As ADODB.Recordset
    Dim rsRANKFLE        As ADODB.Recordset
    Set rsSTKSTAT = New ADODB.Recordset
    Set rsSTKSTAT = gconDMIS.Execute("Select * from PMIOS_RankFle where date_gen = '10/30/2004'")
    If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
        rsSTKSTAT.MoveFirst
        Screen.MousePointer = 11
        'gconDMIS.Execute "delete from PMIOS_StkStat where date_gen = '10/30/2004'"
        Do While Not rsSTKSTAT.EOF
            gconDMIS.Execute "update PMIOS_StkStat Set Onhand = " & N2Str2Zero(rsSTKSTAT!ONHAND) & _
                           " Where MATCDE = " & N2Str2Null(rsSTKSTAT!MATCDE) & " And Date_Gen = '10/30/2004'  "
            Me.Caption = Null2String(rsSTKSTAT!MATCDE)
            rsSTKSTAT.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Command7_Click()

    Dim rsRo_det         As ADODB.Recordset
    Dim rsOrd_Hd         As ADODB.Recordset
    Dim rsDAYTRAN        As ADODB.Recordset
    Dim rsREPOR          As ADODB.Recordset
    Dim rsPARTMAS        As ADODB.Recordset
    Dim vDate_Rel        As String

    Set rsRo_det = New ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("Select Ro_Det.ID,Ro_Det.DetAmt,Ro_Det.Rep_Or,detcde,repor.Dte_rel from Ro_Det inner join repor on ro_det.rep_or = repor.rep_or where ro_det.livil = '2' and month(repor.dte_rel) = 7 and year(repor.dte_rel) = 2006 Order by ro_det.Rep_Or asc")
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        rsRo_det.MoveFirst
        MsgBox "Ready"
        Do While Not rsRo_det.EOF
            'Set rsRepor = New ADODB.Recordset
            'Set rsRepor = gconDMIS.Execute("Select Dte_rel,Rep_or from Repor Where rep_or = " & N2Str2Null(rsRo_det!rep_or))
            'If Not rsRepor.EOF And Not rsRepor.BOF Then
            '   vDate_Rel = N2Date2Null(rsRepor!dte_rel)
            'Else
            '   vDate_Rel = "NULL"
            'End If
            Set rsOrd_Hd = New ADODB.Recordset
            Set rsOrd_Hd = gconDMIS.Execute("Select tranno from PMIOS_Ord_Hist where trantype='RIV' and RONO = '" & Left(rsRo_det!REP_OR, 1) & Right(rsRo_det!REP_OR, 6) & "'")
            If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                rsOrd_Hd.MoveFirst
                Do While Not rsOrd_Hd.EOF
                    Set rsDAYTRAN = New ADODB.Recordset
                    Set rsDAYTRAN = gconDMIS.Execute("Select tranucost from PMIOS_DayTran where trantype = 'RIV' and tranno = '" & rsOrd_Hd!Tranno & "' and STOCK_ORD = " & N2Str2Null(rsRo_det!detcde) & " order by trandate desc")
                    If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
                        If N2Str2Zero(rsDAYTRAN!TRANUCOST) > 0 Then
                            gconDMIS.Execute "Update Ro_Det Set " & _
                                             "DetCost = " & rsDAYTRAN!TRANUCOST & _
                                           " Where id = " & rsRo_det!ID
                            Me.Caption = Null2String(rsRo_det!REP_OR) & " with DetAmt: " & N2Str2Zero(rsRo_det!detamt) & " Cost = " & N2Str2Zero(rsDAYTRAN!TRANUCOST)
                        End If
                        'Else
                        '  Stop
                    End If
                    'Set rsPartmas = New ADODB.Recordset
                    'Set rsPartmas = gconDMIS.Execute("Select MAC from new_partmas where STOCKNO = " & N2Str2Null(rsRo_det!detcde))
                    'If Not rsPartmas.EOF And Not rsPartmas.BOF Then
                    '   Me.Caption = Null2String(rsRo_det!rep_or) & " with DetAmt: " & N2Str2Zero(rsRo_det!detamt) & " Cost = " & N2Str2Zero(rsDAYTRAN!tranucost) & " Pcost = " & N2Str2Zero(rsPartmas!MAC)
                    'Else
                    'End If
                    rsOrd_Hd.MoveNext
                Loop
            End If
            DoEvents
            'Stop
            rsRo_det.MoveNext
        Loop
        MsgBox "Tapos"
    End If
End Sub

Private Sub Command8_Click()
    frmPMIOSReconcileMaster.Show
End Sub

Private Sub Command9_Click()
    frmPMIOSRECONCheckPrevBal.Show
End Sub

Private Sub Form_Load()
    ComboBox1.Clear
    ComboBox1.AddItem "kim" & Chr(9) & "piya" & Chr(9) & "lim"
End Sub

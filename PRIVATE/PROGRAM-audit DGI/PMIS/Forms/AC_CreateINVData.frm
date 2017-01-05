VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPMISAC_CreateINVDATA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Cut-Off Master File"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4365
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AC_CreateINVData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "AC_CreateINVData.frx":1472
   ScaleHeight     =   780
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreateData 
      Caption         =   "Create Cut-Off Database Now"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   750
      TabIndex        =   0
      ToolTipText     =   "Create Cut-Off Database Now"
      Top             =   75
      Width           =   3315
   End
   Begin MSComDlg.CommonDialog cmdDialogINV 
      Left            =   150
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPMISAC_CreateINVDATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wsINVENTORY                                                       As Workspace
Dim dbINVENTORY                                                       As DATABASE
Dim FILNAME                                                           As String

Sub Create_Database()
    On Error GoTo ErrCode
    Set wsINVENTORY = DBEngine.Workspaces(0)
    If Exists(FILNAME) Then
        If MsgQuestionBox(FILNAME & " Already Exist! Overwrite?", "File Exist") = False Then
            Exit Sub
        Else
            Kill FILNAME
        End If
    End If
    Set dbINVENTORY = wsINVENTORY.CreateDatabase(FILNAME, dbLangGeneral)
    Set dbINVENTORY = wsINVENTORY.OpenDatabase(FILNAME)
    frmSplash.Show
    frmSplash.labCon.Caption = "Creating CutOff Accessories Master File... Please wait..."
    DoEvents
    Create_Part_MasterFile_Table
    frmSplash.labCon.Caption = "Creating Adjustment Master File... Please wait..."
    DoEvents
    Create_Adjustment_MasterFile_Table
    frmSplash.labCon.Caption = "Creating Consolidated Physical Count Master File... Please wait..."
    DoEvents
    Create_ConPhy_MasterFile_Table
    frmSplash.labCon.Caption = "Creating CutOff Accessories Master File... Please wait..."
    DoEvents
    Create_PhyCnt_MasterFile_Table
    frmSplash.labCon.Caption = "Creating Physical Count Master File... Please wait..."
    DoEvents
    Create_TagNumber_MasterFile_Table
    frmSplash.labCon.Caption = "Creating Tag Number Master File... Please wait..."
    DoEvents
    Create_Ledger_MasterFile_Table
    frmSplash.labCon.Caption = "Creating Ledger Master File... Please wait..."
    DoEvents
    dbINVENTORY.Close
    Unload frmSplash
    MsgSpeechBox "INVENTORY DATABASE AND TABLES Successfully Created!!"
    Unload Me
    Exit Sub
ErrCode:
    If Err.Number = 3204 Then
        Resume Next
    Else
        ShowVBError
    End If
End Sub

Sub Create_Part_MasterFile_Table()
    Dim i                                                             As Integer
    Dim tdCUTOFF                                                      As TableDef
    Dim FLDCUTOFF(60)                                                 As Field
    Dim CUTOFFIDIndex                                                 As Index
    Dim CUTOFFIDFLD                                                   As Field

    Set tdCUTOFF = dbINVENTORY.CreateTableDef("CUTOFF")
    Set FLDCUTOFF(0) = tdCUTOFF.CreateField("ID", dbLong)
    FLDCUTOFF(0).Required = True
    Set FLDCUTOFF(1) = tdCUTOFF.CreateField("PARTNO", dbText, 30)
    FLDCUTOFF(1).Required = True
    FLDCUTOFF(1).AllowZeroLength = False
    Set FLDCUTOFF(2) = tdCUTOFF.CreateField("PARTDESC", dbText, 150)
    Set FLDCUTOFF(3) = tdCUTOFF.CreateField("PRICECLASS", dbText, 1)
    Set FLDCUTOFF(4) = tdCUTOFF.CreateField("INVCLASS", dbText, 1)
    Set FLDCUTOFF(5) = tdCUTOFF.CreateField("VEHTYPE", dbText, 1)
    Set FLDCUTOFF(6) = tdCUTOFF.CreateField("MODELCODE", dbText, 100)
    Set FLDCUTOFF(7) = tdCUTOFF.CreateField("LOCATION", dbText, 50)
    Set FLDCUTOFF(8) = tdCUTOFF.CreateField("MAC", dbDouble)
    Set FLDCUTOFF(9) = tdCUTOFF.CreateField("MAD", dbLong)
    Set FLDCUTOFF(10) = tdCUTOFF.CreateField("OLDNO", dbText, 12)
    Set FLDCUTOFF(11) = tdCUTOFF.CreateField("NEWNO", dbText, 12)
    Set FLDCUTOFF(12) = tdCUTOFF.CreateField("GENNO", dbText, 12)
    Set FLDCUTOFF(13) = tdCUTOFF.CreateField("WFP", dbDouble)
    Set FLDCUTOFF(14) = tdCUTOFF.CreateField("SRP", dbDouble)
    Set FLDCUTOFF(15) = tdCUTOFF.CreateField("PROMO_WFP", dbDouble)
    Set FLDCUTOFF(16) = tdCUTOFF.CreateField("PROMO_SRP", dbDouble)
    Set FLDCUTOFF(17) = tdCUTOFF.CreateField("TSRP", dbDouble)
    Set FLDCUTOFF(18) = tdCUTOFF.CreateField("NOSHIP", dbLong)
    Set FLDCUTOFF(19) = tdCUTOFF.CreateField("KEYMARK", dbText, 1)
    Set FLDCUTOFF(20) = tdCUTOFF.CreateField("LASTM_MAC", dbDouble)
    Set FLDCUTOFF(21) = tdCUTOFF.CreateField("LASTM_MAD", dbDouble)
    Set FLDCUTOFF(22) = tdCUTOFF.CreateField("LASTM_SELL", dbDouble)
    Set FLDCUTOFF(23) = tdCUTOFF.CreateField("LASTM_OH", dbLong)
    Set FLDCUTOFF(24) = tdCUTOFF.CreateField("LASTM_OO", dbLong)
    Set FLDCUTOFF(25) = tdCUTOFF.CreateField("ONHAND", dbLong)
    Set FLDCUTOFF(26) = tdCUTOFF.CreateField("TRECQTY", dbLong)
    Set FLDCUTOFF(27) = tdCUTOFF.CreateField("TISSQTY", dbLong)
    Set FLDCUTOFF(28) = tdCUTOFF.CreateField("ONORDER", dbLong)
    Set FLDCUTOFF(29) = tdCUTOFF.CreateField("TPOQTY", dbLong)
    Set FLDCUTOFF(30) = tdCUTOFF.CreateField("PRQTY", dbLong)
    Set FLDCUTOFF(31) = tdCUTOFF.CreateField("TPRQTY", dbLong)
    Set FLDCUTOFF(32) = tdCUTOFF.CreateField("QTYSERVICE", dbLong)
    Set FLDCUTOFF(33) = tdCUTOFF.CreateField("UNDERREC", dbLong)
    Set FLDCUTOFF(34) = tdCUTOFF.CreateField("LAST_RECQ", dbDate)
    Set FLDCUTOFF(35) = tdCUTOFF.CreateField("LAST_RECD", dbDate)
    Set FLDCUTOFF(36) = tdCUTOFF.CreateField("LASTY_OH", dbLong)
    Set FLDCUTOFF(37) = tdCUTOFF.CreateField("LASTY_MAC", dbDouble)
    Set FLDCUTOFF(38) = tdCUTOFF.CreateField("LASTY_OO", dbLong)
    Set FLDCUTOFF(39) = tdCUTOFF.CreateField("LASTY_ADJ", dbLong)
    Set FLDCUTOFF(40) = tdCUTOFF.CreateField("HOLD", dbBoolean)
    Set FLDCUTOFF(41) = tdCUTOFF.CreateField("SUPCODE", dbText, 6)
    Set FLDCUTOFF(42) = tdCUTOFF.CreateField("VARIANCE", dbLong)
    Set FLDCUTOFF(43) = tdCUTOFF.CreateField("SUBINVCLAS", dbText, 1)
    Set FLDCUTOFF(44) = tdCUTOFF.CreateField("PHYCOUNT", dbLong)
    Set FLDCUTOFF(45) = tdCUTOFF.CreateField("ADJPHYCNT", dbLong)
    Set FLDCUTOFF(46) = tdCUTOFF.CreateField("CUTOFFQTY", dbLong)
    Set FLDCUTOFF(47) = tdCUTOFF.CreateField("CUTOFFMAC", dbDouble)
    Set FLDCUTOFF(48) = tdCUTOFF.CreateField("MAC2", dbDouble)
    Set FLDCUTOFF(49) = tdCUTOFF.CreateField("TAGNO", dbText, 10)
    Set FLDCUTOFF(50) = tdCUTOFF.CreateField("MATCH", dbText, 1)
    Set FLDCUTOFF(51) = tdCUTOFF.CreateField("RECEIPTS", dbLong)
    Set FLDCUTOFF(52) = tdCUTOFF.CreateField("ISSUANCES", dbLong)
    Set FLDCUTOFF(53) = tdCUTOFF.CreateField("TYPE", dbText, 1)
    Set FLDCUTOFF(54) = tdCUTOFF.CreateField("CHECK", dbText, 1)
    Set FLDCUTOFF(55) = tdCUTOFF.CreateField("USERCODE", dbText, 3)
    Set FLDCUTOFF(56) = tdCUTOFF.CreateField("LASTUPDATE", dbDate)
    Set FLDCUTOFF(57) = tdCUTOFF.CreateField("DNP", dbDouble)
    Set FLDCUTOFF(58) = tdCUTOFF.CreateField("VALID_ICC", dbText, 1)
    Set FLDCUTOFF(59) = tdCUTOFF.CreateField("DATE_ENTERED", dbDate)
    For i = 0 To 59
        tdCUTOFF.Fields.Append FLDCUTOFF(i)
    Next i
    Set CUTOFFIDIndex = tdCUTOFF.CreateIndex("ID")
    CUTOFFIDIndex.Primary = True
    CUTOFFIDIndex.Unique = True
    Set CUTOFFIDFLD = CUTOFFIDIndex.CreateField("ID")
    CUTOFFIDIndex.Fields.Append CUTOFFIDFLD
    tdCUTOFF.Indexes.Append CUTOFFIDIndex
    dbINVENTORY.TableDefs.Append tdCUTOFF
    Set CUTOFFIDIndex = Nothing
    Set CUTOFFIDFLD = Nothing
    Set tdCUTOFF = Nothing
    For i = 0 To 60
        Set FLDCUTOFF(i) = Nothing
    Next i
End Sub

Sub Create_Adjustment_MasterFile_Table()
    Dim i                                                             As Integer
    Dim tdADJUST                                                      As TableDef
    Dim FLDADJUST(9)                                                  As Field
    Dim ADJUSTIDIndex                                                 As Index
    Dim ADJUSTIDFLD                                                   As Field

    Set tdADJUST = dbINVENTORY.CreateTableDef("ADJUST")
    Set FLDADJUST(0) = tdADJUST.CreateField("ID", dbLong)
    Set FLDADJUST(1) = tdADJUST.CreateField("PARTNO", dbText, 30)
    FLDADJUST(1).Required = True
    FLDADJUST(1).AllowZeroLength = False
    Set FLDADJUST(2) = tdADJUST.CreateField("PARTDESC", dbText, 150)
    Set FLDADJUST(3) = tdADJUST.CreateField("MINUS", dbLong)
    Set FLDADJUST(4) = tdADJUST.CreateField("ADD", dbLong)
    Set FLDADJUST(5) = tdADJUST.CreateField("COST", dbDouble)
    Set FLDADJUST(6) = tdADJUST.CreateField("STATUS", dbText, 1)
    Set FLDADJUST(7) = tdADJUST.CreateField("USERCODE", dbText, 3)
    Set FLDADJUST(8) = tdADJUST.CreateField("LASTUPDATE", dbDate)
    For i = 0 To 8
        tdADJUST.Fields.Append FLDADJUST(i)
    Next i
    Set ADJUSTIDIndex = tdADJUST.CreateIndex("ID")
    ADJUSTIDIndex.Primary = True
    ADJUSTIDIndex.Unique = True
    Set ADJUSTIDFLD = ADJUSTIDIndex.CreateField("ID")
    ADJUSTIDIndex.Fields.Append ADJUSTIDFLD
    tdADJUST.Indexes.Append ADJUSTIDIndex
    dbINVENTORY.TableDefs.Append tdADJUST
    Set ADJUSTIDIndex = Nothing
    Set ADJUSTIDFLD = Nothing
    Set tdADJUST = Nothing
    For i = 0 To 9
        Set FLDADJUST(i) = Nothing
    Next i
End Sub

Sub Create_ConPhy_MasterFile_Table()
    Dim i                                                             As Integer
    Dim tdCONPHY                                                      As TableDef
    Dim FLDCONPHY(20)                                                 As Field
    Dim CONPHYIDIndex                                                 As Index
    Dim CONPHYIDFLD                                                   As Field

    Set tdCONPHY = dbINVENTORY.CreateTableDef("CONPHY")
    Set FLDCONPHY(0) = tdCONPHY.CreateField("ID", dbLong)
    FLDCONPHY(0).Required = True
    Set FLDCONPHY(1) = tdCONPHY.CreateField("PARTNO", dbText, 30)
    FLDCONPHY(1).Required = True
    FLDCONPHY(1).AllowZeroLength = False
    Set FLDCONPHY(2) = tdCONPHY.CreateField("PARTDESC", dbText, 150)
    Set FLDCONPHY(3) = tdCONPHY.CreateField("LOCATION", dbText, 50)
    Set FLDCONPHY(4) = tdCONPHY.CreateField("ONHAND", dbLong)
    Set FLDCONPHY(5) = tdCONPHY.CreateField("QCOUNT", dbLong)
    Set FLDCONPHY(6) = tdCONPHY.CreateField("VARIANCE", dbLong)
    Set FLDCONPHY(7) = tdCONPHY.CreateField("AMARK", dbText, 1)
    Set FLDCONPHY(8) = tdCONPHY.CreateField("ADATE", dbDate)
    Set FLDCONPHY(9) = tdCONPHY.CreateField("TAGNO", dbText, 10)
    Set FLDCONPHY(10) = tdCONPHY.CreateField("DATE_ISS", dbDate)
    Set FLDCONPHY(11) = tdCONPHY.CreateField("MAC", dbDouble)
    Set FLDCONPHY(12) = tdCONPHY.CreateField("STATUS", dbText, 2)
    Set FLDCONPHY(13) = tdCONPHY.CreateField("LASTUPDATE", dbDate)
    Set FLDCONPHY(14) = tdCONPHY.CreateField("TIME", dbText, 12)
    Set FLDCONPHY(15) = tdCONPHY.CreateField("GROUP_NO", dbText, 2)
    Set FLDCONPHY(16) = tdCONPHY.CreateField("PRINT_STAT", dbText, 1)
    Set FLDCONPHY(17) = tdCONPHY.CreateField("USERCODE", dbText, 2)
    Set FLDCONPHY(18) = tdCONPHY.CreateField("TOTALMAC", dbDouble)
    Set FLDCONPHY(19) = tdCONPHY.CreateField("NEWPARTNO", dbText, 13)
    For i = 0 To 19
        tdCONPHY.Fields.Append FLDCONPHY(i)
    Next i
    Set CONPHYIDIndex = tdCONPHY.CreateIndex("ID")
    CONPHYIDIndex.Primary = True
    CONPHYIDIndex.Unique = True
    Set CONPHYIDFLD = CONPHYIDIndex.CreateField("ID")
    CONPHYIDIndex.Fields.Append CONPHYIDFLD
    tdCONPHY.Indexes.Append CONPHYIDIndex
    dbINVENTORY.TableDefs.Append tdCONPHY
    Set CONPHYIDIndex = Nothing
    Set CONPHYIDFLD = Nothing
    Set tdCONPHY = Nothing
    For i = 0 To 20
        Set FLDCONPHY(i) = Nothing
    Next i
End Sub

Sub Create_PhyCnt_MasterFile_Table()
    Dim i                                                             As Integer
    Dim tdPHYCNT                                                      As TableDef
    Dim FLDPHYCNT(20)                                                 As Field
    Dim PHYCNTIDIndex                                                 As Index
    Dim PHYCNTIDFLD                                                   As Field

    Set tdPHYCNT = dbINVENTORY.CreateTableDef("PHYCNT")
    Set FLDPHYCNT(0) = tdPHYCNT.CreateField("ID", dbLong)
    FLDPHYCNT(0).Required = True
    Set FLDPHYCNT(1) = tdPHYCNT.CreateField("PARTNO", dbText, 30)
    FLDPHYCNT(1).Required = True
    FLDPHYCNT(1).AllowZeroLength = False
    Set FLDPHYCNT(2) = tdPHYCNT.CreateField("PARTDESC", dbText, 150)
    Set FLDPHYCNT(3) = tdPHYCNT.CreateField("LOCATION", dbText, 50)
    Set FLDPHYCNT(4) = tdPHYCNT.CreateField("ONHAND", dbLong)
    Set FLDPHYCNT(5) = tdPHYCNT.CreateField("QCOUNT", dbLong)
    Set FLDPHYCNT(6) = tdPHYCNT.CreateField("VARIANCE", dbLong)
    Set FLDPHYCNT(7) = tdPHYCNT.CreateField("AMARK", dbText, 1)
    Set FLDPHYCNT(8) = tdPHYCNT.CreateField("ADATE", dbDate)
    Set FLDPHYCNT(9) = tdPHYCNT.CreateField("TAGNO", dbText, 10)
    Set FLDPHYCNT(10) = tdPHYCNT.CreateField("DATE_ISS", dbDate)
    Set FLDPHYCNT(11) = tdPHYCNT.CreateField("MAC", dbDouble)
    Set FLDPHYCNT(12) = tdPHYCNT.CreateField("STATUS", dbText, 2)
    Set FLDPHYCNT(13) = tdPHYCNT.CreateField("LASTUPDATE", dbDate)
    Set FLDPHYCNT(14) = tdPHYCNT.CreateField("TIME", dbText, 12)
    Set FLDPHYCNT(15) = tdPHYCNT.CreateField("GROUP_NO", dbText, 2)
    Set FLDPHYCNT(16) = tdPHYCNT.CreateField("PRINT_STAT", dbText, 1)
    Set FLDPHYCNT(17) = tdPHYCNT.CreateField("USERCODE", dbText, 2)
    Set FLDPHYCNT(18) = tdPHYCNT.CreateField("TOTALMAC", dbDouble)
    Set FLDPHYCNT(19) = tdPHYCNT.CreateField("NEWPARTNO", dbText, 12)
    For i = 0 To 19
        tdPHYCNT.Fields.Append FLDPHYCNT(i)
    Next i
    Set PHYCNTIDIndex = tdPHYCNT.CreateIndex("ID")
    PHYCNTIDIndex.Primary = True
    PHYCNTIDIndex.Unique = True
    Set PHYCNTIDFLD = PHYCNTIDIndex.CreateField("ID")
    PHYCNTIDIndex.Fields.Append PHYCNTIDFLD
    tdPHYCNT.Indexes.Append PHYCNTIDIndex
    dbINVENTORY.TableDefs.Append tdPHYCNT
    Set PHYCNTIDIndex = Nothing
    Set PHYCNTIDFLD = Nothing
    Set tdPHYCNT = Nothing
    For i = 0 To 20
        Set FLDPHYCNT(i) = Nothing
    Next i
End Sub

Sub Create_TagNumber_MasterFile_Table()
    Dim i                                                             As Integer
    Dim tdTAGS                                                        As TableDef
    Dim FLDTAGS(6)                                                    As Field
    Dim TAGSIDIndex                                                   As Index
    Dim TAGSIDFLD                                                     As Field

    Set tdTAGS = dbINVENTORY.CreateTableDef("TAGS")
    Set FLDTAGS(0) = tdTAGS.CreateField("ID", dbLong)
    Set FLDTAGS(1) = tdTAGS.CreateField("TAG", dbText, 10)
    FLDTAGS(1).Required = True
    FLDTAGS(1).AllowZeroLength = False
    Set FLDTAGS(2) = tdTAGS.CreateField("PARTNO", dbText, 30)
    Set FLDTAGS(3) = tdTAGS.CreateField("STATUS", dbText, 1)
    Set FLDTAGS(4) = tdTAGS.CreateField("REMARKS", dbText, 50)
    Set FLDTAGS(5) = tdTAGS.CreateField("DUPLICATE", dbText, 8)
    For i = 0 To 5
        tdTAGS.Fields.Append FLDTAGS(i)
    Next i
    Set TAGSIDIndex = tdTAGS.CreateIndex("ID")
    TAGSIDIndex.Primary = True
    TAGSIDIndex.Unique = True
    Set TAGSIDFLD = TAGSIDIndex.CreateField("ID")
    TAGSIDIndex.Fields.Append TAGSIDFLD
    tdTAGS.Indexes.Append TAGSIDIndex
    dbINVENTORY.TableDefs.Append tdTAGS
    Set TAGSIDIndex = Nothing
    Set TAGSIDFLD = Nothing
    Set tdTAGS = Nothing
    For i = 0 To 6
        Set FLDTAGS(i) = Nothing
    Next i
End Sub

Sub Create_Ledger_MasterFile_Table()
    Dim i                                                             As Integer
    Dim tdLEDGER                                                      As TableDef
    Dim FLDLEDGER(14)                                                 As Field
    Dim LEDGERIDIndex                                                 As Index
    Dim LEDGERIDFLD                                                   As Field

    Set tdLEDGER = dbINVENTORY.CreateTableDef("LEDGER")
    Set FLDLEDGER(0) = tdLEDGER.CreateField("ID", dbLong)
    FLDLEDGER(0).Required = True
    Set FLDLEDGER(1) = tdLEDGER.CreateField("PARTNO", dbText, 30)
    FLDLEDGER(1).Required = True
    FLDLEDGER(1).AllowZeroLength = False
    Set FLDLEDGER(2) = tdLEDGER.CreateField("PARTDESC", dbText, 150)
    Set FLDLEDGER(3) = tdLEDGER.CreateField("TRANDATE", dbDate)
    Set FLDLEDGER(4) = tdLEDGER.CreateField("TRANNO", dbText, 14)
    Set FLDLEDGER(5) = tdLEDGER.CreateField("RONUMBER", dbText, 8)
    Set FLDLEDGER(6) = tdLEDGER.CreateField("WHO", dbText, 40)
    Set FLDLEDGER(7) = tdLEDGER.CreateField("RECEIVED", dbLong)
    Set FLDLEDGER(8) = tdLEDGER.CreateField("ISSUED", dbLong)
    Set FLDLEDGER(9) = tdLEDGER.CreateField("BALANCE", dbLong)
    Set FLDLEDGER(10) = tdLEDGER.CreateField("UCOST", dbDouble)
    Set FLDLEDGER(11) = tdLEDGER.CreateField("MAC", dbDouble)
    Set FLDLEDGER(12) = tdLEDGER.CreateField("TTLCOST", dbDouble)
    Set FLDLEDGER(13) = tdLEDGER.CreateField("STATUS", dbText, 10)
    For i = 0 To 13
        tdLEDGER.Fields.Append FLDLEDGER(i)
    Next i
    Set LEDGERIDIndex = tdLEDGER.CreateIndex("ID")
    LEDGERIDIndex.Primary = True
    LEDGERIDIndex.Unique = True
    Set LEDGERIDFLD = LEDGERIDIndex.CreateField("ID")
    LEDGERIDIndex.Fields.Append LEDGERIDFLD
    tdLEDGER.Indexes.Append LEDGERIDIndex
    dbINVENTORY.TableDefs.Append tdLEDGER
    Set tdLEDGER = Nothing
    For i = 0 To 14
        Set FLDLEDGER(i) = Nothing
    Next i
End Sub

Private Sub cmdCreateData_Click()

    On Error Resume Next
    Dim MYPATH, PAYLNAME                                              As String
    MYPATH = App.Path
    cmdDialogINV.Filter = "Access Files (*.MDB)|*.MDB"
    cmdDialogINV.FilterIndex = 1
    cmdDialogINV.DefaultExt = "MDB"
    PAYLNAME = cmdDialogINV.FileName
    If MYPATH <> "\" Then
        cmdDialogINV.FileName = MYPATH & "\" & cmdDialogINV.FileName
    End If
    If PAYLNAME = "" Then
        cmdDialogINV.FileName = "*.MDB"
    End If
    cmdDialogINV.Action = 2
    If Err = 32755 Then Exit Sub
    FILNAME = cmdDialogINV.FileName
    Create_Database
    LogAudit "G", "CREATE ACCESSORIES INVENTORY DATABASE"
    If Err = 32755 Then Exit Sub
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe Me, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
End Sub


Attribute VB_Name = "ImportDetailsToRO"
Option Explicit
Dim xRO_RIV_TRANNO(100)                                  As Long
Dim xRO_RIV_TRANNO_COUNTER                                As Integer
Dim xRO_MRIS_TRANNO(100)                                  As Long
Dim xRO_MRIS_TRANNO_COUNTER                               As Integer
Dim xROTOTAL                                             As Double
Dim xTOTJOBAMT                                           As Double
Dim xTOTJOBDISC                                          As Double
Dim xTOTJOBDISCVAL                                       As Double
Dim xTOTJOBTAX                                           As Double

Dim xTOTPARTSAMT                                         As Double
Dim xTOTPARTSDISC                                        As Double
Dim xTOTPARTSDISCVAL                                     As Double
Dim xTOTPARTSTAX                                         As Double

Dim xTOTMATAMT                                           As Double
Dim xTOTMATDISC                                          As Double
Dim xTOTMATDISCVAL                                       As Double
Dim xTOTMATTAX                                           As Double

Dim xTOTACCAMT                                           As Double
Dim xTOTACCDISC                                          As Double
Dim xTOTACCDISCVAL                                       As Double
Dim xTOTACCTAX                                           As Double

Dim xJOBTOTAL                                           As Double
Dim xJOBCOMTOTAL                                        As Double
Dim xJOBSALESTOTAL                                      As Double
Dim xJOBWARTOTAL                                        As Double
Dim xJOBDISCTOTAL                                       As Double
Dim xJOBVATTOTAL                                        As Double

Dim xPARTSTOTAL                                         As Double
Dim xPARTSCOMTOTAL                                      As Double
Dim xPARTSSALESTOTAL                                    As Double
Dim xPARTSWARTOTAL                                      As Double
Dim xPARTSDISCTOTAL                                     As Double
Dim xPARTSVATTOTAL                                      As Double

Dim xMATTOTAL                                           As Double
Dim xMATCOMTOTAL                                        As Double
Dim xMATSALESTOTAL                                      As Double
Dim xMATWARTOTAL                                        As Double
Dim xMATDISCTOTAL                                       As Double
Dim xMATVATTOTAL                                        As Double

Dim xACCTOTAL                                           As Double
Dim xACCCOMTOTAL                                        As Double
Dim xACCSALESTOTAL                                      As Double
Dim xACCWARTOTAL                                        As Double
Dim xACCDISCTOTAL                                       As Double
Dim xACCVATTOTAL                                        As Double
Dim VPIS_NO_CHARGE_TO                                   As String

Dim xPcnt                                               As Integer
Dim xMCNT                                               As Integer
Dim xACNT                                               As Integer
Dim xKCNT                                               As Integer

Function ImportDetails(vREP_OR As String, XTYPE As String, xLIVIL As String, Optional xTRANTYPE As String) As Boolean
    Dim rsCSMS_REPOR                                   As New ADODB.Recordset
    Dim rsfindloc                                      As New ADODB.Recordset
    
    On Error GoTo Errorcode
    If CheckIfRoIsAlreadyInvoice(vREP_OR) = True Then
        MsgBox "Repair Order is Already Billed. Transaction will not be imported!", vbInformation, vREP_OR & " Already Billed"
        Exit Function
    End If

    Dim RONOFORMAT                                     As String
    Dim YZA                                            As Integer
    Dim TISOY                                          As String
    Dim KEIKEI                                         As String
    
    RONOFORMAT = ""
    KEIKEI = ""
    TISOY = ""
    YZA = 0

    For YZA = 1 To Len(vREP_OR)
        TISOY = Mid(vREP_OR, YZA, 1)
        KEIKEI = KEIKEI + TISOY
    Next

    RONOFORMAT = KEIKEI

    Dim VARPARTSLINE_NO                                As String
    Dim VarPartNo                                      As String
    Dim VARDESCRIPTION                                 As String
    Dim VARPARTCODE                                    As String
    Dim VARQTY                                         As Double
    Dim VARUNITCOST                                    As Double
    Dim VARUNITPRICE                                   As Double
    Dim VARPARTAMOUNT                                  As String
    Dim VARCHARGETO                                    As String
    Dim VARPARTDISCOUNT                                As String
    Dim PARTSREP_OR                                    As String
    Dim PARTSLEVEL                                     As String
    Dim PARTSLINE_NO                                   As String
    Dim PARTSDETCDE                                    As String
    Dim PARTSDETDSC                                    As String
    Dim PARTSDETUNT                                    As String
    Dim PARTSDETVOL                                    As Double
    Dim PARTSDETPRC                                    As Double
    Dim PARTSDETAMT                                    As Double
    Dim PARTSCODE                                      As String
    Dim PARTSWCODE                                     As String
    Dim PARTSTAXRATE                                   As Double
    Dim PARTSDISCRATE                                  As Double
    Dim PARTSTAXVAL                                    As Double
    Dim PARTSDISVAL                                    As Double
    Dim PARTSPOCODE                                    As String
    Dim PARTSREP_OR2                                   As String
    Dim PARTSDETAIL                                    As String
    Dim PARTSDET_AMT                                   As Double
    Dim PARTSDETCOST                                   As Double
    Dim PARTSDIS_VAL                                   As Double
    Dim PARTSDISCOUNT_2                                As Double
    Dim PARTSREMARKS                                   As String
    Dim REF_RIV_ADB                                    As String
    Dim RSRR_HDCHECK                                   As New ADODB.Recordset
    Dim RSRR_HDTDAYTRANCHECK                           As New ADODB.Recordset
    Dim VGJORBP                                        As String
    Dim MTRANNO                                        As String
    Dim RSCSMSORD_HIST                                  As New ADODB.Recordset
    Dim RSCSMSDAYTRAN                                   As New ADODB.Recordset
    Dim RSCSMSORD_HD                                    As New ADODB.Recordset
    Dim RSCSMSTDAYTRAN                                  As New ADODB.Recordset
    VPIS_NO_CHARGE_TO = "NULL"
    
    VGJORBP = "NULL"
            
    gconDMIS.Execute "delete from CSMS_RO_Det where " & _
        " livil = " & N2Str2Null(xLIVIL) & _
        " AND ISNULL(ROTYPE,'') <> 'SR' " & _
        " and rep_or = " & N2Str2Null(vREP_OR) & _
        " and detcde <> 'MISC'"
        
    
    xPcnt = 0
    xRO_RIV_TRANNO_COUNTER = 0
    Set RSCSMSORD_HIST = New ADODB.Recordset
    Set RSCSMSORD_HIST = gconDMIS.Execute("select rono,trandate,trantype,tranno,REFPISNO from PMIS_ord_hIST " & _
        " where [TYPE] = " & N2Str2Null(XTYPE) & _
        " AND status <> 'C' " & _
        " AND ISNULL(STATUS2,'') <> 'R' " & _
        " and status <> 'N' " & _
        " and trantype IN('RIV','ADB') " & _
        " and rono = '" & RONOFORMAT & "'")
    If Not RSCSMSORD_HIST.EOF And Not RSCSMSORD_HIST.BOF Then
        RSCSMSORD_HIST.MoveFirst
        Do While Not RSCSMSORD_HIST.EOF
            If Mid(Null2String(RSCSMSORD_HIST!REFPISNO), 5, 1) = "C" Then VPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(RSCSMSORD_HIST!REFPISNO), 5, 1) = "I" Then VPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(RSCSMSORD_HIST!REFPISNO), 5, 1) = "W" Then VPIS_NO_CHARGE_TO = "'W'"
            If Mid(Null2String(RSCSMSORD_HIST!REFPISNO), 4, 1) = "B" Then VGJORBP = "'BP'"
            If Mid(Null2String(RSCSMSORD_HIST!REFPISNO), 4, 1) = "G" Then VGJORBP = "'GJ'"

            Set RSCSMSDAYTRAN = New ADODB.Recordset
            Set RSCSMSDAYTRAN = gconDMIS.Execute("SELECT A.[TYPE],A.TRANDATE,A.TRANTYPE,A.TRANNO,A.ITEMNO,A.STOCK_ORD,A.STOCK_SUP, " & _
                " CAST(ISNULL(A.TRANQTY,0) - ISNULL(B.QTY_REQ,0) AS INT) AS TRANQTY,TRANUCOST,TRANUPRICE, " & _
                " NETCOST , NETPRICE, A.STATUS, A.IN_OUT, A.Mac, A.USERCODE, A.LASTUPDATE, A.ID " & _
                " FROM PMIS_DAYTRAN A LEFT OUTER JOIN " & _
                " ( " & _
                " SELECT A.REP_OR,A.STATUS,B.STOCKNO,B.ITEMID,B.STOCK_TYPE,B.QTY_REQ FROM CSMS_RETURN_HD A INNER JOIN CSMS_RETURN_DET B " & _
                " ON A.ID = B.ID_HD WHERE A.STATUS = 'P' AND A.VERI_BY IS NOT NULL " & _
                " )B " & _
                " ON A.ID = B.ITEMID AND A.[TYPE] = B.STOCK_TYPE " & _
                " WHERE (ISNULL(A.TRANQTY,0) - ISNULL(B.QTY_REQ,0)) > 0 " & _
                " AND A.TRANTYPE =  " & N2Str2Null(RSCSMSORD_HIST!TranType) & _
                " AND A.[TYPE] = " & N2Str2Null(XTYPE) & _
                " AND A.STATUS = 'P' " & _
                " AND A.TRANNO = " & N2Str2Null(RSCSMSORD_HIST!TRANNO) & _
                " ORDER BY ITEMNO ASC")
                
            If Not RSCSMSDAYTRAN.EOF And Not RSCSMSDAYTRAN.BOF Then
                Screen.MousePointer = 11
                RSCSMSDAYTRAN.MoveFirst
                xRO_RIV_TRANNO_COUNTER = xRO_RIV_TRANNO_COUNTER + 1
                xRO_RIV_TRANNO(xRO_RIV_TRANNO_COUNTER) = Null2String(RSCSMSORD_HIST!TRANNO)
                Do While Not RSCSMSDAYTRAN.EOF
                    MTRANNO = Null2String(RSCSMSDAYTRAN!TRANNO)
                    xPcnt = xPcnt + 1
                    VARPARTSLINE_NO = ""
                    VarPartNo = ""
                    VARDESCRIPTION = ""
                    VARPARTCODE = ""
                    VARQTY = 0
                    VARUNITPRICE = 0
                    VARPARTAMOUNT = ""
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO

                    VARPARTSLINE_NO = Format(xPcnt, "00")
                    VarPartNo = Null2String(RSCSMSDAYTRAN!STOCK_ORD)
                    VARDESCRIPTION = SetSTOCKDESC(Null2String(RSCSMSDAYTRAN!STOCK_ORD))
                    VARPARTCODE = "01"
                    VARQTY = Format(N2Str2IntZero(RSCSMSDAYTRAN!TRANQTY), "####0")
                    VARUNITCOST = N2Str2Zero(RSCSMSDAYTRAN!TRANUCOST)
                    VARUNITPRICE = N2Str2Zero(RSCSMSDAYTRAN!TRANUPRICE)
                    VARPARTAMOUNT = N2Str2Zero(RSCSMSDAYTRAN!TRANQTY) * N2Str2Zero(RSCSMSDAYTRAN!TRANUPRICE)
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO
                    REF_RIV_ADB = "'RIV" & Format(Null2String(RSCSMSDAYTRAN!TRANNO), "000000") & Format(Null2String(RSCSMSDAYTRAN!itemno), "000") & "'"
                    PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                    PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0
                    PARTSREP_OR = N2Str2Null(vREP_OR)
                    PARTSLEVEL = N2Str2Null(xLIVIL)
                    PARTSLINE_NO = N2Str2Null(Format(VARPARTSLINE_NO, "00"))
                    PARTSDETCDE = N2Str2Null(Left(VarPartNo, 50))
                    PARTSDETDSC = N2Str2Null(Mid(VARDESCRIPTION, 1, 100))
                    PARTSDETUNT = "NULL"
                    PARTSDETVOL = N2Str2Zero(VARQTY)
                    PARTSDETCOST = NumericVal(VARUNITCOST)
                    PARTSDETPRC = NumericVal(VARUNITPRICE)
                    PARTSDETAMT = Round(NumericVal(VARPARTAMOUNT) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    PARTSCODE = "NULL"
                    PARTSWCODE = VPIS_NO_CHARGE_TO
                    PARTSTAXRATE = (VAT_RATE / 100)
                    PARTSDISCRATE = NumericVal(VARPARTDISCOUNT) / 100
                    PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                    PARTSPOCODE = N2Str2Null(VARPARTCODE)
                    PARTSREP_OR2 = "NULL"
                    PARTSDETAIL = "NULL"
                    PARTSDET_AMT = NumericVal(VARPARTAMOUNT)
                    PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                    PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                    
                    PARTSTAXVAL = Round((PARTSDET_AMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                    SQL_STATEMENT = "insert into CSMS_RO_Det " & _
                        "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detcost,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                        " values (" & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                        " " & VGJORBP & ", " & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                        " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                        " " & PARTSDETCOST & ", " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                        ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                        ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                        ", " & PARTSREP_OR2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                        ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    
                    '-----------------------------------------------------------------------------------------
                    'updated by:    IEBV 12032010_0310pm
                    'description:   Update LOCnumber for the ro number
                            If COMPANY_CODE = "HLU" Then
                                Call RODET_LOCNUM(PARTSDETCDE, RSCSMSORD_HD!TRANNO, PARTSREP_OR)
                            End If
                    '-----------------------------------------------------------------------------------------
                    
                    'NEW LOG AUDIT ---------------------------------------------------------
                        Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "P", "PART NO: " & Null2String(RSCSMSDAYTRAN!TRANNO) & " ,PART NO: " & Null2String(PARTSDETCDE) & " - HISTORY", "RIV", "")
                    'NEW LOG AUDIT ---------------------------------------------------------
                    Screen.MousePointer = 0
10000               RSCSMSDAYTRAN.MoveNext
                Loop
            End If
            RSCSMSORD_HIST.MoveNext
        Loop
    End If

    'PRESENT************************************************
    Set RSCSMSORD_HD = gconDMIS.Execute("select rono,tranno,trantype,REFPISNO from PMIS_ord_hd where " & _
         " [TYPE] = " & N2Str2Null(XTYPE) & _
         " AND status <> 'C' " & _
         " AND ISNULL(STATUS2,'') <> 'R' " & _
         " and status <> 'N' " & _
         " and trantype IN('RIV','ADB') " & _
         " and rono = '" & RONOFORMAT & "'")
    If Not RSCSMSORD_HD.EOF And Not RSCSMSORD_HD.BOF Then
        RSCSMSORD_HD.MoveFirst
        Do While Not RSCSMSORD_HD.EOF
            If Mid(Null2String(RSCSMSORD_HD!REFPISNO), 5, 1) = "C" Then VPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(RSCSMSORD_HD!REFPISNO), 5, 1) = "I" Then VPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(RSCSMSORD_HD!REFPISNO), 5, 1) = "W" Then VPIS_NO_CHARGE_TO = "'W'"
            If Mid(Null2String(RSCSMSORD_HD!REFPISNO), 4, 1) = "B" Then VGJORBP = "'BP'"
            If Mid(Null2String(RSCSMSORD_HD!REFPISNO), 4, 1) = "G" Then VGJORBP = "'GJ'"

            Set RSCSMSTDAYTRAN = gconDMIS.Execute("select itemno,trantype,tranno,STOCK_ord,STOCK_sup,tranqty,tranucost,tranuprice from PMIS_TdayTran where " & _
                " [TYPE] = " & N2Str2Null(XTYPE) & _
                " AND trantype = '" & Null2String(RSCSMSORD_HD!TranType) & _
                "' and tranno = " & N2Str2Null(RSCSMSORD_HD!TRANNO) & _
                " order by itemno asc")
            If Not RSCSMSTDAYTRAN.EOF And Not RSCSMSTDAYTRAN.BOF Then
                Screen.MousePointer = 11
                RSCSMSTDAYTRAN.MoveFirst
                xRO_RIV_TRANNO_COUNTER = xRO_RIV_TRANNO_COUNTER + 1
                xRO_RIV_TRANNO(xRO_RIV_TRANNO_COUNTER) = Null2String(RSCSMSORD_HD!TRANNO)

                Do While Not RSCSMSTDAYTRAN.EOF
                    MTRANNO = Null2String(RSCSMSTDAYTRAN!TRANNO)
                    xPcnt = xPcnt + 1
                    VARPARTSLINE_NO = ""
                    VarPartNo = ""
                    VARDESCRIPTION = ""
                    VARPARTCODE = ""
                    VARQTY = 0
                    VARUNITPRICE = 0
                    VARPARTAMOUNT = ""
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO
                    VARPARTSLINE_NO = Format(xPcnt, "00")
                    VarPartNo = Null2String(RSCSMSTDAYTRAN!STOCK_ORD)
                    VARDESCRIPTION = SetSTOCKDESC(Null2String(RSCSMSTDAYTRAN!STOCK_ORD))
                    VARPARTCODE = "01"
                    VARQTY = Format(N2Str2IntZero(RSCSMSTDAYTRAN!TRANQTY), "####0")
                    VARUNITCOST = N2Str2Zero(RSCSMSTDAYTRAN!TRANUCOST)
                    VARUNITPRICE = N2Str2Zero(RSCSMSTDAYTRAN!TRANUPRICE)
                    VARPARTAMOUNT = N2Str2Zero(RSCSMSTDAYTRAN!TRANQTY) * N2Str2Zero(RSCSMSTDAYTRAN!TRANUPRICE)
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO
                    REF_RIV_ADB = "'RIV" & Format(Null2String(RSCSMSTDAYTRAN!TRANNO), "000000") & Format(Null2String(RSCSMSTDAYTRAN!itemno), "000") & "'"
                    PARTSDISVAL = 0
                    PARTSTAXVAL = 0
                    PARTSDETAMT = 0
                    PARTSDIS_VAL = 0
                    PARTSDISCOUNT_2 = 0
                    PARTSDISCRATE = 0
                    PARTSREP_OR = N2Str2Null(vREP_OR)
                    PARTSLEVEL = N2Str2Null(xLIVIL)
                    
                    PARTSLINE_NO = N2Str2Null(Format(VARPARTSLINE_NO, "00"))
                    PARTSDETCDE = N2Str2Null(Left(VarPartNo, 50))
                    PARTSDETDSC = N2Str2Null(Mid(VARDESCRIPTION, 1, 100))
                    PARTSDETUNT = "NULL"
                    PARTSDETVOL = N2Str2Zero(VARQTY)
                    PARTSDETCOST = NumericVal(VARUNITCOST)
                    PARTSDETPRC = NumericVal(VARUNITPRICE)
                    PARTSDETAMT = Round(NumericVal(VARPARTAMOUNT) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    PARTSCODE = "NULL"
                    PARTSWCODE = VPIS_NO_CHARGE_TO    'N2Str2Null(VarChargeTo)
                    PARTSTAXRATE = (VAT_RATE / 100)
                    PARTSDISCRATE = Round(NumericVal(VARPARTDISCOUNT) / 100, 2)
                    PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                    PARTSPOCODE = N2Str2Null(VARPARTCODE)
                    PARTSREP_OR2 = "NULL"
                    PARTSDETAIL = "NULL"
                    PARTSDET_AMT = NumericVal(VARPARTAMOUNT)
                    PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                    PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                    PARTSTAXVAL = Round((PARTSDET_AMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                    SQL_STATEMENT = "INSERT INTO CSMS_RO_DET " & _
                        "(REP_OR,LIVIL,LINE_NO,JOBTYPE,DETCDE,DETDSC,DETUNT,DETVOL,DETCOST,DETPRC,DETAMT,CODE,WCODE,TAXRATE,DISCRATE,TAXVAL,DISVAL,POCODE,REP_OR2,DETAIL,DET_AMT,DIS_VAL,DISCOUNT_2,REF_RIV_ADB)" & _
                        " VALUES (" & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                        " " & VGJORBP & ", " & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                        " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                        " " & PARTSDETCOST & ", " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                        ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                        ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                        ", " & PARTSREP_OR2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                        ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    
                    '-----------------------------------------------------------------------------------------
                    'updated by:    IEBV 12032010_0310pm
                    'description:   Update LOCnumber for the ro number
                            If COMPANY_CODE = "HLU" Then
                                Call RODET_LOCNUM(PARTSDETCDE, RSCSMSORD_HD!TRANNO, PARTSREP_OR)
                            End If
                    '-----------------------------------------------------------------------------------------
                    
                    'NEW LOG AUDIT ---------------------------------------------------------
                        Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "P", "TRAN NO: " & Null2String(RSCSMSTDAYTRAN!TRANNO) & " ,PART NO: " & Null2String(PARTSDETCDE), "RIV", "")
                    'NEW LOG AUDIT ---------------------------------------------------------
                    Screen.MousePointer = 0
20000               RSCSMSTDAYTRAN.MoveNext
                Loop
            End If
            RSCSMSORD_HD.MoveNext
        Loop
    End If

    Set RSCSMSORD_HD = New ADODB.Recordset
    Set RSCSMSORD_HD = gconDMIS.Execute("Select PARTICIPAT, INSAMT, PARTLABOR, PARTPARTS, PARTMATERIALS, PARTACCESSORIES FROM CSMS_REPOR WHERE " & _
        " REP_OR = " & N2Str2Null(vREP_OR))
    If Not RSCSMSORD_HD.EOF And Not RSCSMSORD_HD.BOF Then
        If Null2String(RSCSMSORD_HD!PARTICIPAT) = "" Then
            Call FillDetails(vREP_OR, "1")
            Call FillDetails(vREP_OR, "2")
            Call FillDetails(vREP_OR, "3")
            Call FillDetails(vREP_OR, "4")
            
            xROTOTAL = xTOTJOBAMT + xTOTPARTSAMT + xTOTMATAMT + xTOTACCAMT
        Else
            If COMPANY_CODE = "HII" Or COMPANY_CODE = "HBI" Then
                xROTOTAL = xTOTJOBAMT + xTOTPARTSAMT + xTOTMATAMT + xTOTACCAMT
            Else
                If NumericVal(RSCSMSORD_HD!INSAMT) = 0 Then
                    xROTOTAL = xTOTJOBAMT + xTOTPARTSAMT + xTOTMATAMT + xTOTACCAMT
                Else
                    xROTOTAL = xTOTJOBAMT + xTOTPARTSAMT + xTOTMATAMT + xTOTACCAMT - (NumericVal(RSCSMSORD_HD!INSAMT))
                End If
            End If
        End If
        
        If NumericVal(RSCSMSORD_HD!INSAMT) = 0 Then
            SQL_STATEMENT = "update CSMS_RepOr set" & _
                " labor = " & xTOTJOBAMT - xTOTJOBTAX - (NumericVal(RSCSMSORD_HD!PARTLABOR)) & "," & _
                " l_amtvalue = " & Round(xTOTJOBAMT, 2) - (NumericVal(RSCSMSORD_HD!PARTLABOR)) & "," & _
                " l_disc = " & Round(xTOTJOBDISCVAL, 2) & ", l_disc2 = " & Round(xTOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
                " l_taxval = " & Round(xTOTJOBTAX, 2) & ", l_discount = " & Round(xTOTJOBDISC, 2) & "," & _
                " wl_amt = " & 0 & "," & _
                " parts = " & xTOTPARTSAMT - xTOTPARTSTAX - (NumericVal(RSCSMSORD_HD!PARTPARTS)) & "," & _
                " p_amtvalue = " & xTOTPARTSAMT - NumericVal(RSCSMSORD_HD!PARTPARTS) & "," & _
                " p_disc = " & xTOTPARTSDISCVAL & ", p_disc2 = " & xTOTPARTSDISC * (VAT_RATE / 100) & "," & _
                " p_taxval = " & xTOTPARTSTAX & ", p_discount = " & xTOTPARTSDISC & "," & _
                " wp_amt = " & 0 & "," & _
                " material = " & xTOTMATAMT - xTOTMATTAX - NumericVal(RSCSMSORD_HD!PARTMATERIALS) & "," & _
                " m_amtvalue = " & xTOTMATAMT - NumericVal(RSCSMSORD_HD!PARTMATERIALS) & "," & _
                " m_disc = " & xTOTMATDISCVAL & ", m_disc2 = " & xTOTMATDISC * (VAT_RATE / 100) & "," & _
                " m_taxval = " & xTOTMATTAX & ", m_discount = " & xTOTMATDISC & "," & _
                " Accessories = " & xTOTACCAMT - xTOTACCTAX - NumericVal(RSCSMSORD_HD!PARTACCESSORIES) & "," & _
                " A_amtvalue = " & xTOTACCAMT - NumericVal(RSCSMSORD_HD!PARTACCESSORIES) & "," & _
                " A_disc = " & xTOTACCDISCVAL & ", A_disc2 = " & xTOTACCDISC * (VAT_RATE / 100) & "," & _
                " A_taxval = " & xTOTACCTAX & ", A_discount = " & xTOTACCDISC & "," & _
                " amount = " & xROTOTAL - xTOTJOBDISC - xTOTPARTSDISC - xTOTMATDISC - xTOTACCDISC & "," & _
                " rovat = " & xTOTJOBTAX + xTOTPARTSTAX + xTOTMATTAX + xTOTACCTAX & "," & _
                " WA_amt = " & 0 & "," & _
                " ro_amount = " & Round(xROTOTAL - xTOTJOBDISC - xTOTPARTSDISC - xTOTMATDISC - xTOTACCDISC, 2) & _
                " where rep_or = " & N2Str2Null(vREP_OR)
        Else
            SQL_STATEMENT = "update CSMS_RepOr set" & _
                " labor = " & xTOTJOBAMT - (NumericVal(RSCSMSORD_HD!PARTLABOR)) & "," & _
                " l_amtvalue = " & Round(xTOTJOBAMT, 2) - (NumericVal(RSCSMSORD_HD!PARTLABOR)) & "," & _
                " l_disc = " & Round(xTOTJOBDISCVAL, 2) & ", l_disc2 = " & Round(xTOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
                " l_taxval = " & Round(xTOTJOBTAX, 2) & ", l_discount = " & Round(xTOTJOBDISC, 2) & "," & _
                " wl_amt = " & 0 & "," & _
                " parts = " & xTOTPARTSAMT - (NumericVal(RSCSMSORD_HD!PARTPARTS)) & "," & _
                " p_amtvalue = " & xTOTPARTSAMT - NumericVal(RSCSMSORD_HD!PARTPARTS) & "," & _
                " p_disc = " & xTOTPARTSDISCVAL & ", p_disc2 = " & xTOTPARTSDISC * (VAT_RATE / 100) & "," & _
                " p_taxval = " & xTOTPARTSTAX & ", p_discount = " & xTOTPARTSDISC & "," & _
                " wp_amt = " & 0 & "," & _
                " material = " & xTOTMATAMT - NumericVal(RSCSMSORD_HD!PARTMATERIALS) & "," & _
                " m_amtvalue = " & xTOTMATAMT - NumericVal(RSCSMSORD_HD!PARTMATERIALS) & "," & _
                " m_disc = " & xTOTMATDISCVAL & ", m_disc2 = " & xTOTMATDISC * (VAT_RATE / 100) & "," & _
                " m_taxval = " & xTOTMATTAX & ", m_discount = " & xTOTMATDISC & "," & _
                " Accessories = " & xTOTACCAMT - NumericVal(RSCSMSORD_HD!PARTACCESSORIES) & "," & _
                " A_amtvalue = " & xTOTACCAMT - NumericVal(RSCSMSORD_HD!PARTACCESSORIES) & "," & _
                " A_disc = " & xTOTACCDISCVAL & ", A_disc2 = " & xTOTACCDISC * (VAT_RATE / 100) & "," & _
                " A_taxval = " & xTOTACCTAX & ", A_discount = " & xTOTACCDISC & "," & _
                " amount = " & xROTOTAL - xTOTJOBDISC - xTOTPARTSDISC - xTOTMATDISC - xTOTACCDISC & "," & _
                " rovat = " & xTOTJOBTAX + xTOTPARTSTAX + xTOTMATTAX + xTOTACCTAX & "," & _
                " WA_amt = " & 0 & "," & _
                " ro_amount = " & Round(xROTOTAL - xTOTJOBDISC - xTOTPARTSDISC - xTOTMATDISC - xTOTACCDISC, 2) & _
                " where rep_or = " & N2Str2Null(vREP_OR)
        End If
        gconDMIS.Execute SQL_STATEMENT
            
        'NEW LOG AUDIT ---------------------------------------------------------
            Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), N2Str2Null(XTYPE), "TRAN NO: " & MTRANNO, N2Str2Null(xTRANTYPE), "")
        'NEW LOG AUDIT ---------------------------------------------------------
    End If
    
    Screen.MousePointer = 0
    ImportDetails = True
    Exit Function

Errorcode:
    Screen.MousePointer = 0:
    ImportDetails = False
    ShowVBError
    MsgBox err.Description
End Function

Sub FillDetails(vREP_OR As String, vLIVIL As String)
    Dim rsRO_DET                                       As New ADODB.Recordset
    
    If Null2String(vLIVIL) = "1" Then
        xTOTJOBAMT = 0:                 xTOTJOBDISC = 0
        xTOTJOBDISCVAL = 0:             xTOTJOBTAX = 0
        xKCNT = 0:                      xJOBCOMTOTAL = 0
        xJOBSALESTOTAL = 0:             xJOBWARTOTAL = 0
    ElseIf Null2String(vLIVIL) = "2" Then
        xTOTPARTSAMT = 0:               xTOTPARTSDISC = 0
        xTOTPARTSDISCVAL = 0:           xTOTPARTSTAX = 0
        xPcnt = 0:                      xPARTSCOMTOTAL = 0
        xPARTSSALESTOTAL = 0:           xPARTSWARTOTAL = 0
    ElseIf Null2String(vLIVIL) = "3" Then
        xTOTMATAMT = 0:                 xTOTMATDISC = 0
        xTOTMATDISCVAL = 0:             xTOTMATTAX = 0
        xMCNT = 0: xMATCOMTOTAL = 0
        xMATSALESTOTAL = 0:             xMATWARTOTAL = 0
    Else
        xTOTACCAMT = 0:                 xTOTACCDISC = 0
        xTOTACCDISCVAL = 0:             xTOTACCTAX = 0
        xACNT = 0:                      xACCCOMTOTAL = 0
        xACCSALESTOTAL = 0:             xACCWARTOTAL = 0
    End If

    Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det " & _
        " where rep_or = " & N2Str2Null(vREP_OR) & _
        " and livil = " & N2Str2Null(vLIVIL) & _
        " order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Do While Not rsRO_DET.EOF
            If Null2String(vLIVIL) = "1" Then
                xKCNT = xKCNT + 1
                If Null2String(rsRO_DET!wCode) = "C" Then
                    xJOBCOMTOTAL = xJOBCOMTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                ElseIf Null2String(rsRO_DET!wCode) = "S" Then
                    xJOBSALESTOTAL = xJOBSALESTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                ElseIf Null2String(rsRO_DET!wCode) = "W" Then
                    xJOBWARTOTAL = xJOBWARTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                Else
                    xTOTJOBAMT = xTOTJOBAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                    xTOTJOBDISC = xTOTJOBDISC + N2Str2Zero(rsRO_DET!discount_2)
                    xTOTJOBDISCVAL = xTOTJOBDISCVAL + N2Str2Zero(rsRO_DET!disval)
                    xTOTJOBTAX = xTOTJOBTAX + N2Str2Zero(rsRO_DET!taxval)
                End If
            ElseIf Null2String(vLIVIL) = "2" Then
                xPcnt = xPcnt + 1
                If Null2String(rsRO_DET!wCode) = "C" Then
                    xPARTSCOMTOTAL = xPARTSCOMTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                ElseIf Null2String(rsRO_DET!wCode) = "S" Then
                    xPARTSSALESTOTAL = xPARTSSALESTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                ElseIf Null2String(rsRO_DET!wCode) = "W" Then
                    xPARTSWARTOTAL = xPARTSWARTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                Else
                    xTOTPARTSAMT = xTOTPARTSAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                    xTOTPARTSDISC = xTOTPARTSDISC + N2Str2Zero(rsRO_DET!discount_2)
                    xTOTPARTSDISCVAL = xTOTPARTSDISCVAL + N2Str2Zero(rsRO_DET!disval)
                    xTOTPARTSTAX = xTOTPARTSTAX + N2Str2Zero(rsRO_DET!taxval)
                End If
            ElseIf Null2String(vLIVIL) = "3" Then
                xMCNT = xMCNT + 1
                If Null2String(rsRO_DET!wCode) = "C" Then
                    xMATCOMTOTAL = xMATCOMTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                ElseIf Null2String(rsRO_DET!wCode) = "S" Then
                    xMATSALESTOTAL = xMATSALESTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                ElseIf Null2String(rsRO_DET!wCode) = "W" Then
                    xMATWARTOTAL = xMATWARTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                Else
                    xTOTMATAMT = xTOTMATAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                    xTOTMATDISC = xTOTMATDISC + N2Str2Zero(rsRO_DET!discount_2)
                    xTOTMATDISCVAL = xTOTMATDISCVAL + N2Str2Zero(rsRO_DET!disval)
                    xTOTMATTAX = xTOTMATTAX + N2Str2Zero(rsRO_DET!taxval)
                End If
            Else
                xACNT = xACNT + 1
                If Null2String(rsRO_DET!wCode) = "C" Then
                    xACCCOMTOTAL = xACCCOMTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                ElseIf Null2String(rsRO_DET!wCode) = "S" Then
                    xACCSALESTOTAL = xACCSALESTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                ElseIf Null2String(rsRO_DET!wCode) = "W" Then
                    xACCWARTOTAL = xACCWARTOTAL + N2Str2Zero(rsRO_DET!Det_AMT)
                Else
                    xTOTACCAMT = xTOTACCAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                    xTOTACCDISC = xTOTACCDISC + N2Str2Zero(rsRO_DET!discount_2)
                    xTOTACCDISCVAL = xTOTACCDISCVAL + N2Str2Zero(rsRO_DET!disval)
                    xTOTACCTAX = xTOTACCTAX + N2Str2Zero(rsRO_DET!taxval)
                End If
            End If
            rsRO_DET.MoveNext
        Loop
    End If
    Set rsRO_DET = Nothing
    
    If Null2String(vLIVIL) = "1" Then
        xTOTJOBAMT = Round(xTOTJOBAMT, 2):              xTOTJOBDISC = Round(xTOTJOBDISC, 2)
        xTOTJOBDISCVAL = Round(xTOTJOBDISCVAL, 2):      xTOTJOBTAX = Round(xTOTJOBTAX, 2)
    ElseIf Null2String(vLIVIL) = "2" Then
        xTOTPARTSAMT = Round(xTOTPARTSAMT, 2):          xTOTPARTSDISC = Round(xTOTPARTSDISC, 2)
        xTOTPARTSDISCVAL = Round(xTOTPARTSDISCVAL, 2):  xTOTPARTSTAX = Round(xTOTPARTSTAX, 2)
    ElseIf Null2String(vLIVIL) = "3" Then
        xTOTMATAMT = Round(xTOTMATAMT, 2):              xTOTMATDISC = Round(xTOTMATDISC, 2)
        xTOTMATDISCVAL = Round(xTOTMATDISCVAL, 2):      xTOTMATTAX = Round(xTOTMATTAX, 2)
    Else
        xTOTACCAMT = Round(xTOTACCAMT, 2):              xTOTACCDISC = Round(xTOTACCDISC, 2)
        xTOTACCDISCVAL = Round(xTOTACCDISCVAL, 2):      xTOTACCTAX = Round(xTOTACCTAX, 2)
    End If
End Sub

'updated by: IEBV 12032010_0310pm
'description:   to save LOC number(s) of part number in service
Function RODET_LOCNUM(partcode As String, vtranno As String, part_rep_or As String)
    Dim rsLOC_NUM As ADODB.Recordset
    Dim rsro As ADODB.Recordset
    Dim LOC_NO As String
    Dim stakno As String
    Dim NEWLOC_NO As String
    
    Set rsLOC_NUM = New ADODB.Recordset
    Set rsLOC_NUM = gconDMIS.Execute("SELECT DISTINCT LOC_NUMBER , stockno from pmis_fifo where stockno = " & partcode & " and Status = 'ISSUED' and tranno = '" & vtranno & "'")
    If Not rsLOC_NUM.EOF And Not rsLOC_NUM.BOF Then
        LOC_NO = ""
        rsLOC_NUM.MoveFirst
        Do While Not rsLOC_NUM.EOF
            LOC_NO = LOC_NO & Trim(Null2String(rsLOC_NUM!LOC_NUMBER)) & ","
            stakno = rsLOC_NUM!STOCKNO
            rsLOC_NUM.MoveNext
        Loop
    NEWLOC_NO = ""
    NEWLOC_NO = Left(LOC_NO, (Len(LOC_NO) - 1))
    gconDMIS.Execute ("Update CSMS_RO_DET set LOC_NUMBER = '" & UCase(NEWLOC_NO) & "' where rep_or =  " & part_rep_or & " and DETCDE = '" & stakno & "'")
    End If
End Function


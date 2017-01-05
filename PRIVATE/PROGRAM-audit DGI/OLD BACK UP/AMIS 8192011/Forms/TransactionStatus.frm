VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmTransactionStatus 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Status"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "TransactionStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   75
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox picLoading 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   3180
      ScaleHeight     =   1005
      ScaleWidth      =   5895
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   5925
      Begin MSComctlLib.ProgressBar PROGBAR 
         Height          =   405
         Left            =   60
         TabIndex        =   5
         Top             =   510
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblPercent 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   60
         TabIndex        =   7
         Top             =   270
         Width           =   5085
      End
      Begin XtremeShortcutBar.ShortcutCaption lblData 
         Height          =   255
         Left            =   -30
         TabIndex        =   6
         Top             =   0
         Width           =   5955
         _Version        =   655364
         _ExtentX        =   10504
         _ExtentY        =   450
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   8421504
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   1065
         Left            =   -30
         TabIndex        =   8
         Top             =   -60
         Width           =   6105
         _Version        =   655364
         _ExtentX        =   10769
         _ExtentY        =   1879
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   405
      Left            =   9630
      TabIndex        =   2
      Top             =   60
      Width           =   1605
   End
   Begin MSFlexGridLib.MSFlexGrid gridStatus 
      Height          =   6375
      Left            =   30
      TabIndex        =   3
      Top             =   570
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   0
      BackColorSel    =   16744448
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483633
      GridColor       =   0
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "TransactionStatus.frx":09AA
   End
   Begin XtremeSuiteControls.TabControl SearchTab 
      CausesValidation=   0   'False
      Height          =   615
      Left            =   30
      TabIndex        =   1
      Top             =   6930
      Width           =   11460
      _Version        =   655364
      _ExtentX        =   20214
      _ExtentY        =   1085
      _StockProps     =   64
      Appearance      =   9
      Color           =   4
      PaintManager.Layout=   1
      PaintManager.Position=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   6
      Item(0).Caption =   "Trial Balance"
      Item(0).ControlCount=   0
      Item(1).Caption =   "Vendor Code in A/R Account"
      Item(1).ControlCount=   0
      Item(2).Caption =   "Invalid Reference"
      Item(2).Tooltip =   "Search Sales Journals by Voucher Number"
      Item(2).ControlCount=   0
      Item(3).Caption =   "Customer Opening Bal vs GL"
      Item(3).Tooltip =   "Search Sales Journals by Customer Name"
      Item(3).ControlCount=   0
      Item(4).Caption =   "Vendor Opening Bal vs GL"
      Item(4).ControlCount=   0
      Item(5).Caption =   "Wrong Customer Code"
      Item(5).Tooltip =   "Search Sales Journals by Invoice Number"
      Item(5).ControlCount=   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View"
      Height          =   405
      Left            =   8040
      TabIndex        =   11
      Top             =   60
      Width           =   1605
   End
   Begin XtremeShortcutBar.ShortcutCaption sc2 
      Height          =   525
      Left            =   4560
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2235
      _Version        =   655364
      _ExtentX        =   3942
      _ExtentY        =   926
      _StockProps     =   14
      Caption         =   "Search Vendor Code:"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      ForeColor       =   4210752
   End
   Begin XtremeShortcutBar.ShortcutCaption sc1 
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11445
      _Version        =   655364
      _ExtentX        =   20188
      _ExtentY        =   926
      _StockProps     =   14
      Caption         =   "Transaction Status Tool:"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      ForeColor       =   4210752
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Visible         =   0   'False
      Begin VB.Menu mnuSalesJournal 
         Caption         =   "View Sales Journal"
      End
      Begin VB.Menu mnuCRJ 
         Caption         =   "View Cash Receipts Journal"
      End
   End
End
Attribute VB_Name = "frmTransactionStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsVoucherNo                                        As ADODB.Recordset
Dim i, X                                               As Integer
Attribute X.VB_VarUserMemId = 1073938433
Dim xREMARKS                                           As String
Attribute xREMARKS.VB_VarUserMemId = 1073938435
Dim xVOUCHERNO                                         As String
Attribute xVOUCHERNO.VB_VarUserMemId = 1073938436
Dim xLoading                                           As String
Attribute xLoading.VB_VarUserMemId = 1073938437
Dim Status                                             As String
Attribute Status.VB_VarUserMemId = 1073938438

Private Sub cmdInvalid_Click()
    Status = "INVALID"
    initGrid
    StoreMemVars
End Sub

Private Sub cmdOpening_Click()
    Status = "SLGL"
    initGrid
    StoreMemVars
    If gridStatus.Rows = 1 Then Exit Sub
    FlexGrid_To_Excel gridStatus, gridStatus.Rows, gridStatus.Cols, 1, Status
End Sub

Private Sub cmdTrialBalance_Click()
    Status = "TRIAL BALANCE"
    initGrid
    StoreMemVars
End Sub

Public Sub FlexGrid_To_Excel(TheFlexgrid As MSFlexGrid, TheRows As Integer, TheCols As Integer, Optional GridStyle As Integer = 1, Optional WorkSheetName As String)
    Dim objXL                                          As New Excel.Application
    Dim wbXL                                           As New Excel.Workbook
    Dim wsXL                                           As New Excel.Worksheet
    Dim intRow                                         As Integer    ' counter
    Dim intCol                                         As Integer    ' counter

    If Not IsObject(objXL) Then
        MsgBox "You need Microsoft Excel to use this function", _
               vbExclamation, "Print to Excel"
        Exit Sub
    End If

    'On Error Resume Next is necessary because
    'someone may pass more rows
    'or columns than the flexgrid has

    'you can instead check for this,
    'or rewrite the function so that
    'it exports all non-fixed cells
    'to Excel

    On Error Resume Next

    ' open Excel

    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet

    ' name the worksheet
    With wsXL
        If Not WorkSheetName = "" Then
            .Name = WorkSheetName
        End If
    End With

    ' fill worksheet
    For intRow = 1 To TheRows
        For intCol = 1 To TheCols
            With TheFlexgrid
                wsXL.Cells(intRow, intCol).Value = _
                .TextMatrix(intRow - 1, intCol - 1) & " "
            End With
        Next
    Next

    ' format the look
    For intCol = 1 To TheCols
        wsXL.Columns(intCol).AutoFit
        'wsXL.Columns(intCol).AutoFormat (1)
        wsXL.Range("A1", Right(wsXL.Columns(TheCols).AddressLocal, 1) & TheRows).AutoFormat GridStyle
    Next
    objXL.Visible = True
End Sub

Private Sub cmdPrint_Click()
    If gridStatus.Rows = 1 Then
        MsgBox " No Records To Print!", vbInformation
        Exit Sub
    End If
    FlexGrid_To_Excel gridStatus, gridStatus.Rows, gridStatus.Cols, 1, Status
    cmdPrint.Enabled = False
End Sub

Private Sub Command1_Click()

    Select Case SearchTab.SelectedItem
    Case 0
        'Status = "TRIAL BALANCE"
        initGrid
        StoreMemVars
    Case 1
        'Status = "VENDOR"
        initGrid
        StoreMemVars
    Case 2
        'Status = "INVALIDREF"
        initGrid
        StoreMemVars
    Case 3
        'Status = "SLGLCUSTOMER"
        initGrid
        StoreMemVars
    Case 4
        'Status = "SLGLVENDOR"
        initGrid
        StoreMemVars
    Case 5
        'Status = "WRONGCODE"
        initGrid
        StoreMemVars
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    SearchTab.SelectedItem = 0
    sc1.Caption = "Transaction Status Tool: "

    Status = "TRIAL BALANCE"
    initGrid
    StoreMemVars
End Sub

Sub StoreMemVars()
    cmdPrint.Enabled = False
    gridStatus.Redraw = True
    picLoading.Visible = True
    picLoading.Refresh
    If Status = "TRIAL BALANCE" Then
        Dim rsTrialBalance                             As ADODB.Recordset
        Set rsTrialBalance = New ADODB.Recordset
        rsTrialBalance.Open "SELECT * FROM (SELECT VOUCHERNO,JTYPE,SUM(DEBIT) AS DEBIT, SUM(CREDIT) AS CREDIT FROM AMIS_JOURNAL_DET " & _
                            " WHERE STATUS = 'P' GROUP BY VOUCHERNO,JTYPE) A WHERE DEBIT <> CREDIT", gconDMIS
        If Not rsTrialBalance.EOF And Not rsTrialBalance.BOF Then
            If rsTrialBalance!Debit <> rsTrialBalance!Credit Then
                xREMARKS = "Trial Balance NOT Balanced"
            End If
            PROGBAR.Value = 0
            PROGBAR.Max = rsTrialBalance.RecordCount
            Do While Not rsTrialBalance.EOF
                DoEvents
                If N2Str2Zero(rsTrialBalance!Credit) - N2Str2Zero(rsTrialBalance!Debit) <> 0 Then
                    xLoading = rsTrialBalance!VOUCHERNO
                    gridStatus.AddItem Null2String(rsTrialBalance!VOUCHERNO) & Chr(9) & Null2String(rsTrialBalance!jtype) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsTrialBalance!Debit)) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsTrialBalance!Credit)) & Chr(9) & xREMARKS
                    SearchTab(0).Enabled = True: SearchTab(1).Enabled = False: SearchTab(2).Enabled = False: SearchTab(3).Enabled = False: SearchTab(4).Enabled = False: SearchTab(5).Enabled = False
                End If
                LoadPercent
                rsTrialBalance.MoveNext
            Loop

        Else
            MessagePop Star, "", "Trial Balance is BALANCED"
        End If
        Set rsTrialBalance = Nothing
        picLoading.Visible = False
        SearchTab(0).Enabled = True: SearchTab(1).Enabled = True: SearchTab(2).Enabled = True: SearchTab(3).Enabled = True: SearchTab(4).Enabled = True: SearchTab(5).Enabled = True
    ElseIf Status = "VENDOR" Then
        Dim xVOUCHERNO                                 As String
        Dim rsVendorAr                                 As ADODB.Recordset
        Set rsVendorAr = New ADODB.Recordset
        Dim PayableAmnt                                As Double
        rsVendorAr.Open "SELECT  DET.Acct_Code,DET.Acct_Name,DET.Debit,DET.Credit,HD.VoucherNo,HD.VendorCode,AV.NameofVendor,HD.JDate,HD.JType,HD.Status FROM " & _
                        "AMIS_Journal_HD HD INNER JOIN AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.jtype = DET.jtype INNER JOIN ALL_VENDOR AV ON AV.CODE = HD.VendorCode " & _
                        "WHERE (HD.JType = 'CDJ') AND (LEFT(DET.Acct_Code, 5) = '11-02') AND HD.status = 'P' ORDER BY DET.Acct_Code ASC, AV.Code ASC", gconDMIS
        If Not rsVendorAr.EOF And Not rsVendorAr.BOF Then
            PROGBAR.Value = 0
            PROGBAR.Max = rsVendorAr.RecordCount
            Do While Not rsVendorAr.EOF
                DoEvents
                If rsVendorAr!Debit = 0 Then
                    PayableAmnt = NumericVal(rsVendorAr!Credit) * (-1)
                Else
                    PayableAmnt = NumericVal(rsVendorAr!Debit)
                End If
                xLoading = rsVendorAr!VOUCHERNO
                gridStatus.AddItem Null2String(rsVendorAr!VOUCHERNO) & Chr(9) & _
                                   Null2String(Format(rsVendorAr!JDATE, "mm/dd/yyyy")) & Chr(9) & _
                                   ToDoubleNumber(N2Str2Zero(PayableAmnt)) & Chr(9) & _
                                   Null2String(rsVendorAr!jtype) & Chr(9) & _
                                   Null2String(rsVendorAr!ACCT_CODE) & Chr(9) & _
                                   Null2String(rsVendorAr!acct_Name) & Chr(9) & _
                                   Null2String(rsVendorAr!VendorCode) & Chr(9) & _
                                   Null2String(rsVendorAr!nameofvendor)
                rsVendorAr.MoveNext
                gridStatus.TopRow = gridStatus.Rows - 1
                LoadPercent
                SearchTab(0).Enabled = False: SearchTab(1).Enabled = True: SearchTab(2).Enabled = False: SearchTab(3).Enabled = False: SearchTab(4).Enabled = False: SearchTab(5).Enabled = False
            Loop
        End If
        Set rsVendorAr = Nothing
        picLoading.Visible = False
        SearchTab(0).Enabled = True: SearchTab(1).Enabled = True: SearchTab(2).Enabled = True: SearchTab(3).Enabled = True: SearchTab(4).Enabled = True: SearchTab(5).Enabled = True
    ElseIf Status = "INVALIDREF" Then
        Dim rsInvalid_refNO                            As ADODB.Recordset
        Dim rsIs_Found_SJ                              As ADODB.Recordset
        Set rsInvalid_refNO = New ADODB.Recordset
        rsInvalid_refNO.Open "SELECT VOUCHERNO,INVOICEDATE,INVOICENO,INVOICETYPE,INVOICEAMOUNT FROM AMIS_CRJ_DETAIL", gconDMIS
        PROGBAR.Value = 0
        PROGBAR.Max = rsInvalid_refNO.RecordCount
        If Not rsInvalid_refNO.EOF And Not rsInvalid_refNO.BOF Then
            Do While Not rsInvalid_refNO.EOF
                DoEvents
                Set rsIs_Found_SJ = gconDMIS.Execute("Select INVOICENO,INVOICETYPE FROM AMIS_JOURNAL_HD WHERE INVOICENO = '" & Null2String(rsInvalid_refNO!INVOICENO) & "' AND INVOICETYPE = '" & rsInvalid_refNO!InvoiceType & "' and JTYPE = 'SJ'")
                If rsIs_Found_SJ.EOF And rsIs_Found_SJ.BOF Then
                    gridStatus.AddItem Null2String(rsInvalid_refNO!VOUCHERNO) & Chr(9) & Format(Null2String(rsInvalid_refNO!invoicedate), "mm/dd/yyyy") & Chr(9) & Null2String(rsInvalid_refNO!INVOICENO) & Chr(9) & Null2String(rsInvalid_refNO!InvoiceType) & Chr(9) & ToDoubleNumber(rsInvalid_refNO!invoiceamount)
                Else
                    'CORRECT REFERENCE
                End If
                xLoading = Null2String(rsInvalid_refNO!VOUCHERNO)
                rsInvalid_refNO.MoveNext
                '                    gridStatus.TopRow = gridStatus.Rows - 1
                LoadPercent
                SearchTab(0).Enabled = False: SearchTab(1).Enabled = False: SearchTab(2).Enabled = True: SearchTab(3).Enabled = False: SearchTab(4).Enabled = False: SearchTab(5).Enabled = False
            Loop
        End If
        Set rsInvalid_refNO = Nothing
        Set rsIs_Found_SJ = Nothing
        picLoading.Visible = False
        SearchTab(0).Enabled = True: SearchTab(1).Enabled = True: SearchTab(2).Enabled = True: SearchTab(3).Enabled = True: SearchTab(4).Enabled = True: SearchTab(5).Enabled = True
    ElseIf Status = "SLGLCUSTOMER" Then
        Dim rsSLGLCustomer                             As ADODB.Recordset
        Set rsSLGLCustomer = New ADODB.Recordset
        If COMPANY_CODE = "HCA" Then
            rsSLGLCustomer.Open "SELECT * FROM (SELECT ACCTCODE AS ACCT_CODE,DESCRIPTION AS ACCT_NAME," & _
                                "CAST(ISNULL((SELECT  CASE WHEN H.JTYPE='COB' THEN ABS(SUM(H.INVOICEAMT)) ELSE SUM(H.AMOUNTTOPAY) END  FROM AMIS_JOURNAL_HD H  INNER JOIN AMIS_JOURNAL_DET D  ON H.VOUCHERNO=D.VOUCHERNO  AND  H.JTYPE=D.JTYPE WHERE  H.JTYPE IN ('COB') AND H.STATUS='P' AND ACCT_CODE =ACCTCODE GROUP BY H.JTYPE),0) AS DECIMAL(18,2)) [SLBALANCE] , " & _
                                "ISNULL((SELECT SUM(ISNULL(DEBIT,0))-SUM(ISNULL(CREDIT,0)) FROM AMIS_JOURNAL_DET WHERE JTYPE='OPB' AND STATUS='P' AND ACCT_CODE=ACCTCODE),0) [GLBALANCE] " & _
                                "FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1 AND LEFT(ACCTCODE,5) IN ('11-02','11-03') " & _
                                ") T WHERE (SLBALANCE<>0 OR GLBALANCE<>0) ORDER BY ACCT_CODE", gconDMIS, adOpenKeyset
        Else
            rsSLGLCustomer.Open "SELECT * FROM (SELECT ACCTCODE AS ACCT_CODE,DESCRIPTION AS ACCT_NAME," & _
                                "CAST(ISNULL((SELECT  CASE WHEN H.JTYPE='COB' THEN ABS(SUM(H.INVOICEAMT)) ELSE SUM(H.AMOUNTTOPAY) END  FROM AMIS_JOURNAL_HD H  INNER JOIN AMIS_JOURNAL_DET D  ON H.VOUCHERNO=D.VOUCHERNO  AND  H.JTYPE=D.JTYPE WHERE  H.JTYPE IN ('COB','VPJ') AND H.STATUS='P' AND ACCT_CODE =ACCTCODE GROUP BY H.JTYPE),0) AS DECIMAL(18,2)) [SLBALANCE] , " & _
                                "ISNULL((SELECT SUM(ISNULL(DEBIT,0))-SUM(ISNULL(CREDIT,0)) FROM AMIS_JOURNAL_DET WHERE JTYPE='OPB' AND STATUS='P' AND ACCT_CODE=ACCTCODE),0) [GLBALANCE] " & _
                                "FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1 AND LEFT(ACCTCODE,5) IN ('11-02','11-03') " & _
                                ") T WHERE (SLBALANCE<>0 OR GLBALANCE<>0) ORDER BY ACCT_CODE", gconDMIS, adOpenKeyset
        End If
        If Not rsSLGLCustomer.EOF And Not rsSLGLCustomer.BOF Then
            X = 0
            PROGBAR.Value = 0
            PROGBAR.Max = rsSLGLCustomer.RecordCount
            Do While Not rsSLGLCustomer.EOF
                DoEvents
                xLoading = rsSLGLCustomer!ACCT_CODE
                If NumericVal(rsSLGLCustomer![SLBALANCE]) <> NumericVal(rsSLGLCustomer![GLBALANCE]) Then
                    xREMARKS = "Unbalanced"
                Else
                    xREMARKS = "Balanced"
                End If
                gridStatus.AddItem Null2String(rsSLGLCustomer!ACCT_CODE) & Chr(9) & Null2String(rsSLGLCustomer!acct_Name) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsSLGLCustomer![SLBALANCE])) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsSLGLCustomer![GLBALANCE])) & Chr(9) & xREMARKS
                rsSLGLCustomer.MoveNext
                SearchTab(0).Enabled = False: SearchTab(1).Enabled = False: SearchTab(2).Enabled = False: SearchTab(3).Enabled = True: SearchTab(4).Enabled = False: SearchTab(5).Enabled = False
                LoadPercent
                X = X + 1
                If xREMARKS = "Unbalanced" Then
                    With gridStatus
                        .Row = X                      ' the row you want to highlight
                        For i = 0 To .Cols - 1
                            .Col = i
                            .CellBackColor = QBColor(14)
                        Next i
                    End With
                End If
            Loop

        Else
            MessagePop Star, "", "Customer Opening Balance and GL is BALANCE"
        End If
        Set rsSLGLCustomer = Nothing
        picLoading.Visible = False
        SearchTab(0).Enabled = True: SearchTab(1).Enabled = True: SearchTab(2).Enabled = True: SearchTab(3).Enabled = True: SearchTab(4).Enabled = True: SearchTab(5).Enabled = True
    ElseIf Status = "SLGLVENDOR" Then
        Dim rsSLGLVendor                               As ADODB.Recordset
        Set rsSLGLVendor = New ADODB.Recordset
        rsSLGLVendor.Open "SELECT * FROM (SELECT ACCTCODE AS ACCT_CODE, DESCRIPTION AS ACCT_NAME, " & _
                          "CAST(ISNULL((SELECT CASE WHEN H.JTYPE = 'VPJ' THEN ABS(SUM(H.AMOUNTTOPAY)) ELSE SUM(H.INVOICEAMT) END FROM AMIS_JOURNAL_HD H  INNER JOIN AMIS_JOURNAL_DET D  ON H.VOUCHERNO=D.VOUCHERNO  AND  H.JTYPE=D.JTYPE WHERE  H.JTYPE IN ('VPJ','COB') AND H.STATUS='P' AND ACCT_CODE =ACCTCODE GROUP BY H.JTYPE),0) AS DECIMAL(18,2)) [SLBALANCE] , " & _
                          "ISNULL((SELECT SUM(ISNULL(CREDIT,0))-SUM(ISNULL(DEBIT,0)) FROM AMIS_JOURNAL_DET WHERE JTYPE='OPB' AND STATUS='P' AND ACCT_CODE=ACCTCODE),0) [GLBALANCE] " & _
                          "FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1 AND LEFT(ACCTCODE,5) IN('21-01','21-02','21-07') " & _
                          ") T WHERE (SLBALANCE<>0 OR GLBALANCE<>0) ORDER BY ACCT_CODE", gconDMIS, adOpenKeyset
        If Not rsSLGLVendor.EOF And Not rsSLGLVendor.BOF Then
            X = 0
            PROGBAR.Value = 0
            PROGBAR.Max = rsSLGLVendor.RecordCount
            Do While Not rsSLGLVendor.EOF
                DoEvents
                xLoading = rsSLGLVendor!ACCT_CODE
                If NumericVal(rsSLGLVendor![SLBALANCE]) <> NumericVal(rsSLGLVendor![GLBALANCE]) Then
                    xREMARKS = "Unbalanced"
                Else
                    xREMARKS = "Balanced"
                End If
                gridStatus.AddItem Null2String(rsSLGLVendor!ACCT_CODE) & Chr(9) & Null2String(rsSLGLVendor!acct_Name) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsSLGLVendor![SLBALANCE])) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsSLGLVendor![GLBALANCE])) & Chr(9) & xREMARKS
                rsSLGLVendor.MoveNext
                LoadPercent
                SearchTab(0).Enabled = False: SearchTab(1).Enabled = False: SearchTab(2).Enabled = False: SearchTab(3).Enabled = False: SearchTab(4).Enabled = True: SearchTab(5).Enabled = False
                X = X + 1
                If xREMARKS = "Unbalanced" Then
                    With gridStatus
                        .Row = X                      ' the row you want to highlight
                        For i = 0 To .Cols - 1
                            .Col = i
                            .CellBackColor = QBColor(14)
                        Next i
                    End With
                End If
            Loop
        Else
            MessagePop Star, "INFORMATION", "Vendor Opening Balance and GL is BALANCE"
        End If
        Set rsSLGLVendor = Nothing
        picLoading.Visible = False
        SearchTab(0).Enabled = True: SearchTab(1).Enabled = True: SearchTab(2).Enabled = True: SearchTab(3).Enabled = True: SearchTab(4).Enabled = True: SearchTab(5).Enabled = True
    ElseIf Status = "WRONGCODE" Then
        Dim rsWrongCode                                As ADODB.Recordset
        Dim rsCheck_Code                               As ADODB.Recordset
        Dim rsINV_INVTYPE                              As ADODB.Recordset
        Set rsWrongCode = New ADODB.Recordset
        rsWrongCode.Open "SELECT VOUCHERNO,CR_TYPE,INVOICENO,INVOICETYPE FROM " & _
                         "AMIS_CRJ_DETAIL where VoucherNo is not NULL " & _
                         "GROUP BY VOUCHERNO,CR_TYPE,INVOICENO,INVOICETYPE ", gconDMIS


        If Not rsWrongCode.EOF And Not rsWrongCode.BOF Then
            PROGBAR.Value = 0
            PROGBAR.Max = rsWrongCode.RecordCount

            DoEvents
            Do While Not rsWrongCode.EOF
                If rsWrongCode!VOUCHERNO = "" Then
                    rsWrongCode.MoveNext
                    DoEvents
                Else
                    Set rsINV_INVTYPE = New ADODB.Recordset
                    rsINV_INVTYPE.Open "Select CustomerCode,VoucherNo,Jtype from Amis_Journal_hd where VoucherNo = '" & Null2String(rsWrongCode!VOUCHERNO) & "' and Jtype = 'CRJ'", gconDMIS
                    If Not rsINV_INVTYPE.EOF And Not rsINV_INVTYPE.BOF Then
                        Set rsCheck_Code = gconDMIS.Execute("Select VoucherNo,CustomerCode from Amis_Journal_hd where InvoiceNo = '" & Null2String(rsWrongCode!INVOICENO) & "' and  InvoiceType  = '" & rsWrongCode!InvoiceType & "' and Jtype = 'SJ'")
                        If Not rsCheck_Code.EOF And Not rsCheck_Code.BOF Then
                            If Null2String(rsINV_INVTYPE!CustomerCode) <> Null2String(rsCheck_Code!CustomerCode) Then
                                gridStatus.AddItem Null2String(rsCheck_Code!VOUCHERNO) & Chr(9) & _
                                                   Null2String(rsCheck_Code!CustomerCode) & Chr(9) & _
                                                   GetSJ_CustomerName(Null2String(rsINV_INVTYPE!CustomerCode)) & Chr(9) & _
                                                   Null2String(rsINV_INVTYPE!VOUCHERNO) & Chr(9) & _
                                                   Null2String(rsINV_INVTYPE!CustomerCode) & Chr(9) & _
                                                   GetSJ_CustomerName(Null2String(rsINV_INVTYPE!CustomerCode)) & Chr(9) & _
                                                   Null2String(rsWrongCode!INVOICENO)
                            Else
                                'correct code
                            End If
                        End If
                    End If
                    rsWrongCode.MoveNext
                    DoEvents
                End If
                'xLoading = Null2String(rsWrongCode!VOUCHERNO)
                'rsWrongCode.MoveNext
                LoadPercent
                If gridStatus.Row = 0 Then
                Else
                    gridStatus.TopRow = gridStatus.Rows - 1
                End If

            Loop
        End If

        Set rsWrongCode = Nothing
        Set rsCheck_Code = Nothing
        Set rsINV_INVTYPE = Nothing
        picLoading.Visible = False
        SearchTab(0).Enabled = True: SearchTab(1).Enabled = True: SearchTab(2).Enabled = True: SearchTab(3).Enabled = True: SearchTab(4).Enabled = True: SearchTab(5).Enabled = True
    End If
    If gridStatus.Rows > 1 Then
        cmdPrint.Enabled = True
    End If
    gridStatus.Redraw = True
    gridStatus.Refresh
End Sub

Private Sub gridStatus_DblClick()
    If Status = "TRIAL BALANCE" Then
        'gridStatus.Row = gridStatus.Row
        gridStatus.Col = 1
        JOURNALTYPE = gridStatus.Text
        gridStatus.Col = 0
        Set rsVoucherNo = gconDMIS.Execute("SELECT * FROM (SELECT VOUCHERNO,JTYPE,SUM(DEBIT) AS DEBIT, SUM(CREDIT) AS CREDIT FROM AMIS_JOURNAL_DET " & _
                                           " WHERE STATUS = 'P' AND VOUCHERNO = '" & gridStatus.Text & "' GROUP BY VOUCHERNO,JTYPE) A WHERE DEBIT <> CREDIT")
        If Not rsVoucherNo.EOF And Not rsVoucherNo.BOF Then
            xVOUCHERNO = rsVoucherNo!VOUCHERNO
            frmAMISJournalEntry.Show
            frmAMISJournalEntry.SearchVoucherNo Trim(xVOUCHERNO)
            frmAMISJournalEntry.ZOrder 0
        End If
        Set rsVoucherNo = Nothing
    ElseIf Status = "VENDOR" Then
        'gridStatus.Row = gridStatus.Row
        gridStatus.Col = 3
        JOURNALTYPE = gridStatus.Text
        gridStatus.Col = 0
        JOURNALTYPE = gridStatus.TextMatrix(gridStatus.RowSel, gridStatus.ColSel)
        Set rsVoucherNo = gconDMIS.Execute("SELECT  DET.Acct_Code,DET.Acct_Name,DET.Debit,DET.Credit,HD.VoucherNo,HD.VendorCode,AV.NameofVendor,HD.JDate,HD.JType,HD.Status FROM " & _
                                           "AMIS_Journal_HD HD INNER JOIN AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.jtype = DET.jtype INNER JOIN ALL_VENDOR AV ON AV.CODE = HD.VendorCode " & _
                                           "WHERE (HD.JType = 'CDJ') AND (LEFT(DET.Acct_Code, 5) = '11-02') AND HD.status = 'P' AND HD.VOUCHERNO = '" & gridStatus.Text & "' ORDER BY DET.Acct_Code ASC, AV.Code ASC")
        If Not rsVoucherNo.EOF And Not rsVoucherNo.BOF Then

            xVOUCHERNO = rsVoucherNo!VOUCHERNO
            frmAMISJournalEntry.Show
            frmAMISJournalEntry.SearchVoucherNo Trim(xVOUCHERNO)
            frmAMISJournalEntry.ZOrder 0
        End If
        Set rsVoucherNo = Nothing
    ElseIf Status = "INVALIDREF" Then
        Dim xInvoiceType, xInvoiceNo                   As String
        'gridStatus.Row = gridStatus.Row
        gridStatus.Col = 2
        xInvoiceNo = gridStatus.Text
        gridStatus.Col = 3
        xInvoiceType = gridStatus.Text
        JOURNALTYPE = "SJ"
        gridStatus.Col = 0
        Dim rsInvalid_refNO                            As ADODB.Recordset
        Dim rsIs_Found_SJ                              As ADODB.Recordset
        Set rsInvalid_refNO = New ADODB.Recordset
        rsInvalid_refNO.Open "SELECT VOUCHERNO,INVOICEDATE,INVOICENO,INVOICETYPE,INVOICEAMOUNT FROM AMIS_CRJ_DETAIL WHERE INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "' and VoucherNo = '" & gridStatus.Text & "'", gconDMIS
        If Not rsInvalid_refNO.EOF And Not rsInvalid_refNO.BOF Then
            'Set rsIs_Found_SJ = gconDMIS.Execute("Select INVOICENO,INVOICETYPE,VOUCHERNO FROM AMIS_JOURNAL_HD WHERE INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "' and JTYPE = 'CRJ' and VoucherNo = '" & gridStatus.Text & "'")
            'If Not rsIs_Found_SJ.EOF And Not rsIs_Found_SJ.BOF Then
            xVOUCHERNO = Null2String(rsInvalid_refNO!VOUCHERNO)
            frmAMISJournalEntry.Show
            frmAMISJournalEntry.SearchVoucherNo Trim(xVOUCHERNO)
            frmAMISJournalEntry.ZOrder 0
            'End If
        End If
        Set rsInvalid_refNO = Nothing
        Set rsIs_Found_SJ = Nothing
    ElseIf Status = "SLGLCUSTOMER" Then
    ElseIf Status = "SLGLVENDOR" Then
    ElseIf Status = "WRONGCODE" Then
    End If
End Sub

Sub initGrid()
    gridStatus.Clear
    If Status = "TRIAL BALANCE" Then
        With gridStatus
            .Rows = 1
            .Cols = 5
            .ColWidth(0) = 1400
            .ColWidth(1) = 1200
            .ColWidth(2) = 1400
            .ColWidth(3) = 1400
            .ColWidth(4) = 2100
            .Row = 0
            .Col = 0: .Text = "Voucher #"
            .Col = 1: .Text = "Journal Type"
            .Col = 2: .Text = "Debit"
            .Col = 3: .Text = "Credit"
            .Col = 4: .Text = "Remarks"
            .ColAlignment(0) = flexAlignLeftCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .ColAlignment(2) = flexAlignRightCenter
            .ColAlignment(3) = flexAlignRightCenter
            .ColAlignment(4) = flexAlignLeftCenter
        End With
        sc1.Caption = "Transaction Status Tool: " & SearchTab.Item(0).Caption
        txtSearch.Visible = False
        sc2.Visible = False
    ElseIf Status = "VENDOR" Then
        With gridStatus
            .Rows = 1
            .Cols = 8
            .ColWidth(0) = 900
            .ColWidth(1) = 1200
            .ColWidth(2) = 1200
            .ColWidth(3) = 600
            .ColWidth(4) = 1200
            .ColWidth(5) = 3800
            .ColWidth(6) = 1400
            .ColWidth(7) = 4200
            .Row = 0
            .Col = 0: .Text = "Voucher #"
            .Col = 1: .Text = "Journal Date"
            .Col = 2: .Text = "Payables"
            .Col = 3: .Text = "Journal Type"
            .Col = 4: .Text = "Account Code"
            .Col = 5: .Text = "Account Name"
            .Col = 6: .Text = "Vendor Code"
            .Col = 7: .Text = "Name"
            .ColAlignment(0) = flexAlignLeftCenter
            .ColAlignment(1) = flexAlignCenterCenter
            .ColAlignment(2) = flexAlignRightCenter
            .ColAlignment(3) = flexAlignCenterCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .ColAlignment(5) = flexAlignLeftCenter
            .ColAlignment(6) = flexAlignCenterCenter
            .ColAlignment(7) = flexAlignLeftCenter
        End With
        sc1.Caption = "Transaction Status Tool: " & SearchTab.Item(1).Caption
        txtSearch.Visible = True
        sc2.Visible = True
    ElseIf Status = "INVALIDREF" Then
        With gridStatus
            .Rows = 1
            .Cols = 5
            .ColWidth(0) = 1600
            .ColWidth(1) = 1600
            .ColWidth(2) = 1500
            .ColWidth(3) = 1500
            .ColWidth(4) = 1800
            .Row = 0
            .Col = 0: .Text = "VOUCHERNO"
            .Col = 1: .Text = "INVOICE DATE"
            .Col = 2: .Text = "INVOICE NO."
            .Col = 3: .Text = "INVOICE TYPE"
            .Col = 4: .Text = "INVOICE AMOUNT"
            .ColAlignment(0) = flexAlignCenterCenter
            .ColAlignment(1) = flexAlignCenterCenter
            .ColAlignment(2) = flexAlignLeftCenter
            .ColAlignment(3) = flexAlignCenterCenter
            .ColAlignment(4) = flexAlignRightCenter
        End With
        sc1.Caption = "Transaction Status Tool: " & SearchTab.Item(2).Caption
        txtSearch.Visible = False
        sc2.Visible = False
    ElseIf Status = "SLGLCUSTOMER" Then
        With gridStatus
            .Rows = 1
            .Cols = 5
            .ColWidth(0) = 1600
            .ColWidth(1) = 4700
            .ColWidth(2) = 1500
            .ColWidth(3) = 1500
            .ColWidth(4) = 1300
            .Row = 0
            .Col = 0: .Text = "ACCOUNT CODE"
            .Col = 1: .Text = "ACCOUNT NAME"
            .Col = 2: .Text = "SL"
            .Col = 3: .Text = "GL"
            .Col = 4: .Text = "REMARKS":
            .ColAlignment(0) = flexAlignCenterCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .ColAlignment(2) = flexAlignRightCenter
            .ColAlignment(3) = flexAlignRightCenter
            .ColAlignment(4) = flexAlignLeftCenter

        End With
        sc1.Caption = "Transaction Status Tool: " & SearchTab.Item(3).Caption
        txtSearch.Visible = False
        sc2.Visible = False
    ElseIf Status = "SLGLVENDOR" Then
        With gridStatus
            .Rows = 1
            .Cols = 5
            .ColWidth(0) = 1600
            .ColWidth(1) = 4700
            .ColWidth(2) = 1500
            .ColWidth(3) = 1500
            .ColWidth(4) = 1300
            .Row = 0
            .Col = 0: .Text = "ACCOUNT CODE"
            .Col = 1: .Text = "ACCOUNT NAME"
            .Col = 2: .Text = "SL"
            .Col = 3: .Text = "GL"
            .Col = 4: .Text = "REMARKS"
            .ColAlignment(0) = flexAlignCenterCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .ColAlignment(2) = flexAlignRightCenter
            .ColAlignment(3) = flexAlignRightCenter
            .ColAlignment(4) = flexAlignLeftCenter
        End With
        sc1.Caption = "Transaction Status Tool: " & SearchTab.Item(4).Caption
        txtSearch.Visible = False
        sc2.Visible = False
    ElseIf Status = "WRONGCODE" Then
        With gridStatus
            .Rows = 1
            .Cols = 6
            .ColWidth(0) = 1600
            .ColWidth(1) = 1600
            .ColWidth(2) = 1600
            .ColWidth(3) = 1600
            .ColWidth(4) = 1600
            .ColWidth(5) = 4200
            .Row = 0
            .Col = 0: .Text = "SJ VOUCHERNO"
            .Col = 1: .Text = "SJ CUST. CODE"
            .Col = 2: .Text = "CRJ VOUCHERNO"
            .Col = 3: .Text = "INVOICENO"
            .Col = 4: .Text = "CUST. CODE"
            .Col = 5: .Text = "CUST. NAME"
            .ColAlignment(0) = flexAlignCenterCenter
            .ColAlignment(1) = flexAlignCenterCenter
            .ColAlignment(2) = flexAlignCenterCenter
            .ColAlignment(3) = flexAlignCenterCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .ColAlignment(5) = flexAlignLeftCenter
        End With
        sc1.Caption = "Transaction Status Tool: " & SearchTab.Item(5).Caption
        txtSearch.Visible = False
        sc2.Visible = False
    End If
End Sub

Private Sub gridStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Status = "WRONGCODE" Then
        If Button = vbRightButton Then
            PopupMenu mnuOption
        End If
    End If
End Sub

Private Sub mnuCRJ_Click()
    gridStatus.Row = gridStatus.Row
    gridStatus.Col = 3
    '    frmAMIS_CRJ_JOURNAL_ENTRY.LoadJournal ("CRJ")
    '    frmAMIS_CRJ_JOURNAL_ENTRY.rsRefresh
    '    frmAMIS_CRJ_JOURNAL_ENTRY.SearchVoucherNo (gridStatus.Text)

End Sub

Private Sub mnuSalesJournal_Click()
    Dim xVOUCHERNO                                     As String
    xVOUCHERNO = gridStatus.Text
    '    frmAMIS_SJ_JOURNAL_ENTRY.LoadJournal ("SJ")
    '    frmAMIS_SJ_JOURNAL_ENTRY.rsRefresh
    '    frmAMIS_SJ_JOURNAL_ENTRY.SearchVoucherNo (Trim(xVoucherNo))
    JOURNALTYPE = "SJ"
    frmAMISJournalEntry.rsRefresh
    frmAMISJournalEntry.SearchVoucherNo (Trim(xVOUCHERNO))

End Sub

Private Sub SearchTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    SEARCH_TAB = SearchTab.SelectedItem
    Select Case SEARCH_TAB
    Case 0
        Status = "TRIAL BALANCE"
        initGrid
        'StoreMemvars
    Case 1
        Status = "VENDOR"
        initGrid
        'StoreMemvars
    Case 2
        Status = "INVALIDREF"
        initGrid
        'StoreMemvars
    Case 3
        Status = "SLGLCUSTOMER"
        initGrid
        'StoreMemvars
    Case 4
        Status = "SLGLVENDOR"
        initGrid
        'StoreMemvars
    Case 5
        Status = "WRONGCODE"
        initGrid
        'StoreMemvars
    End Select
End Sub

Sub LoadPercent()
    PROGBAR.Value = PROGBAR.Value + 1
    lblPercent = Round((PROGBAR.Value / PROGBAR.Max) * 100, 0) & "%"
    lblData.Caption = xLoading
End Sub

Private Sub txtSearch_Change()
    Dim xVOUCHERNO                                     As String
    Dim rsVendorAr                                     As ADODB.Recordset
    Set rsVendorAr = New ADODB.Recordset
    Dim PayableAmnt                                    As Double
    initGrid
    rsVendorAr.Open "SELECT  DET.Acct_Code,DET.Acct_Name,DET.Debit,DET.Credit,HD.VoucherNo,HD.VendorCode,AV.NameofVendor,HD.JDate,HD.JType,HD.Status FROM " & _
                    "AMIS_Journal_HD HD INNER JOIN AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.jtype = DET.jtype INNER JOIN ALL_VENDOR AV ON AV.CODE = HD.VendorCode " & _
                    "WHERE (HD.JType = 'CDJ') AND (LEFT(DET.Acct_Code, 5) = '11-02') AND HD.status = 'P' AND HD.VendorCode like '%" & txtSearch & "%' ORDER BY DET.Acct_Code ASC, AV.Code ASC", gconDMIS
    If Not rsVendorAr.EOF And Not rsVendorAr.BOF Then
        rsVendorAr.MoveFirst
        Do While Not rsVendorAr.EOF
            If rsVendorAr!Debit = 0 Then
                PayableAmnt = NumericVal(rsVendorAr!Credit) * (-1)
            Else
                PayableAmnt = NumericVal(rsVendorAr!Debit)
            End If
            xLoading = rsVendorAr!VOUCHERNO
            gridStatus.AddItem Null2String(rsVendorAr!VOUCHERNO) & Chr(9) & Null2String(Format(rsVendorAr!JDATE, "mm/dd/yyyy")) & Chr(9) & ToDoubleNumber(N2Str2Zero(PayableAmnt)) & Chr(9) & _
                               Null2String(rsVendorAr!jtype) & Chr(9) & _
                               Null2String(rsVendorAr!ACCT_CODE) & Chr(9) & _
                               Null2String(rsVendorAr!acct_Name) & Chr(9) & _
                               Null2String(rsVendorAr!VendorCode) & Chr(9) & _
                               Null2String(rsVendorAr!nameofvendor)
            rsVendorAr.MoveNext
        Loop
    End If
End Sub

Function GetSJ_CustomerName(xSJ_CRJ_Code As String) As String
    Dim rsSJ_CRJ_Name                                  As ADODB.Recordset
    Set rsSJ_CRJ_Name = gconDMIS.Execute("Select AcctName from All_Customer_Table where CusCde = '" & xSJ_CRJ_Code & "'")
    If Not rsSJ_CRJ_Name.EOF And Not rsSJ_CRJ_Name.BOF Then
        GetSJ_CustomerName = Null2String(rsSJ_CRJ_Name!AcctName)
    Else
        GetSJ_CustomerName = ""
    End If
    Set rsSJ_CRJ_Name = Nothing
End Function

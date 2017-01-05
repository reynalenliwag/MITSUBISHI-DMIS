VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frm_TOOLS_ARTOOLS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AR TOOLS"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   Icon            =   "ARTOOLS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   10335
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   5535
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   10305
      _Version        =   655364
      _ExtentX        =   18177
      _ExtentY        =   9763
      _StockProps     =   64
      PaintManager.BoldSelected=   -1  'True
      PaintManager.FixedTabWidth=   100
      PaintManager.MaxTabWidth=   100
      PaintManager.MinTabWidth=   100
      ItemCount       =   3
      Item(0).Caption =   "GJ W/AP"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "Grid2"
      Item(0).Control(1)=   "Label(0)"
      Item(0).Control(2)=   "Label(1)"
      Item(0).Control(3)=   "Texttotal"
      Item(0).Control(4)=   "cboacct"
      Item(0).Control(5)=   "Command"
      Item(1).Caption =   "AR TOOLS"
      Item(1).ControlCount=   0
      Item(2).Caption =   "UN BALANCED TB"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "Command3"
      Item(2).Control(1)=   "Grid1"
      Item(2).Control(2)=   "Command4"
      Begin VB.CommandButton Command4 
         Caption         =   "View Unbalanced Entry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -69910
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   2565
      End
      Begin FlexCell.Grid Grid1 
         Height          =   4665
         Left            =   -69940
         TabIndex        =   14
         Top             =   810
         Visible         =   0   'False
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   8229
         BackColorBkg    =   -2147483645
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         Rows            =   1
      End
      Begin VB.CommandButton Command3 
         Caption         =   "TB Refresher"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -62380
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.CommandButton Command 
         Caption         =   "generate"
         Height          =   435
         Left            =   8850
         TabIndex        =   12
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox Texttotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8010
         TabIndex        =   11
         Top             =   5130
         Width           =   1965
      End
      Begin VB.ComboBox cboacct 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   450
         Width           =   3795
      End
      Begin FlexCell.Grid Grid2 
         Height          =   4065
         Left            =   90
         TabIndex        =   7
         Top             =   1020
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   7170
         BackColorBkg    =   -2147483645
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label Label 
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   7350
         TabIndex        =   10
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label Label 
         Caption         =   "Account"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   480
         Width           =   1365
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   10440
      TabIndex        =   5
      Top             =   1470
      Width           =   615
   End
   Begin FlexCell.Grid thegrid 
      Height          =   4005
      Left            =   210
      TabIndex        =   4
      Top             =   1320
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   7064
      BackColorBkg    =   -2147483645
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin wizProgBar.Prg loadme 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   556
      Picture         =   "ARTOOLS.frx":058A
      ForeColor       =   0
      BarPicture      =   "ARTOOLS.frx":05A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   795
      Left            =   8400
      TabIndex        =   1
      Top             =   90
      Width           =   1635
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   1800
      TabIndex        =   16
      Top             =   1470
      Visible         =   0   'False
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   450
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "CRJ Date is greter the the date of the SAles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4635
   End
End
Attribute VB_Name = "frm_TOOLS_ARTOOLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboacct_Click()
getAccount cboacct.Text
End Sub

Private Sub Command_Click()
displayGJ
End Sub

Private Sub Command1_Click()
DisplayCRJ
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    If MsgBox("Are you sure ", vbInformation + vbYesNo) = vbNo Then Exit Sub
    Dim rsJournal_Det                                  As ADODB.Recordset
    Dim rsJournal_hd                                   As ADODB.Recordset
    Dim x                                              As Double

    Dim XVOUCHERNO As String
    Dim xJtype                                   As String



    For i = 1 To Grid1.Rows
        XVOUCHERNO = Grid1.Cell(i, 1).Text
        xJtype = Grid1.Cell(i, 2).Text
        DEBIT = Grid1.Cell(i, 3).Text
        CREDIT = Grid1.Cell(i, 4).Text
        i = i + 1
'        MsgBox xVoucherno & DEBIT & CREDIT
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("SELECT SUM(DEBIT) AS TOTALDEBIT,SUM(CREDIT) AS TOTALCREDIT FROM AMIS_JOURNAL_DET WHERE JTYPE = " & N2Str2Null(xJtype) & " AND VOUCHERNO = " & N2Str2Null(XVOUCHERNO) & "")

        If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
            If DEBIT <> CREDIT Then
           
           End If
        Else
            
        End If
        
        
    Next

     

'
'
'
'    'Check Lost
'    Set rsJournal_Det = New ADODB.Recordset
'    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_Det where jdate <= '3/1/2009' order by JNO ASC")
'    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
'        rsJournal_Det.MoveFirst
'        Do While Not rsJournal_Det.EOF
'            x = x + 1
'            Set rsJournal_hd = New ADODB.Recordset
'            Set rsJournal_hd = gconDMIS.Execute("Select * from AMIS_journal_HD where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & "  and VoucherNo = " & N2Str2Null(rsJournal_Det!VOUCHERNO))
'            If rsJournal_hd.EOF And rsJournal_hd.BOF Then
'                gconDMIS.Execute ("Delete from AMIS_Journal_Det where ID = " & rsJournal_Det!ID)
'            End If
'            Me.Caption = rsJournal_Det!jtype & rsJournal_Det!VOUCHERNO & "-" & x: DoEvents
'            rsJournal_Det.MoveNext
'        Loop
'    End If
'
'    'Refresh Amounts
'    Dim rsJournal_Det_Trans                            As ADODB.Recordset
'    Set rsJournal_Det = New ADODB.Recordset
'    Set rsJournal_Det = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_Det where jdate <= '5/31/2008' Group By Jtype,VoucherNo Having (SUM(Debit) <> SUM(Credit)) Order by Jtype,VoucherNo")
'    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
'        rsJournal_Det.MoveFirst
'        x = 0
'        Do While Not rsJournal_Det.EOF
'            x = x + 1
'            Set rsJournal_Det_Trans = New ADODB.Recordset
'            Set rsJournal_Det_Trans = gconDMIS.Execute("Select * from AMIS_Journal_Det Where JType = " & N2Str2Null(rsJournal_Det!jtype) & " AND VoucherNo = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " Order by JItemNo asc")
'            If Not rsJournal_Det_Trans.EOF And Not rsJournal_Det_Trans.BOF Then
'                rsJournal_Det_Trans.MoveFirst
'                Do While Not rsJournal_Det_Trans.EOF
'                    gconDMIS.Execute ("Update AMIS_Journal_Det set Debit = " & Round(NumericVal(rsJournal_Det_Trans!DEBIT), 2) & ", Credit = " & Round(NumericVal(rsJournal_Det_Trans!CREDIT), 2) & " where id = " & rsJournal_Det_Trans!ID)
'                    rsJournal_Det_Trans.MoveNext
'                Loop
'            End If
'            DoEvents
'            Me.Caption = rsJournal_Det!VOUCHERNO & "-" & x
'            rsJournal_Det.MoveNext
'        Loop
'    End If

    MsgBox "Completed"
End Sub

Private Sub Command4_Click()

    Dim rsx                                            As ADODB.Recordset
    Set rsx = gconDMIS.Execute("SELECT * FROM (SELECT VOUCHERNO,JTYPE,  CAST(SUM(DEBIT)AS DECIMAL(18,4)) DEBIT, CAST(SUM(CREDIT)AS DECIMAL(18,4)) CREDIT FROM AMIS_JOURNAL_DET WHERE STATUS='P' GROUP BY JTYPE,VOUCHERNO )T WHERE T.DEBIT<>T.CREDIT")
    Grid1.Rows = 1


    Grid1.AutoRedraw = False
    Dim SYS_REM                                        As String
    While Not rsx.EOF
        DoEvents
        SYS_REM = ""
        If Round(rsx(2), 2) = Round(rsx(3), 2) Then
            SYS_REM = "Rounding Err"
        Else
            SYS_REM = "Transaction Err"
        End If
        Grid1.AddItem rsx(0) & Chr(9) & rsx(1) & Chr(9) & rsx(2) & Chr(9) & rsx(3) & Chr(9) & SYS_REM, False

        rsx.MoveNext
    Wend
    Grid1.AutoRedraw = True
    Grid1.Refresh
    If Grid1.Rows > 1 Then
        Command3.Enabled = True
    End If

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
initthegrid
returnAcctDescription
InitGrid
End Sub
Sub DisplayCRJ()
    ' to display CRJ having jdate is greater then SJ jdate
    Dim RSSJ As New ADODB.Recordset
    Dim RSCRJ As New ADODB.Recordset
    Dim i As Integer
    Set RSSJ = gconDMIS.Execute("SELECT dbo.AMIS_Journal_HD.VoucherNo as SJVo,dbo.AMIS_Journal_HD.jdate as SJDATE,dbo.AMIS_Journal_HD.CustomerCode as SJCODE, dbo.AMIS_CRJ_Detail.INVOICETYPE as INV_TYPE,dbo.AMIS_CRJ_Detail.voucherno as CRJ_voucherno, dbo.AMIS_CRJ_Detail.INVOICENO as INV_NO,dbo.AMIS_CRJ_Detail.INVOICEAMOUNT as CRJ_AMOUNT FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_CRJ_Detail ON dbo.AMIS_Journal_HD.InvoiceNo = dbo.AMIS_CRJ_Detail.INVOICENO AND dbo.AMIS_Journal_HD.InvoiceType = dbo.AMIS_CRJ_Detail.InvoiceType WHERE (dbo.AMIS_Journal_HD.JType = 'SJ') and dbo.AMIS_Journal_HD.jdate <='8/31/2008'")
    loadme.Value = 0
    i = 0
    loadme.Max = RSSJ.RecordCount
    If Not (RSSJ.EOF And RSSJ.BOF) Then
        Do While Not RSSJ.EOF
            i = i + 1
            Set RSCRJ = gconDMIS.Execute("Select voucherno,invoiceamt,invoiceno,invoicetype,jdate as crj_date,invoicedate,customercode from AMIS_JOURNAL_HD where jtype='CRJ' and voucherno='" & RSSJ!CRJ_voucherno & "'")
                If Not (RSCRJ.EOF And RSCRJ.BOF) Then
                    'If RSCRJ!crj_date > RSSJ!sjdate Then
                    If RSSJ!sjdate > RSCRJ!crj_date And RSSJ!SJCODE = RSCRJ!CustomerCode Then
                         thegrid.AddItem RSCRJ![VOUCHERNO] & vbTab & getcutomername(RSCRJ!CustomerCode) & vbTab & _
                                          Null2String(RSSJ!INV_NO) & vbTab & Null2String(RSSJ!INV_TYPE) & vbTab & _
                                          Null2String(RSSJ!SJVo) & vbTab & Null2String(RSSJ!CRJ_AMOUNT)
                    End If
                End If
            Me.Caption = (RSSJ!INV_NO)
            DoEvents
            loadme.Value = loadme.Value + 1
            Label2.Caption = Round((loadme.Value / loadme.Max * 100), 0) & "%"
            RSSJ.MoveNext
        Loop
    End If
    If i > 0 Then thegrid.RemoveItem 1
    MsgBox "TAPOS na "
    Set RSCRJ = Nothing
    Set RSSJ = Nothing
End Sub
Function getcutomername(xxx As String)
    Dim rs As New ADODB.Recordset
    Set rs = gconDMIS.Execute("SELECT ACCTNAME from ALL_CUSTOMER_TABLE where cuscde ='" & xxx & "'")
    If Not (rs.EOF And rs.BOF) Then
            getcutomername = Null2String(rs!AcctName)
        Else
            getcutomername = ""
    End If
    Set rs = Nothing
End Function
Sub initthegrid()
    With thegrid
         .Cols = 7: .Rows = 2
        .DisplayFocusRect = True: .AllowUserResizing = True

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = "CRJ Vounherno"
        .Cell(0, 2).Text = "Customer Name"
        .Cell(0, 3).Text = "Invoice No"
        .Cell(0, 4).Text = "Invoice type"
        .Cell(0, 5).Text = "SJ voucherno"
        .Cell(0, 6).Text = "AMOUNT"
        

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox:    '.Column(2).MaxLength = 50
        .Column(3).CellType = cellTextBox:    '.Column(3).MaxLength = 50
        .Column(4).CellType = cellTextBox
        .Column(5).CellType = cellTextBox
        .Column(6).CellType = cellTextBox
       
       
        .Column(1).Width = 100: .Column(1).Locked = True
        .Column(2).Width = 295: .Column(2).Locked = True
        .Column(3).Width = 90: .Column(3).Locked = True
        .Column(4).Width = 80: .Column(4).Locked = True
        .Column(5).Width = 80: .Column(5).Locked = True
        .Column(6).Width = 200
       
        
        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 6, .Rows - 1, 6).ForeColor = RGB(0, 0, 128)
    End With
    With Grid2
         .Cols = 5: .Rows = 2
        .DisplayFocusRect = True: .AllowUserResizing = True

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = "Voucher No"
        .Cell(0, 2).Text = "Particular"
        .Cell(0, 3).Text = "Journard Date"
        .Cell(0, 4).Text = "AMOUNT"
        

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox:    '.Column(2).MaxLength = 50
        .Column(3).CellType = cellTextBox:    '.Column(3).MaxLength = 50
        .Column(4).CellType = cellTextBox
        
       
       
        .Column(1).Width = 100: .Column(1).Locked = True
        .Column(2).Width = 350: .Column(2).Locked = True
        .Column(3).Width = 90: .Column(3).Locked = True
        .Column(4).Width = 80: .Column(4).Locked = True
       
        
        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 4, .Rows - 1, 4).ForeColor = RGB(0, 0, 128)
    End With
End Sub



Sub InitGrid()
    With Grid1
         .Cols = 6: .Rows = 1
        .DisplayFocusRect = True: .AllowUserResizing = True

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = "Voucher#"
        .Cell(0, 2).Text = "Jtype"
        .Cell(0, 3).Text = "Sum of Debit"
        .Cell(0, 4).Text = "Sum of Credit"
        .Cell(0, 5).Text = "Remarks"
         
        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True
 
    End With
    
End Sub

Sub returnAcctDescription()
    Dim rs As New ADODB.Recordset
    Set rs = gconDMIS.Execute("select description,acctcode from AMIS_chartaccount where left(acctcode,5) in ('21-02','21-01')")
        cboacct.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            cboacct.AddItem rs!Description
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
End Sub
Sub displayGJ()
    Dim nard As New ADODB.Recordset
    Dim i As Integer
    Dim amount As Double
    Dim Total As Double
    Set nard = gconDMIS.Execute("SELECT  AMIS_Journal_HD.VoucherNo,AMIS_Journal_HD.remarks, AMIS_Journal_HD.JType,AMIS_Journal_HD.jdate,AMIS_Journal_HD.status,AMIS_Journal_HD.invoicetype,AMIS_Journal_HD.invoicedate, AMIS_Journal_HD.VendorCode,AMIS_Journal_HD.invoiceamt,AMIS_Journal_HD.amounttopay,AMIS_Journal_HD.amountpaid,AMIS_Journal_Det.Debit, AMIS_Journal_Det.CREDIT,AMIS_Journal_HD.DUEDATE,AMIS_Journal_Det.Acct_Code " & _
                                  "FROM AMIS_Journal_HD INNER JOIN AMIS_Journal_Det " & _
                                  "ON AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo AND AMIS_Journal_HD.jtype = AMIS_Journal_Det.jtype " & _
                                  "WHERE AMIS_Journal_HD.JType = 'GJ' and  AMIS_Journal_det.acct_Code ='" & getAccount(cboacct.Text) & "'")
    i = 0
    Grid2.Rows = 1
    If Not (nard.EOF And nard.BOF) Then
        Do While Not nard.EOF
                i = i + 1
                If nard!CREDIT = 0 Then
                    amount = nard!DEBIT
                    Else
                    amount = nard!CREDIT
                End If
                Grid2.AddItem nard![VOUCHERNO] & vbTab & nard!remarks & vbTab & _
                              nard!jdate & vbTab & amount
                Total = Total + amount
            nard.MoveNext
        Loop
    End If
    Texttotal = ToDoubleNumber(Total)
    If i > 1 Then Grid2.RemoveItem 1
    Grid2.Refresh:  Grid2.AutoRedraw = True
    Set nard = Nothing
End Sub
Function getAccount(xxx As String) As String
    Dim RSSSS As New ADODB.Recordset
    Set RSSSS = gconDMIS.Execute("select acctcode from AMIS_chartaccount where description='" & xxx & "'")
    If Not (RSSSS.EOF And RSSSS.BOF) Then
        getAccount = Null2String(RSSSS!acctcode)
    End If
    Set RSSSS = Nothing
End Function

Private Sub Grid1_DblClick()
    Dim VARVOUCHERNO                                   As String
    Dim VOUCHERNO                                      As String
    Dim RETURNVOUCHERNO                                As ADODB.Recordset
    VOUCHERNO = Grid1.Cell(Grid1.ActiveCell.Row, 1).Text
    JOURNALTYPE = Grid1.Cell(Grid1.ActiveCell.Row, 2).Text
    Set RETURNVOUCHERNO = gconDMIS.Execute("SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = " & N2Str2Null(VOUCHERNO) & " AND JTYPE=" & N2Str2Null(JOURNALTYPE))
    If Not RETURNVOUCHERNO.EOF And Not RETURNVOUCHERNO.BOF Then
        VARVOUCHERNO = Null2String(RETURNVOUCHERNO!VOUCHERNO)    'RIGHT(GRID1.TEXT, 6)
        Screen.MousePointer = 11
        If JOURNALTYPE = "COB" Then
            On Error Resume Next
            Unload frmAMISCustomerAROpening
            frmAMISCustomerAROpening.Show
            frmAMISCustomerAROpening.StoreSearch (VARVOUCHERNO)
        Else
            On Error Resume Next
            Unload frmAMISJournalEntry
            frmAMISJournalEntry.Show
            frmAMISJournalEntry.StoreSearch (VARVOUCHERNO)
        End If
        Screen.MousePointer = 0
    Else
    End If
End Sub

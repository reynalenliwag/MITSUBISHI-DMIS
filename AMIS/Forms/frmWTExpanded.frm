VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO50BF~1.OCX"
Begin VB.Form frmWTExpanded 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Withholding Tax - Expanded"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11760
   Begin VB.ComboBox cboATC 
      Height          =   315
      ItemData        =   "frmWTExpanded.frx":0000
      Left            =   4800
      List            =   "frmWTExpanded.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   90
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      Height          =   7785
      Left            =   0
      TabIndex        =   16
      Top             =   600
      Width           =   2715
      Begin XtremeReportControl.ReportControl rptEntity 
         Height          =   6825
         Left            =   60
         TabIndex        =   17
         Top             =   900
         Width           =   2595
         _Version        =   655364
         _ExtentX        =   4577
         _ExtentY        =   12039
         _StockProps     =   64
      End
      Begin VB.TextBox TextSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   2595
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Entity Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   150
         Width           =   2445
      End
   End
   Begin VB.ComboBox cboAR_ACCT_CODE 
      Height          =   315
      Left            =   1800
      Style           =   1  'Simple Combo
      TabIndex        =   9
      Text            =   "cboAR_ACCT_CODE"
      Top             =   90
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Height          =   435
      Left            =   11250
      Picture         =   "frmWTExpanded.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   60
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Prev"
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
      Left            =   8280
      MouseIcon       =   "frmWTExpanded.frx":0706
      MousePointer    =   99  'Custom
      Picture         =   "frmWTExpanded.frx":0858
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Move to Previous Record"
      Top             =   7590
      Width           =   705
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
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
      Left            =   8970
      MouseIcon       =   "frmWTExpanded.frx":0BB7
      MousePointer    =   99  'Custom
      Picture         =   "frmWTExpanded.frx":0D09
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Move to Next Record"
      Top             =   7590
      Width           =   705
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
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
      Left            =   9660
      MouseIcon       =   "frmWTExpanded.frx":1061
      MousePointer    =   99  'Custom
      Picture         =   "frmWTExpanded.frx":11B3
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Find a Record"
      Top             =   7590
      Width           =   705
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
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
      Left            =   10320
      MouseIcon       =   "frmWTExpanded.frx":14AD
      MousePointer    =   99  'Custom
      Picture         =   "frmWTExpanded.frx":15FF
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print this Record"
      Top             =   7590
      Width           =   705
   End
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
      Left            =   11040
      MouseIcon       =   "frmWTExpanded.frx":1965
      MousePointer    =   99  'Custom
      Picture         =   "frmWTExpanded.frx":1AB7
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit Window"
      Top             =   7590
      Width           =   705
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   2760
      ScaleHeight     =   765
      ScaleWidth      =   5445
      TabIndex        =   0
      Top             =   7590
      Width           =   5475
      Begin VB.Image Image1 
         Height          =   360
         Left            =   150
         Picture         =   "frmWTExpanded.frx":1E1D
         Top             =   180
         Width           =   360
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Double click the line item to open the corresponding journal entry of the particular voucher no."
         Height          =   435
         Left            =   630
         TabIndex        =   1
         Top             =   150
         Width           =   4725
      End
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   90
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97583105
      CurrentDate     =   40210
   End
   Begin MSComCtl2.DTPicker dtTO 
      Height          =   375
      Left            =   9720
      TabIndex        =   20
      Top             =   90
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97583105
      CurrentDate     =   40210
   End
   Begin VB.Frame Frame2 
      Height          =   6945
      Left            =   2760
      TabIndex        =   10
      Top             =   600
      Width           =   8985
      Begin XtremeReportControl.ReportControl rptLEDGER 
         Height          =   6495
         Left            =   0
         TabIndex        =   11
         Top             =   960
         Width           =   8925
         _Version        =   655364
         _ExtentX        =   15743
         _ExtentY        =   11456
         _StockProps     =   64
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1710
         TabIndex        =   13
         Top             =   150
         Width           =   1275
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1710
         TabIndex        =   12
         Top             =   570
         Width           =   7215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   15
         Top             =   210
         Width           =   1665
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   14
         Top             =   630
         Width           =   1635
      End
   End
   Begin Crystal.CrystalReport rptEWT 
      Left            =   12000
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "ATC Codes"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ATC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   4320
      TabIndex        =   25
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      TabIndex        =   22
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9360
      TabIndex        =   21
      Top             =   120
      Width           =   315
   End
End
Attribute VB_Name = "frmWTExpanded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLOAD_VENDOR                                           As ADODB.Recordset
Dim rsLOAD_AP                                               As ADODB.Recordset
Dim REC                                                     As XtremeReportControl.ReportRecord
Dim TOTAL_DEBIT                                             As Double
Dim TOTAL_CREDIT                                            As Double
Dim xBALANCE                                                As Double
Dim FWD_BALANCE                                             As Double
Dim FWD_DEBIT                                               As Double
Dim FWD_CREDIT                                              As Double
Dim rsREF                                                   As ADODB.Recordset
Dim xlApplication                                           As Excel.Application
Dim xlWorkbook                                              As Excel.Workbook
Dim xlWorksheet                                             As Excel.Worksheet
Dim xlRange                                                 As Excel.Range
Dim xCounter                                                As Integer

Sub INIT_CTRL_LEDGER()
    With rptLEDGER
        .Columns.DeleteAll
        .Columns.Add 0, "DOCDATE", 80, True: .Columns(0).Alignment = xtpAlignmentRight: .Columns(0).AllowRemove = False: .Columns(0).AutoSortWhenGrouped = True
        .Columns.Add 1, "REFERENCE", 110, True: .Columns(1).Alignment = xtpAlignmentCenter: .Columns(1).AllowRemove = False
        .Columns.Add 2, "INVOICE NO/CHECK NO", 110, True: .Columns(2).Alignment = xtpAlignmentCenter: .Columns(2).AllowRemove = False
        .Columns.Add 3, "DEBIT", 110, True: .Columns(3).Alignment = xtpAlignmentRight: .Columns(3).AllowRemove = False
        .Columns.Add 4, "CREDIT", 110, True: .Columns(4).Alignment = xtpAlignmentRight: .Columns(4).AllowRemove = False
        .Columns.Add 5, "BALANCE", 80, True: .Columns(5).Alignment = xtpAlignmentRight: .Columns(5).AllowRemove = False
        .Columns.Add 6, "ID", 0, True: .Columns(6).Alignment = xtpAlignmentIconRight: .Columns(6).AllowRemove = False: .Columns(6).Visible = False
        .Columns.Add 7, "JTYPE", 0, True: .Columns(7).Alignment = xtpAlignmentIconRight: .Columns(7).AllowRemove = False: .Columns(7).Visible = False

        .PaintManager.HorizontalGridStyle = xtpGridSolid    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSolid    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        .AllowColumnSort = False

        .ShowFooter = True

        .Columns(0).DrawFooterDivider = False
        .Columns(1).DrawFooterDivider = False
        .Columns(2).FooterText = "TOTAL : ": .Columns(2).FooterAlignment = xtpAlignmentCenter
        .Columns(3).FooterText = 0
        .Columns(4).FooterText = 0
        .Columns(5).FooterText = 0
        .Columns(3).FooterAlignment = xtpAlignmentRight
        .Columns(4).FooterAlignment = xtpAlignmentRight
        .Columns(5).FooterAlignment = xtpAlignmentRight
        .Columns(6).DrawFooterDivider = False
        .Columns(7).DrawFooterDivider = False
    End With
End Sub
Sub INIT_CTRL_ENTITY()
    With rptEntity
        .Columns.DeleteAll
        .Columns.Add 0, "ENTITY NAME", 150, True: .Columns(0).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False: .Columns(0).AutoSortWhenGrouped = True
        .Columns.Add 1, "ID", 0, True: .Columns(1).Alignment = xtpAlignmentCenter: .Columns(1).AllowRemove = False

        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    'xtpGridSolid    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSolid    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        .AllowColumnSort = False
    End With
End Sub

Private Sub cboATC_Change()
Dim rsATC_Code                      As ADODB.Recordset
Set rsATC_Code = New ADODB.Recordset

'rsATC_Code.Open "SELECT DISTINCT atc.NATURE  FROM AMIS_AP AP INNER JOIN AMIS_JOURNAL_DET DT ON LEFT(AP.VOUCHERNO,3)= DT.JTYPE AND RIGHT(AP.VOUCHERNO,6)= DT.VOUCHERNO INNER JOIN AMIS_CHARTACCOUNT AC ON DT.Acct_Code = AC.AcctCode INNER JOIN AMIS_ATC ATC ON ATC.ATC = DT.ATC WHERE AC.TranType1 = 'EXPANDED' AND AP.VENDOR_NAME = '" & txtName & "' ", gconDMIS, adOpenKeyset
rsATC_Code.Open "SELECT DISTINCT atc.NATURE  FROM AMIS_AP AP INNER JOIN AMIS_JOURNAL_DET DT ON LEFT(AP.VOUCHERNO,3)= DT.JTYPE AND RIGHT(AP.VOUCHERNO,6)= DT.VOUCHERNO INNER JOIN AMIS_CHARTACCOUNT AC ON DT.Acct_Code = AC.AcctCode INNER JOIN AMIS_ATC ATC ON ATC.ATC = DT.ATC WHERE AC.TranType1 = 'EXPANDED' AND AP.VENDOR_NAME = '" & txtName & "' ", gconDMIS, adOpenKeyset
cboATC.AddItem "ALL ATC"
If Not rsATC_Code.EOF And Not rsATC_Code.BOF Then
        Do While Not rsATC_Code.EOF
            cboATC.AddItem Null2String(rsATC_Code!NATURE)
            rsATC_Code.MoveNext
        Loop
    End If
    Set rsATC_Code = Nothing
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    TextSearch.SetFocus
    TextSearch.BackColor = &HFFFFC0
    TextSearch.Text = ""
End Sub

Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:
    rsREF.MoveNext
    If rsREF.EOF Then
        rsREF.MoveLast
        ShowLastRecordMsg
    End If
    Call StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:
    rsREF.MovePrevious
    If rsREF.BOF Then
        rsREF.MoveFirst
        ShowFirstRecordMsg
    End If
    Call StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Print", "WITHHOLDING TAX EXPANDED") = False Then Exit Sub
    
    Screen.MousePointer = 11
    rptEWT.Reset
    rptEWT.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    rptEWT.Formulas(2) = "CompanyAddress = '" & UCase(COMPANY_ADDRESS) & "'"
    rptEWT.Formulas(3) = "TIN = '" & COMPANY_TIN & "'"
    rptEWT.Formulas(4) = "FromJdate = '" & dtFrom & "'"
    rptEWT.Formulas(5) = "ToJdate = '" & dtTO & "'"
    rptEWT.Formulas(6) = "FromMonth = '" & Format(Month(CDate(dtFrom)), "00") & "'"
    rptEWT.Formulas(7) = "FromDate = '" & Format(Day(CDate(dtFrom)), "00") & "'"
    rptEWT.Formulas(8) = "FromYear = '" & Right(Format(Year(CDate(dtFrom)), "00"), 2) & "'"
    rptEWT.Formulas(9) = "ToMonth = '" & Format(Month(CDate(dtTO)), "00") & "'"
    rptEWT.Formulas(10) = "ToDate = '" & Format(Day(CDate(dtTO)), "00") & "'"
    rptEWT.Formulas(11) = "ToYear = '" & Right(Format(Year(CDate(dtTO)), "00"), 2) & "'"
    rptEWT.Formulas(12) = "PAYORS = '" & LOGNAME & "'"
    
    If cboATC.Text = "ALL ATC" Then
        PrintSQLReport rptEWT, AMIS_REPORT_PATH & "\files\WithholdingTax.rpt", "{Journal_HD.VendorCode} = '" & txtCode & "' and {Journal_HD.JDate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_HD.JDate} <= date(" & Year(dtTO) & "," & Month(dtTO) & "," & Day(dtTO) & ")", DMIS_REPORT_Connection, 1
    Else
        PrintSQLReport rptEWT, AMIS_REPORT_PATH & "\files\WithholdingTax.rpt", "{Journal_HD.VendorCode} = '" & txtCode & "' and {Journal_HD.JDate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_HD.JDate} <= date(" & Year(dtTO) & "," & Month(dtTO) & "," & Day(dtTO) & ") and {AMIS_ATC.NATURE} = '" & cboATC.Text & "'", DMIS_REPORT_Connection, 1
    End If
    Screen.MousePointer = 0

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdOK_Click()
    If cboAR_ACCT_CODE.Text = "" Then
        MessagePop InfoFriend, "SYSTEM MESSAGE", "Please select the account code."
        cboAR_ACCT_CODE.SetFocus
        Exit Sub
    End If

    If CDate(dtFrom.Value) > CDate(dtTO.Value) Then
        MessagePop InfoFriend, "System Message", "Invalid date range Date from is greater than date to."
        Exit Sub
    End If

    If txtCode.Text = "" Then
        MessagePop InfoFriend, "System Message", "Entity name not yet selected.Please select entity name."
        Exit Sub
    End If
    
    If cboATC.Text = "" Then
        MsgBox "Select ATC.", vbInformation, "ATC CODE"
        cboATC.Enabled = True
        cboATC.SetFocus
        Exit Sub
    End If
  
    Call LOAD_AP_ENTITY
End Sub

Function Setacctcode(XXX As String) As String
    Dim rsSetAcctCode                                       As ADODB.Recordset
    Set rsSetAcctCode = New ADODB.Recordset
    rsSetAcctCode.Open "SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "'", gconDMIS, adOpenForwardOnly
    If Not rsSetAcctCode.EOF And Not rsSetAcctCode.BOF Then
        Setacctcode = rsSetAcctCode!AcctCode
    End If
    Set rsSetAcctCode = Nothing
End Function



Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Call INIT_CTRL_ENTITY
    Call INIT_CTRL_LEDGER
    Call LOAD_ACCT_CODE
    Call LOAD_MAX_MIN_DATE
    Call LOAD_VENDOR
    cboAR_ACCT_CODE.ListIndex = 0
    Call rsRefresh
    Call StoreMemVars
    INIT_CBO_ATC
End Sub

Sub LOAD_ACCT_CODE()
    Dim rsLOAD_ACCT_CODE                                    As ADODB.Recordset
    Set rsLOAD_ACCT_CODE = New ADODB.Recordset
    rsLOAD_ACCT_CODE.Open "SELECT UPPER(DESCRIPTION) AS DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 = 'EXPANDED' ", gconDMIS, adOpenKeyset
    If Not rsLOAD_ACCT_CODE.EOF And Not rsLOAD_ACCT_CODE.BOF Then
        Do While Not rsLOAD_ACCT_CODE.EOF
            cboAR_ACCT_CODE.AddItem UCase(Null2String(rsLOAD_ACCT_CODE!DESCRIPTION))
            rsLOAD_ACCT_CODE.MoveNext
        Loop
    End If
    Set rsLOAD_ACCT_CODE = Nothing
End Sub

Sub LOAD_MAX_MIN_DATE()
    Dim rsMAX_MIN_DATE                                      As ADODB.Recordset
    Set rsMAX_MIN_DATE = New ADODB.Recordset
    rsMAX_MIN_DATE.Open "SELECT * FROM (SELECT MAX(JDATE)AS MAX_JDATE, MIN(JDATE) AS MIN_JDATE FROM AMIS_JOURNAL_HD WHERE STATUS='P') T WHERE MAX_JDATE IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsMAX_MIN_DATE.EOF And Not rsMAX_MIN_DATE.BOF Then
        dtFrom.Value = Null2String(rsMAX_MIN_DATE!MIN_JDATE)
        dtTO.Value = Null2String(rsMAX_MIN_DATE!MAX_JDATE)
    Else
        dtFrom.Value = LOGDATE
        dtTO.Value = LOGDATE
        MessagePop InfoFriend, "Info", "No such Record!"
    End If
    Set rsMAX_MIN_DATE = Nothing
End Sub

Private Sub rptENTITY_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    Metrics.BackColor = RGB(214, 234, 246)
End Sub
Private Sub rptENTITY_SelectionChanged()
    txtCode.Text = UCase(rptEntity.SelectedRows(0).Record(1).Value)
    txtName.Text = UCase(rptEntity.SelectedRows(0).Record(0).Value)
    INIT_CBO_ATC
    LOAD_MAX_MIN_DATE
    rptLEDGER.Records.DeleteAll
    rptLEDGER.Populate
End Sub

Sub LOAD_VENDOR()
    Set rsLOAD_VENDOR = New ADODB.Recordset
    'rsLOAD_VENDOR.Open "SELECT DISTINCT TOP 22 AP.Vendor_code,AP.VENDOR_NAME  FROM AMIS_AP AP INNER JOIN AMIS_JOURNAL_DET DT ON LEFT(AP.VOUCHERNO,3)= DT.JTYPE AND RIGHT(AP.VOUCHERNO,6)= DT.VOUCHERNO INNER JOIN AMIS_CHARTACCOUNT AC ON DT.Acct_Code = AC.AcctCode WHERE AC.TranType1 = 'EXPANDED' AND DT.STATUS = 'P' ORDER BY AP.VENDOR_NAME ASC", gconDMIS, adOpenKeyset
    rsLOAD_VENDOR.Open "SELECT DISTINCT TOP 22 AE.CODE,AE.ACCOUNTNAME  FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DT ON HD.JType = DT.JType AND HD.VoucherNo = DT.VoucherNo INNER JOIN ALL_ENTITY AE ON  HD.VendorCode =AE.CODE AND HD.ENTITY_CLASS = AE.ENTITYCODE INNER JOIN AMIS_CHARTACCOUNT AC ON DT.Acct_Code = AC.AcctCode WHERE AC.TranType1 = 'EXPANDED' AND DT.Status = 'P' ORDER BY AE.ACCOUNTNAME ASC", gconDMIS, adOpenKeyset
    rptEntity.Records.DeleteAll

    If Not rsLOAD_VENDOR.EOF And Not rsLOAD_VENDOR.BOF Then
        Do While Not rsLOAD_VENDOR.EOF
            Set REC = rptEntity.Records.Add
            REC.AddItem (Trim(UCase(Null2String(rsLOAD_VENDOR!ACCOUNTNAME))))
            REC.AddItem (Trim(UCase(Null2String(rsLOAD_VENDOR!Code))))
            rptEntity.Populate
            Set REC = Nothing
            rsLOAD_VENDOR.MoveNext
        Loop
    End If
    Set rsLOAD_VENDOR = Nothing
End Sub

Private Sub rptLEDGER_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record(3).Value <> "0.00" And RTrim(LTrim(Row.Record(2).Value)) <> "TOTAL" Then
        Metrics.BackColor = vbWhite
    Else
        Metrics.BackColor = RGB(214, 234, 246)
    End If
End Sub

Private Sub rptLEDGER_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim xJOURNAL                                            As String
    Dim xVOUCHER                                            As String
    If Len(Row.Record(1).Value) = 10 Then
        xJOURNAL = Left((Row.Record(1).Value), 3)
        xVOUCHER = Right((Row.Record(1).Value), 6)
    Else
        xJOURNAL = Left((Row.Record(1).Value), 2)
        xVOUCHER = Right((Row.Record(1).Value), 6)
    End If

    If xJOURNAL = "APJ" Then
        Call frmAMISJournalEntry_APJ.LOADJOURNAL("APJ")
        FormExistsShow frmAMISJournalEntry_APJ
        Call frmAMISJournalEntry_APJ.SearchVoucherNo(xVOUCHER)
    ElseIf xJOURNAL = "CDJ" Then
        Call frmAMISJournalEntry_CDJ.LOADJOURNAL("CDJ")
        FormExistsShow frmAMISJournalEntry_CDJ
        Call frmAMISJournalEntry_CDJ.SearchVoucherNo(xVOUCHER)
    ElseIf xJOURNAL = "SJ" Then
        Call frmAMISJournalEntry_SJ.LOADJOURNAL("SJ")
        FormExistsShow frmAMISJournalEntry_SJ
        Call frmAMISJournalEntry_SJ.SearchVoucherNo(xVOUCHER)
    ElseIf xJOURNAL = "CRJ" Then
        Call frmAMISJournalEntry_CRJ.LOADJOURNAL("CRJ")
        FormExistsShow frmAMISJournalEntry_CRJ
        Call frmAMISJournalEntry_CRJ.SearchVoucherNo(xVOUCHER)
    ElseIf xJOURNAL = "GJ" Then
        Call frmAMISJournalEntry_GJ.LOADJOURNAL("GJ")
        FormExistsShow frmAMISJournalEntry_GJ
        Call frmAMISJournalEntry_GJ.SearchVoucherNo(xVOUCHER)
    ElseIf xJOURNAL = "VPJ" Then
        Call frmAMISVendorAPOpening.LOADJOURNAL("VPJ")
        FormExistsShow frmAMISVendorAPOpening
        Call frmAMISVendorAPOpening.SearchVoucherNo(xVOUCHER)
    ElseIf xJOURNAL = "CDM" Then
        Call frmAMISJournalEntry_CDM.LOADJOURNAL("CDM")
        FormExistsShow frmAMISJournalEntry_CDM
        Call frmAMISJournalEntry_CDM.SearchVoucherNo(xVOUCHER)
    ElseIf xJOURNAL = "CCM" Then
        Call frmAMISJournalEntry_CCM.LOADJOURNAL("CCM")
        FormExistsShow frmAMISJournalEntry_CCM
        Call frmAMISJournalEntry_CCM.SearchVoucherNo(xVOUCHER)
    ElseIf xJOURNAL = "VCM" Then
        Call frmAMISJournalEntry_VCM.LOADJOURNAL("VCM")
        FormExistsShow frmAMISJournalEntry_VCM
        Call frmAMISJournalEntry_VCM.SearchVoucherNo(xVOUCHER)
    ElseIf xJOURNAL = "VDM" Then
        Call frmAMISJournalEntry_VDM.LOADJOURNAL("VDM")
        FormExistsShow frmAMISJournalEntry_VDM
        Call frmAMISJournalEntry_VDM.SearchVoucherNo(xVOUCHER)
    End If
End Sub

Private Sub textSearch_Change()
    Dim rssearch                                            As ADODB.Recordset
    Set rssearch = New ADODB.Recordset
    If TextSearch.Text <> "" Then
        rssearch.Open "SELECT DISTINCT VENDOR_CODE,VENDOR_NAME FROM AMIS_AP WHERE VENDOR_NAME LIKE '" & Replace(TextSearch.Text, "'", "") & "%' ORDER BY VENDOR_NAME ASC", gconDMIS, adOpenKeyset
    Else
        rssearch.Open "SELECT DISTINCT TOP 22 VENDOR_CODE,VENDOR_NAME FROM AMIS_AP ORDER BY VENDOR_NAME ASC", gconDMIS, adOpenKeyset
    End If

    rptEntity.Records.DeleteAll
    If Not rssearch.EOF And Not rssearch.BOF Then
        Do While Not rssearch.EOF
            Set REC = rptEntity.Records.Add
            REC.AddItem (Trim(UCase(Null2String(rssearch!VENDOR_NAME))))
            REC.AddItem (Trim(UCase(Null2String(rssearch!VENDOR_CODE))))
            rptEntity.Populate
            Set REC = Nothing
            rssearch.MoveNext
        Loop
    End If
    Set rssearch = Nothing
End Sub

Sub LOAD_AP_ENTITY()
    Dim xVOUCHERNO                                          As String
    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0: xBALANCE = 0
    Set rsLOAD_AP = New ADODB.Recordset
    If cboATC.Text = "ALL ATC" Then
'        rsLOAD_AP.Open "SELECT HD.JDate,HD.JTYPE+'-'+HD.VoucherNo AS APVOUCHERNO,AP.INVOICENO,DT.DEBIT,DT.CREDIT,DT.ACCT_CODE,DT.ATC,ATC.NATURE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_AP AP ON LEFT(AP.Voucherno,3) = HD.JTYPE AND RIGHT(AP.Voucherno,6)  = HD.VOUCHERNO INNER JOIN AMIS_JOURNAL_DET DT ON HD.JType = DT.JType AND HD.VoucherNo = DT.VoucherNo INNER JOIN AMIS_ATC ATC ON DT.ATC = ATC.ATC WHERE HD.STATUS='P' AND HD.VENDORCODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND HD.Jdate >= '" & dtFrom.Value & "'and HD.Jdate <= '" & dtTO.Value & "' AND " & _
'                       "DT.ACCT_CODE = '" & GET_ACCTCODE((RTrim(LTrim(cboAR_ACCT_CODE.Text)))) & "' ORDER BY APVOUCHERNO", gconDMIS, adOpenKeyset
        rsLOAD_AP.Open "SELECT HD.JDate,HD.JType+'-'+HD.VoucherNo AS APVOUCHERNO,HD.InvoiceNo,DT.DEBIT,DT.CREDIT,DT.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DT ON HD.JType = DT.JType AND HD.VoucherNo = DT.VoucherNo INNER JOIN AMIS_ATC ATC ON DT.ATC = ATC.ATC WHERE HD.STATUS='P' AND HD.VENDORCODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND HD.Jdate >= '" & dtFrom.Value & "'and HD.Jdate <= '" & dtTO.Value & "' AND " & _
                       "DT.ACCT_CODE = '" & GET_ACCTCODE((RTrim(LTrim(cboAR_ACCT_CODE.Text)))) & "'ORDER BY APVOUCHERNO", gconDMIS, adOpenKeyset
      
      Else
'        rsLOAD_AP.Open "SELECT HD.JDate,HD.JTYPE+'-'+HD.VoucherNo AS APVOUCHERNO,AP.INVOICENO,DT.DEBIT,DT.CREDIT,DT.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_AP AP ON LEFT(AP.Voucherno,3) = HD.JTYPE AND RIGHT(AP.Voucherno,6)  = HD.VOUCHERNO INNER JOIN AMIS_JOURNAL_DET DT ON HD.JType = DT.JType AND HD.VoucherNo = DT.VoucherNo INNER JOIN AMIS_ATC ATC ON DT.ATC = ATC.ATC WHERE HD.STATUS='P' AND HD.VENDORCODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND HD.Jdate >= '" & dtFrom.Value & "'and HD.Jdate <= '" & dtTo.Value & "' AND " & _
'                       "DT.ACCT_CODE = '" & GET_ACCTCODE((RTrim(LTrim(cboAR_ACCT_CODE.Text)))) & "' AND ATC.NATURE = '" & (cboATC.Text) & "' ORDER BY APVOUCHERNO", gconDMIS, adOpenKeyset
        rsLOAD_AP.Open "SELECT HD.JDate,HD.JType+'-'+HD.VoucherNo AS APVOUCHERNO,HD.InvoiceNo,DT.DEBIT,DT.CREDIT,DT.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DT ON HD.JType = DT.JType AND HD.VoucherNo = DT.VoucherNo INNER JOIN AMIS_ATC ATC ON DT.ATC = ATC.ATC WHERE HD.STATUS='P' AND HD.VENDORCODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND HD.Jdate >= '" & dtFrom.Value & "'and HD.Jdate <= '" & dtTO.Value & "' AND " & _
                       "DT.ACCT_CODE = '" & GET_ACCTCODE((RTrim(LTrim(cboAR_ACCT_CODE.Text)))) & "' AND ATC.NATURE = '" & (cboATC.Text) & "' ORDER BY APVOUCHERNO", gconDMIS, adOpenKeyset
    
    End If

    rptLEDGER.Records.DeleteAll
    Screen.MousePointer = 11
    Call FORWARDED_BALANCE

    Set REC = rptLEDGER.Records.Add
    REC.AddItem (Trim(dtFrom.Value))
    REC.AddItem (Trim("FWD BALANCE"))
    REC.AddItem (Trim(""))
    REC.AddItem (Trim("0.00"))
    REC.AddItem (Trim("0.00"))
    REC.AddItem (Trim(ToDoubleNumber(FWD_BALANCE)))
    rptLEDGER.Populate
    Set REC = Nothing

    If Not rsLOAD_AP.EOF And Not rsLOAD_AP.BOF Then
        Do While Not rsLOAD_AP.EOF
            If Len(rsLOAD_AP!APVoucherno) = 10 Then
                xVOUCHERNO = Left(Null2String(rsLOAD_AP!APVoucherno), 3)
            Else
                xVOUCHERNO = Left(Null2String(rsLOAD_AP!APVoucherno), 2)
            End If
            If xVOUCHERNO = "VPJ" Or xVOUCHERNO = "APJ" Or xVOUCHERNO = "CDJ" Or xVOUCHERNO = "GJ" Or xVOUCHERNO = "SJ" Or xVOUCHERNO = "CRJ" Or xVOUCHERNO = "CDM" Or xVOUCHERNO = "CCM" Or xVOUCHERNO = "VDM" Or xVOUCHERNO = "VCM" Then
                Set REC = rptLEDGER.Records.Add
                REC.AddItem (Trim(Null2String(rsLOAD_AP!JDATE)))
                REC.AddItem (Trim(Null2String(rsLOAD_AP!APVoucherno)))
                REC.AddItem (Trim(Null2String(rsLOAD_AP!INVOICENO)))
                If NumericVal(rsLOAD_AP!Credit) <> 0 Then
                    REC.AddItem (Trim("0.00"))
                    REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsLOAD_AP!Credit))))

                    TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_AP!Credit)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_AP!Credit)), 2))

                    REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
                Else
                    REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsLOAD_AP!Debit))))
                    REC.AddItem (Trim("0.00"))

                    TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_AP!Debit)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_AP!Debit)), 2))

                    REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
                End If

                Call REFERENCE_VOUCHER(Null2String(rsLOAD_AP!APVoucherno), txtCode.Text, Null2String(rsLOAD_AP!ACCT_CODE), Null2String(rsLOAD_AP!INVOICENO))
                Call ADJUSTMENT_BYVOUHCHERNO(Null2String(Right(rsLOAD_AP!APVoucherno, 6)), xVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AP!ACCT_CODE))
                rptLEDGER.Populate
                Set REC = Nothing
                rptLEDGER.Populate
                Set REC = Nothing
            End If
            rsLOAD_AP.MoveNext
            DoEvents
        Loop
    End If

    rptLEDGER.Columns(3).FooterText = ToDoubleNumber(TOTAL_DEBIT)
    rptLEDGER.Columns(4).FooterText = ToDoubleNumber(TOTAL_CREDIT)
    rptLEDGER.Columns(5).FooterText = ToDoubleNumber(Round(NumericVal(xBALANCE + FWD_BALANCE), 2))
    Screen.MousePointer = 0
    'Set rsLOAD_AP = Nothing
End Sub



Sub REFERENCE_VOUCHER(xVOUCHERNO As String, xVENDORCODE As String, xAcctCode As String, xInvoiceNo As String)
    Dim rsPAYMENT                                           As ADODB.Recordset
    Set rsPAYMENT = New ADODB.Recordset
    rsPAYMENT.Open "SELECT JDATE,JTYPE,VOUCHERNO,ISNULL(INVOICETYPE,'') AS INVOICETYPE,ISNULL(INVOICENO,'') AS INVOICENO,AMOUNTPAID FROM AMIS_DETAILS WHERE STATUS='P' AND INVOICENO='" & xInvoiceNo & "' AND PV_VOUCHERNO = '" & xVOUCHERNO & "' AND VENDORCODE = '" & xVENDORCODE & "' AND ACCT_CODE = '" & xAcctCode & "' " & _
                   "AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "'", gconDMIS, adOpenKeyset
    If Not rsPAYMENT.EOF And Not rsPAYMENT.BOF Then
        Do While Not rsPAYMENT.EOF
            Set REC = rptLEDGER.Records.Add
            REC.AddItem (Trim(Null2String(rsPAYMENT!JDATE)))
            REC.AddItem (Trim(Null2String(rsPAYMENT!JTYPE) & "-" & Null2String(rsPAYMENT!VOUCHERNO)))
            REC.AddItem (Trim(Null2String(rsPAYMENT!INVOICENO)))
            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsPAYMENT!Debit))))
            REC.AddItem (Trim("0.00"))

            TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsPAYMENT!Debit)), 2))
            xBALANCE = Trim(ToDoubleNumber(xBALANCE - NumericVal(rsPAYMENT!Debit)))

            REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
            rptLEDGER.Populate
            Set REC = Nothing
            rsPAYMENT.MoveNext
            DoEvents
        Loop
    End If
    Set rsPAYMENT = Nothing
End Sub


Sub REFERENCE_VOUCHER_PRINTING(xVOUCHERNO As String, xVENDORCODE As String, xAcctCode As String, xInvoiceNo As String)
    Dim rsPAYMENT                                           As ADODB.Recordset
    Set rsPAYMENT = New ADODB.Recordset
    rsPAYMENT.Open "SELECT JDATE,JTYPE,VOUCHERNO,ISNULL(INVOICETYPE,'') + CASE WHEN INVOICETYPE IS NULL THEN '' ELSE '-' END + ISNULL(INVOICENO,'') AS INVOICENO,AMOUNTPAID FROM AMIS_DETAILS WHERE STATUS='P' AND INVOICENO = '" & xInvoiceNo & "' AND PV_VOUCHERNO = '" & xVOUCHERNO & "' AND VENDORCODE = '" & xVENDORCODE & "' AND ACCT_CODE = '" & xAcctCode & "' " & _
                   "AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "'", gconDMIS, adOpenKeyset
    If Not rsPAYMENT.EOF And Not rsPAYMENT.BOF Then
        Do While Not rsPAYMENT.EOF
        
            xCounter = xCounter + 1
            xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsPAYMENT!JDATE))), "mm/dd/yyyy")
            xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsPAYMENT!JTYPE) & "-" & Null2String(rsPAYMENT!VOUCHERNO)))
            xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsPAYMENT!INVOICENO)))
            xlWorksheet.Cells(xCounter, "D") = (Trim(ToDoubleNumber(NumericVal(rsPAYMENT!AMOUNTPAID))))
            xlWorksheet.Cells(xCounter, "E") = (Trim("0.00"))

            TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsPAYMENT!AMOUNTPAID)), 2))
            xBALANCE = Trim(ToDoubleNumber(xBALANCE - NumericVal(rsPAYMENT!AMOUNTPAID)))

            xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
            Set REC = Nothing
            rsPAYMENT.MoveNext
            DoEvents
        Loop
    End If
    Set rsPAYMENT = Nothing
End Sub



Sub ADJUSTMENT_BYVOUHCHERNO(xADJVOUCHERNO, xADJTYPE, xVENDORCODE As String, xACCT_CODE As String)
    Dim rsADJ                                               As ADODB.Recordset
    Set rsADJ = New ADODB.Recordset
    rsADJ.Open "SELECT JDATE, VOUCHERNO, INVOICETYPE + '-' + INVOICENO AS REF_INVOICE,AMOUNT2PAY,AMOUNTPAID " & _
               "FROM AMIS_AP WHERE STATUS='P' AND VENDOR_CODE = '" & xVENDORCODE & "' AND " & _
               "Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' AND INVOICENO='" & xADJVOUCHERNO & "' AND INVOICETYPE='" & xADJTYPE & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND LEFT(VOUCHERNO,2) = 'GJ'", gconDMIS, adOpenKeyset
    If Not rsADJ.EOF And Not rsADJ.BOF Then
        Do While Not rsADJ.EOF
            Set REC = rptLEDGER.Records.Add
            REC.AddItem (Trim(Null2String(rsADJ!JDATE)))
            REC.AddItem (Trim(Null2String(rsADJ!VOUCHERNO)))
            REC.AddItem (Trim(Null2String(rsADJ!REF_INVOICE)))

            If NumericVal(rsADJ!AMOUNT2PAY) <> 0 Then
                REC.AddItem (Trim("0.00"))
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsADJ!AMOUNT2PAY))))


                TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsADJ!AMOUNT2PAY)), 2))
                xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsADJ!AMOUNT2PAY)), 2))

                REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
            Else
                REC.AddItem (Trim(Null2String(rsADJ!AMOUNT_PAID)))
                REC.AddItem (Trim("0.00"))

                TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsADJ!AMOUNT_PAID)), 2))
                xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsADJ!AMOUNT_PAID)), 2))

                REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
            End If

            rptLEDGER.Populate
            Set REC = Nothing
            rsADJ.MoveNext
        Loop
    End If
    Set rsADJ = Nothing
End Sub

Sub ADJUSTMENT_BYVOUHCHERNO_PRINTING(xVOUCHERNO As String, xJType As String, xCUSCDE As String, xACCT_CODE As String)
    Dim rsADJ                                               As ADODB.Recordset
    Set rsADJ = New ADODB.Recordset
    rsADJ.Open "SELECT JDATE, SJVOUCHERNO, INVOICETYPE + '-' + INVOICENO AS REF_INVOICE,AMOUNT_TOPAY,AMOUNT_PAID " & _
               "FROM AMIS_AP WHERE STATUS='P' AND VENDOR_CODE = '" & xCUSCDE & "' AND INVOICETYPE = '" & xJType & "' AND " & _
               "Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' AND  INVOICENO = '" & xVOUCHERNO & "' AND ACCOUNT_CODE = '" & xACCT_CODE & "' AND LEFT(SJVOUCHERNO,2) = 'GJ'", gconDMIS, adOpenKeyset
    If Not rsADJ.EOF And Not rsADJ.BOF Then
        Do While Not rsADJ.EOF
            xCounter = xCounter + 1
            xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsADJ!JDATE))), "mm/dd/yyyy")
            xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsADJ!SJVoucherno)))
            xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsADJ!REF_INVOICE)))
            If NumericVal(rsADJ!AMOUNT_TOPAY) <> 0 Then
                xlWorksheet.Cells(xCounter, "D") = (Trim(ToDoubleNumber(NumericVal(rsADJ!AMOUNT_TOPAY))))
                xlWorksheet.Cells(xCounter, "E") = (Trim("0.00"))
                TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsADJ!AMOUNT_TOPAY)), 2))
                xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsADJ!AMOUNT_TOPAY)), 2))
                xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
            Else
                xlWorksheet.Cells(xCounter, "D") = (Trim("0.00"))
                xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsADJ!AMOUNT_PAID)))
                TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsADJ!AMOUNT_PAID)), 2))
                xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsADJ!AMOUNT_PAID)), 2))
                xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
            End If
            rsADJ.MoveNext
        Loop
    End If
    Set rsADJ = Nothing
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        rptEntity.SetFocus
    End If
End Sub

Private Sub TextSearch_LostFocus()
    TextSearch.BackColor = vbWhite
End Sub

Sub FORWARDED_BALANCE()
    Dim rsLOAD_AP                                           As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Set rsLOAD_AP = New ADODB.Recordset

    FWD_BALANCE = 0: FWD_CREDIT = 0: FWD_DEBIT = 0

    If cboAR_ACCT_CODE.Text = "" Then
        rsLOAD_AP.Open "SELECT JDATE,VOUCHERNO AS APVOUCHERNO,INVOICETYPE + '-' + INVOICENO AS INVOICE,AMOUNT2PAY,AMOUNTPAID,INVOICENO,INVOICETYPE,VENDOR_CODE,ACCT_CODE,RIGHT(VOUCHERNO,6) AS VOUCHERNO " & _
                       "FROM AMIS_AP WHERE STATUS='P' AND VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate < '" & dtFrom.Value & "'", gconDMIS, adOpenKeyset
    Else
        rsLOAD_AP.Open "SELECT HD.JDate,HD.VoucherNo AS APVOUCHERNO,HD.INVOICETYPE,HD.InvoiceNo,DT.Debit,DT.Credit,HD.VendorCode,DT.Acct_Code,HD.VoucherNo FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DT ON HD.JType = DT.JType AND HD.VoucherNo = DT.VoucherNo " & _
                       "WHERE HD.STATUS='P' AND HD.VENDORCODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND HD.Jdate < '" & dtFrom.Value & "' AND " & _
                       "DT.ACCT_CODE = (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "')", gconDMIS, adOpenKeyset
    End If

    If Not rsLOAD_AP.EOF And Not rsLOAD_AP.BOF Then
        Do While Not rsLOAD_AP.EOF
            If Len(rsLOAD_AP!APVoucherno) = 10 Then
                xVOUCHERNO = Left(Null2String(rsLOAD_AP!APVoucherno), 3)
            Else
                xVOUCHERNO = Left(Null2String(rsLOAD_AP!APVoucherno), 2)
            End If
            If xVOUCHERNO = "VPJ" Or xVOUCHERNO = "APJ" Then
                If NumericVal(rsLOAD_AP!AMOUNT2PAY) <> 0 Then
                    FWD_CREDIT = ToDoubleNumber(Round((FWD_CREDIT + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))
                    FWD_BALANCE = ToDoubleNumber(Round((FWD_BALANCE + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))
                Else
                    FWD_DEBIT = ToDoubleNumber(Round((FWD_DEBIT + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))
                    FWD_BALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_AP!AMOUNTPAID)), 2))
                End If
            End If
            rsLOAD_AP.MoveNext
        Loop
    End If
End Sub

Sub FWD_REFERENCE_INVOICE(xInvoiceNo As String, xInvoiceType As String, xCUSCODE As String, xAcctCode As String)
    Dim rsINVOICE                                           As ADODB.Recordset
    Set rsINVOICE = New ADODB.Recordset
    rsINVOICE.Open "SELECT * FROM AMIS_DETAIL WHERE STATUS='P' AND INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "' AND VENDOR_CODE = '" & xCUSCODE & "' AND ACCT_CODE = '" & xAcctCode & "' " & _
                   "AND Jdate < '" & dtFrom.Value & "'", gconDMIS, adOpenKeyset
    If Not rsINVOICE.EOF And Not rsINVOICE.BOF Then
        Do While Not rsINVOICE.EOF
            FWD_CREDIT = ToDoubleNumber(Round((FWD_CREDIT + NumericVal(rsINVOICE!invoiceamount)), 2))
            FWD_BALANCE = Trim(ToDoubleNumber(FWD_BALANCE - NumericVal(rsINVOICE!invoiceamount)))
            rsINVOICE.MoveNext
        Loop
        Call FWD_ADJUSTMENT_BYVOUHCHERNO(xInvoiceNo, xInvoiceType, txtCode.Text, xAcctCode)
    End If
    Set rsINVOICE = Nothing
End Sub

Sub FWD_ADJUSTMENT_BYVOUHCHERNO(xVOUCHERNO As String, xJType As String, xCUSCDE As String, xACCT_CODE As String)
    Dim rsADJ                                               As ADODB.Recordset
    Set rsADJ = New ADODB.Recordset
    rsADJ.Open "SELECT JDATE, SJVOUCHERNO, INVOICETYPE + '-' + INVOICENO AS REF_INVOICE,AMOUNT_TOPAY,AMOUNT_PAID " & _
               "FROM AMIS_AP WHERE STATUS='P' AND VENDOR_CODE = '" & xCUSCDE & "' AND INVOICETYPE = '" & xJType & "' AND " & _
               "Jdate < '" & dtFrom.Value & "' AND  INVOICENO = '" & xVOUCHERNO & "' AND ACCOUNT_CODE = '" & xACCT_CODE & "' AND LEFT(SJVOUCHERNO,2) = 'GJ'", gconDMIS, adOpenKeyset
    If Not rsADJ.EOF And Not rsADJ.BOF Then
        Do While Not rsADJ.EOF
            If NumericVal(rsADJ!AMOUNT_TOPAY) <> 0 Then
                FWD_DEBIT = ToDoubleNumber(Round((FWD_DEBIT + NumericVal(rsADJ!AMOUNT_TOPAY)), 2))
                FWD_BALANCE = ToDoubleNumber(Round((FWD_BALANCE + NumericVal(rsADJ!AMOUNT_TOPAY)), 2))
            Else
                FWD_CREDIT = ToDoubleNumber(Round((FWD_CREDIT + NumericVal(rsADJ!AMOUNT_PAID)), 2))
                FWD_BALANCE = ToDoubleNumber(Round((FWD_BALANCE - NumericVal(rsADJ!AMOUNT_PAID)), 2))
            End If
            rsADJ.MoveNext
        Loop
    End If
    Set rsADJ = Nothing
End Sub

Sub rsRefresh()
    Set rsREF = New ADODB.Recordset
    rsREF.Open "SELECT DISTINCT VENDORCODE FROM AMIS_JOURNAL_HD WHERE STATUS='P' ORDER BY VENDORCODE ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsREF.EOF And Not rsREF.BOF Then
        txtCode.Text = UCase(Null2String(rsREF!VendorCode))
    End If
    Call LOAD_AP_ENTITY
End Sub

Function GetInvoices(VOUCHERNO As String, JTYPE As String) As String
    Dim rsInvoices                                          As ADODB.Recordset
    Set rsInvoices = New ADODB.Recordset
    rsInvoices.Open "SELECT INV_NO FROM AMIS_PV_DETAIL WHERE STATUS='P' AND VOUCHERNO = '" & VOUCHERNO & "' AND JTYPE='" & JTYPE & "'", gconDMIS, adOpenForwardOnly
    If Not rsInvoices.EOF And Not rsInvoices.BOF Then
        Do While Not rsInvoices.EOF
            GetInvoices = GetInvoices + "," + Null2String(rsInvoices!INV_NO)
            rsInvoices.MoveNext
        Loop
    End If
    GetInvoices = Mid(GetInvoices, 2, Len(GetInvoices))
    Set rsInvoices = Nothing
End Function

Function GET_ACCTCODE(xACCT_NAME As String)
    Dim rsACCT_CODE                                         As ADODB.Recordset
    Set rsACCT_CODE = New ADODB.Recordset
    rsACCT_CODE.Open "SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION ='" & xACCT_NAME & "'", gconDMIS, adOpenForwardOnly
    If Not rsACCT_CODE.EOF And Not rsACCT_CODE.BOF Then
        GET_ACCTCODE = rsACCT_CODE!AcctCode
    End If
    Set rsACCT_CODE = Nothing
End Function

Sub INIT_CBO_ATC()
'    WITHHOLDING TAX PAYABLE - EXPANDED

    Dim rsTAX                                               As ADODB.Recordset
    Set rsTAX = New ADODB.Recordset
    rsTAX.Open "SELECT UPPER(DESCRIPTION) AS DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 = 'EXPANDED' ", gconDMIS, adOpenKeyset
    If Not rsTAX.EOF And Not rsTAX.BOF Then
        Do While Not rsTAX.EOF
            cboAR_ACCT_CODE.AddItem Null2String(rsTAX!DESCRIPTION)
            rsTAX.MoveNext
        Loop
    End If
    Set rsTAX = Nothing

    Dim rsGET_ATC_DESC                                      As ADODB.Recordset
    Set rsGET_ATC_DESC = New ADODB.Recordset
'    rsGET_ATC_DESC.Open "SELECT DISTINCT atc.NATURE  FROM AMIS_AP AP INNER JOIN AMIS_JOURNAL_DET DT ON LEFT(AP.VOUCHERNO,3)= DT.JTYPE AND RIGHT(AP.VOUCHERNO,6)= DT.VOUCHERNO INNER JOIN AMIS_CHARTACCOUNT AC ON DT.Acct_Code = AC.AcctCode INNER JOIN AMIS_ATC ATC ON ATC.ATC = DT.ATC WHERE AC.TranType1 = 'EXPANDED' and AP.VENDOR_NAME = '" & txtName.Text & "'", gconDMIS, adOpenKeyset
    rsGET_ATC_DESC.Open "SELECT DISTINCT ATC.NATURE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DT ON HD.JType = DT.JType AND HD.VoucherNo = DT.VoucherNo INNER JOIN ALL_ENTITY AE ON HD.VendorCode = AE.CODE AND HD.ENTITY_CLASS = AE.ENTITYCODE INNER JOIN AMIS_CHARTACCOUNT AC ON DT.Acct_Code = AC.AcctCode INNER JOIN AMIS_ATC ATC ON DT.ATC = ATC.ATC WHERE AC.TranType1 = 'EXPANDED' AND AE.ACCOUNTNAME = '" & txtName.Text & "'", gconDMIS, adOpenKeyset
    cboATC.Clear
    cboATC.AddItem "ALL ATC"
    If Not rsGET_ATC_DESC.EOF And Not rsGET_ATC_DESC.BOF Then
        Do While Not rsGET_ATC_DESC.EOF
            cboATC.AddItem Null2String(rsGET_ATC_DESC!NATURE)
            rsGET_ATC_DESC.MoveNext
        Loop
    End If
    Set rsGET_ATC_DESC = Nothing
End Sub





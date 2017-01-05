VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmAMIS_APLEDGER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendors A/P Ledger"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAMIS_APLEDGER.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   11820
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   2790
      ScaleHeight     =   765
      ScaleWidth      =   5445
      TabIndex        =   23
      Top             =   7590
      Width           =   5475
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Double click the line item to open the corresponding journal entry of the particular voucher no."
         Height          =   435
         Left            =   630
         TabIndex        =   24
         Top             =   150
         Width           =   4725
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   150
         Picture         =   "frmAMIS_APLEDGER.frx":1082
         Top             =   180
         Width           =   360
      End
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
      Left            =   11070
      MouseIcon       =   "frmAMIS_APLEDGER.frx":2104
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_APLEDGER.frx":2256
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Exit Window"
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
      Left            =   10380
      MouseIcon       =   "frmAMIS_APLEDGER.frx":25BC
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_APLEDGER.frx":270E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Print this Record"
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
      Left            =   9690
      MouseIcon       =   "frmAMIS_APLEDGER.frx":2A74
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_APLEDGER.frx":2BC6
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Find a Record"
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
      Left            =   9000
      MouseIcon       =   "frmAMIS_APLEDGER.frx":2EC0
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_APLEDGER.frx":3012
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Move to Next Record"
      Top             =   7590
      Width           =   705
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
      Left            =   8310
      MouseIcon       =   "frmAMIS_APLEDGER.frx":336A
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_APLEDGER.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Move to Previous Record"
      Top             =   7590
      Width           =   705
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   375
      Left            =   7620
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
      Format          =   52494337
      CurrentDate     =   40210
   End
   Begin VB.CommandButton cmdOK 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11280
      Picture         =   "frmAMIS_APLEDGER.frx":381B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   495
   End
   Begin VB.ComboBox cboAR_ACCT_CODE 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1830
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   90
      Width           =   4935
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6945
      Left            =   2790
      TabIndex        =   1
      Top             =   600
      Width           =   8985
      Begin XtremeReportControl.ReportControl rptLEDGER 
         Height          =   5895
         Left            =   30
         TabIndex        =   9
         Top             =   990
         Width           =   8925
         _Version        =   655364
         _ExtentX        =   15743
         _ExtentY        =   10398
         _StockProps     =   64
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
         TabIndex        =   21
         Top             =   570
         Width           =   7215
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
         TabIndex        =   20
         Top             =   150
         Width           =   1275
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
         TabIndex        =   19
         Top             =   630
         Width           =   1635
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
         TabIndex        =   18
         Top             =   210
         Width           =   1665
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   60
         TabIndex        =   26
         Top             =   150
         Width           =   1665
         _Version        =   655364
         _ExtentX        =   2937
         _ExtentY        =   661
         _StockProps     =   14
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
         GradientColorDark=   16777215
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   375
         Left            =   60
         TabIndex        =   27
         Top             =   570
         Width           =   1665
         _Version        =   655364
         _ExtentX        =   2937
         _ExtentY        =   661
         _StockProps     =   14
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
         GradientColorDark=   16777215
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7785
      Left            =   30
      TabIndex        =   0
      Top             =   600
      Width           =   2715
      Begin XtremeReportControl.ReportControl rptENTITY 
         Height          =   6825
         Left            =   60
         TabIndex        =   10
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
         Left            =   60
         TabIndex        =   16
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
         TabIndex        =   22
         Top             =   150
         Width           =   2445
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   315
         Left            =   60
         TabIndex        =   25
         Top             =   120
         Width           =   2595
         _Version        =   655364
         _ExtentX        =   4577
         _ExtentY        =   556
         _StockProps     =   14
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
         GradientColorDark=   16777215
      End
   End
   Begin MSComCtl2.DTPicker dtTO 
      Height          =   375
      Left            =   9750
      TabIndex        =   8
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
      Format          =   52494337
      CurrentDate     =   40210
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
      Left            =   9300
      TabIndex        =   5
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From "
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
      Left            =   6810
      TabIndex        =   4
      Top             =   120
      Width           =   765
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1635
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   555
      Left            =   30
      TabIndex        =   17
      Top             =   0
      Width           =   11865
      _Version        =   655364
      _ExtentX        =   20929
      _ExtentY        =   979
      _StockProps     =   14
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
   End
End
Attribute VB_Name = "frmAMIS_APLEDGER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLOAD_VENDOR                                      As ADODB.Recordset
Dim rsLOAD_AP                                          As ADODB.Recordset
Dim REC                                                As XtremeReportControl.ReportRecord
Dim TOTAL_DEBIT                                        As Double
Dim TOTAL_CREDIT                                       As Double
Dim xBALANCE                                           As Double
Dim FWD_BALANCE                                        As Double
Dim FWD_DEBIT                                          As Double
Dim FWD_CREDIT                                         As Double
Dim rsREF                                              As ADODB.Recordset
Dim xlApplication                                      As Excel.Application
Dim xlWorkbook                                         As Excel.Workbook
Dim xlWorksheet                                        As Excel.Worksheet
Dim xlRange                                            As Excel.Range
Dim xCounter                                           As Integer

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
    With rptENTITY
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


    If rptLEDGER.Records.Count <= 1 Then
        MsgBox "No record to print."
        Exit Sub
    End If
    rptLEDGER.PrintOptions.Header.TextLeft = "" & COMPANY_NAME & "" & vbCrLf & "" & COMPANY_ADDRESS & "" & vbCrLf & vbCrLf & "CUSTOMER LEDGER " & vbCrLf & vbCrLf & "Period Covered From: " & "" & dtFrom.Value & " " & "To " & "" & dtTO.Value & "" & vbCrLf & vbCrLf & "Customer code: " & "" & txtCode.Text & "" & vbCrLf & "Customer Name: " & "" & txtName.Text & ""
    'rptLEDGER.PrintPreview True
    Screen.MousePointer = 11
    'Call LOAD_AP_ENTITY
    Call LOAD_AP_ENTITY_PRINTING
    Screen.MousePointer = 0
    LogAudit "V", "CUSTOMERS A/R LEDGER", txtCode
End Sub

Private Sub cmdOk_Click()
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
    Call LOAD_AP_ENTITY
    'Call DIRECT_DISBURSEMENT(txtCode.Text, SetAcctCode(cboAR_ACCT_CODE.Text))
End Sub

Function Setacctcode(XXX As String) As String
    Dim rsSetAcctCode                                  As ADODB.Recordset
    Set rsSetAcctCode = New ADODB.Recordset
    rsSetAcctCode.Open "SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "'", gconDMIS, adOpenForwardOnly
    If Not rsSetAcctCode.EOF And Not rsSetAcctCode.BOF Then
        Setacctcode = rsSetAcctCode!ACCTCODE
    End If
    Set rsSetAcctCode = Nothing
End Function
Private Sub Form_Load()
'UPDATED BY: JUN
'DATE UPDATED: 02132010
'DESCRIPTION: THIS CUSTOMER AR LEDGER WAS BASE ON AMIS_AP AND AMIS_DETAIL TABLE
'             WHICH AMIS_AP IS THE AMOUNT TO PAY BY THE CUSTOMER AND AMIS_DETAIL IS THE AMOUNT PAID DETAIL BY THE CUSTOMER
    CenterMe frmMain, Me, 1
    Call INIT_CTRL_ENTITY
    Call INIT_CTRL_LEDGER
    Call LOAD_ACCT_CODE
    Call LOAD_MAX_MIN_DATE
    Call LOAD_VENDOR
    cboAR_ACCT_CODE.ListIndex = 0
    Call rsRefresh
    Call StoreMemVars
End Sub

Sub LOAD_ACCT_CODE()
    Dim rsLOAD_ACCT_CODE                               As ADODB.Recordset
    Set rsLOAD_ACCT_CODE = New ADODB.Recordset
    rsLOAD_ACCT_CODE.Open "SELECT DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE LEFT(ACCTCODE,5) IN ('21-01','21-02','21-07')", gconDMIS, adOpenKeyset
    cboAR_ACCT_CODE.Clear
    cboAR_ACCT_CODE.AddItem "ALL"
    If Not rsLOAD_ACCT_CODE.EOF And Not rsLOAD_ACCT_CODE.BOF Then
        Do While Not rsLOAD_ACCT_CODE.EOF
            cboAR_ACCT_CODE.AddItem UCase(Null2String(rsLOAD_ACCT_CODE!DESCRIPTION))
            rsLOAD_ACCT_CODE.MoveNext
        Loop
    End If
    Set rsLOAD_ACCT_CODE = Nothing
End Sub

Sub LOAD_MAX_MIN_DATE()
    Dim rsMAX_MIN_DATE                                 As ADODB.Recordset
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


Private Sub rptENTITY_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
'If Row.Record Is Nothing Then: Exit Sub
'   MsgBox (Row.Record(1).Value)
End Sub

Private Sub rptENTITY_SelectionChanged()
'If rptENTITY.SelectedRows(0).Row.Record Is Nothing Then: Exit Sub
    txtCode.Text = UCase(rptENTITY.SelectedRows(0).Record(1).Value)
    txtName.Text = UCase(rptENTITY.SelectedRows(0).Record(0).Value)
    LOAD_MAX_MIN_DATE
    cmdOk_Click
End Sub

Sub LOAD_VENDOR()
    Set rsLOAD_VENDOR = New ADODB.Recordset
    'rsLOAD_VENDOR.Open "SELECT DISTINCT TOP 22 VENDOR_CODE,VENDOR_NAME FROM AMIS_AP AP INNER JOIN ALL_VENDOR_TABLE AVT ON AP.VENDOR_CODE=AVT.CODE AND AP.VENDOR_NAME=AVT.NAMEOFVENDOR ORDER BY VENDOR_NAME ASC", gconDMIS, adOpenKeyset
    rsLOAD_VENDOR.Open "SELECT DISTINCT TOP 22 VENDOR_CODE,VENDOR_NAME FROM AMIS_AP ORDER BY VENDOR_NAME ASC", gconDMIS, adOpenKeyset
    rptENTITY.Records.DeleteAll

    If Not rsLOAD_VENDOR.EOF And Not rsLOAD_VENDOR.BOF Then
        Do While Not rsLOAD_VENDOR.EOF
            Set REC = rptENTITY.Records.Add
            REC.AddItem (Trim(UCase(Null2String(rsLOAD_VENDOR!VENDOR_NAME))))
            REC.AddItem (Trim(UCase(Null2String(rsLOAD_VENDOR!VENDOR_CODE))))
            rptENTITY.Populate
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
        'Metrics.Font.Bold = True
    End If
End Sub

Private Sub rptLEDGER_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim xJOURNAL                                       As String
    Dim xVOUCHER                                       As String
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
    End If
End Sub

Private Sub textSearch_Change()
    Dim rssearch                                       As ADODB.Recordset
    Set rssearch = New ADODB.Recordset
    If TextSearch.Text <> "" Then
        'rssearch.Open "SELECT DISTINCT VENDOR_CODE,VENDOR_NAME FROM AMIS_AP AP INNER JOIN ALL_VENDOR_TABLE AVT ON AP.VENDOR_CODE=AVT.CODE AND AP.VENDOR_NAME=AVT.NAMEOFVENDOR WHERE VENDOR_NAME LIKE '" & Replace(TextSearch.Text, "'", "") & "%' ORDER BY VENDOR_NAME ASC", gconDMIS, adOpenKeyset
        rssearch.Open "SELECT DISTINCT VENDOR_CODE,VENDOR_NAME FROM AMIS_AP WHERE VENDOR_NAME LIKE '" & Replace(TextSearch.Text, "'", "") & "%' ORDER BY VENDOR_NAME ASC", gconDMIS, adOpenKeyset
    Else
        'rssearch.Open "SELECT DISTINCT TOP 22 VENDOR_CODE,VENDOR_NAME FROM AMIS_AP AP INNER JOIN ALL_VENDOR_TABLE AVT ON AP.VENDOR_CODE=AVT.CODE AND AP.VENDOR_NAME=AVT.NAMEOFVENDOR ORDER BY VENDOR_NAME ASC", gconDMIS, adOpenKeyset
        rssearch.Open "SELECT DISTINCT TOP 22 VENDOR_CODE,VENDOR_NAME FROM AMIS_AP ORDER BY VENDOR_NAME ASC", gconDMIS, adOpenKeyset
    End If

    rptENTITY.Records.DeleteAll
    If Not rssearch.EOF And Not rssearch.BOF Then
        Do While Not rssearch.EOF
            Set REC = rptENTITY.Records.Add
            REC.AddItem (Trim(UCase(Null2String(rssearch!VENDOR_NAME))))
            REC.AddItem (Trim(UCase(Null2String(rssearch!VENDOR_CODE))))
            rptENTITY.Populate
            Set REC = Nothing
            rssearch.MoveNext
        Loop
    End If
    Set rssearch = Nothing
End Sub

Sub LOAD_AP_ENTITY()
    Dim xVOUCHERNO                                     As String
    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0: xBALANCE = 0
    Set rsLOAD_AP = New ADODB.Recordset
    If cboAR_ACCT_CODE.Text = "ALL" Then
        'rsLOAD_AP.Open "SELECT JDATE,VOUCHERNO AS APVOUCHERNO,ISNULL(INVOICETYPE,'') + CASE WHEN INVOICETYPE IS NULL THEN '' ELSE '-' END + ISNULL(INVOICENO,'') AS INVOICE,AMOUNT2PAY,AMOUNTPAID,INVOICENO,INVOICETYPE,VENDOR_CODE,AP.ACCT_CODE,RIGHT(VOUCHERNO,6) AS VOUCHERNO FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE INNER JOIN ALL_VENDOR_TABLE AVT ON AP.VENDOR_CODE=AVT.CODE AND AP.VENDOR_NAME=AVT.NAMEOFVENDOR WHERE AP.VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' ORDER BY JDATE,VOUCHERNO", gconDMIS, adOpenKeyset
        rsLOAD_AP.Open "SELECT JDATE,VOUCHERNO AS APVOUCHERNO,ISNULL(INVOICETYPE,'') + CASE WHEN INVOICETYPE IS NULL THEN '' ELSE '-' END + ISNULL(INVOICENO,'') AS INVOICE,AMOUNT2PAY,AMOUNTPAID,INVOICENO,INVOICETYPE,VENDOR_CODE,AP.ACCT_CODE,RIGHT(VOUCHERNO,6) AS VOUCHERNO FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE WHERE AC.TRANTYPE2<>'DEPOSIT' AND AP.VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' ORDER BY JDATE,VOUCHERNO", gconDMIS, adOpenKeyset
    Else
        'rsLOAD_AP.Open "SELECT JDATE,VOUCHERNO AS APVOUCHERNO,ISNULL(INVOICETYPE,'') + '-' + ISNULL(INVOICENO,'') AS INVOICE,AMOUNT2PAY,AMOUNTPAID,INVOICENO,INVOICETYPE,VENDOR_CODE,AP.ACCT_CODE,RIGHT(VOUCHERNO,6) AS VOUCHERNO FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE INNER JOIN ALL_VENDOR_TABLE AVT ON AP.VENDOR_CODE=AVT.CODE AND AP.VENDOR_NAME=AVT.NAMEOFVENDOR WHERE AP.VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' AND " & _
                       "ACCT_CODE = (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "') ORDER BY JDATE,APVOUCHERNO", gconDMIS, adOpenKeyset
        rsLOAD_AP.Open "SELECT JDATE,VOUCHERNO AS APVOUCHERNO,ISNULL(INVOICETYPE,'') + '-' + ISNULL(INVOICENO,'') AS INVOICE,AMOUNT2PAY,AMOUNTPAID,INVOICENO,INVOICETYPE,VENDOR_CODE,AP.ACCT_CODE,RIGHT(VOUCHERNO,6) AS VOUCHERNO FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE WHERE AC.TRANTYPE2<>'DEPOSIT' AND AP.VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' AND " & _
                       "ACCT_CODE = (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "') ORDER BY JDATE,APVOUCHERNO", gconDMIS, adOpenKeyset
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
            If xVOUCHERNO = "VPJ" Or xVOUCHERNO = "APJ" Or xVOUCHERNO = "CDJ" Or xVOUCHERNO = "GJ" Or xVOUCHERNO = "SJ" Then
                Set REC = rptLEDGER.Records.Add
                REC.AddItem (Trim(Null2String(rsLOAD_AP!JDATE)))
                REC.AddItem (Trim(Null2String(rsLOAD_AP!APVoucherno)))
                'REC.AddItem (Trim(Null2String(rsLOAD_AP!invoice)))
                If xVOUCHERNO = "APJ" Or xVOUCHERNO = "VPJ" Then
                    REC.AddItem GetInvoices(Right((Trim(Null2String(rsLOAD_AP!APVoucherno))), 6), Left((Trim(Null2String(rsLOAD_AP!APVoucherno))), 3))
                Else
                    REC.AddItem (Trim(Null2String(rsLOAD_AP!INVOICE)))
                End If
                If NumericVal(rsLOAD_AP!AMOUNT2PAY) <> 0 Then
                    REC.AddItem (Trim("0.00"))
                    REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsLOAD_AP!AMOUNT2PAY))))

                    TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))

                    REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
                Else
                    REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsLOAD_AP!AMOUNTPAID))))
                    REC.AddItem (Trim("0.00"))

                    TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_AP!AMOUNTPAID)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_AP!AMOUNTPAID)), 2))

                    REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
                End If

                'If Null2String(rsLOAD_AP!APVoucherno) = "VPJ-000071" Then Stop
                Call REFERENCE_VOUCHER(Null2String(rsLOAD_AP!APVoucherno), txtCode.Text, Null2String(rsLOAD_AP!ACCT_CODE))
                Call ADJUSTMENT_BYVOUHCHERNO(Null2String(Right(rsLOAD_AP!APVoucherno, 6)), xVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AP!ACCT_CODE))
                rptLEDGER.Populate
                Set REC = Nothing
                '                            ElseIf xVOUCHERNO = "SJ" Or xVOUCHERNO = "CRJ" Then
                '                                Set REC = rptLEDGER.Records.Add
                '                                REC.AddItem (Trim(Null2String(rsLOAD_AP!JDate)))
                '                                REC.AddItem (Trim(Null2String(rsLOAD_AP!APVoucherno)))
                '                                REC.AddItem (Trim(Null2String(rsLOAD_AP!invoice)))
                '                                If NumericVal(rsLOAD_AP!AMOUNT2PAY) <> 0 Then
                '                                    REC.AddItem (Trim("0.00"))
                '                                    REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsLOAD_AP!AMOUNT2PAY))))
                '
                '
                '                                    TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))
                '                                    xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))
                '
                '                                    REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
                '                                Else
                '                                    REC.AddItem (Trim(Null2String(ToDoubleNumber(NumericVal(rsLOAD_AP!AMOUNTPAID)))))
                '                                    REC.AddItem (Trim("0.00"))
                '
                '
                '                                    TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_AP!AMOUNTPAID)), 2))
                '                                    xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_AP!AMOUNTPAID)), 2))
                '
                '                                    REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
                '                                End If
                '
                '
                '                Call REFERENCE_VOUCHER(Null2String(rsLOAD_AP!APVoucherno), txtCode.Text, Null2String(rsLOAD_AP!Acct_code))
                '
                '                Call ADJUSTMENT_BYVOUHCHERNO(Null2String(Right(rsLOAD_AP!APVoucherno, 6)), xVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AP!Acct_code))
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

Sub LOAD_AP_ENTITY_PRINTING()
    Dim xVOUCHERNO                                     As String
    Screen.MousePointer = 11
    xCounter = 10
    If Len(Dir(AMIS_REPORT_PATH & "\Ledgers\VendorsSubsidiaryLedger.xlt")) = 0 Then
        MsgBox "No Excel file found", vbInformation, "System Message"
        Exit Sub
    End If
    Set xlApplication = CreateObject("Excel.Application")
    Set xlWorkbook = xlApplication.Workbooks.Open(AMIS_REPORT_PATH & "\Ledgers\VendorsSubsidiaryLedger.xlt")
    Set xlWorksheet = xlWorkbook.Worksheets(1)
    xlWorksheet.Cells(1, "A") = COMPANY_NAME
    xlWorksheet.Cells(1, "A").Font.Bold = True
    xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
    xlWorksheet.Cells(2, "A").Font.Bold = True
    xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTO.Value, "mm/dd/yyyy")
    xlWorksheet.Cells(3, "A").Font.Bold = True
    xlWorksheet.Cells(4, "A") = "VENDOR SUBSIDIARY LEDGER"
    xlWorksheet.Cells(4, "A").Font.Bold = True
    xlWorksheet.Cells(6, "A") = "Vendor Code: " & txtCode.Text
    xlWorksheet.Cells(6, "A").Font.Bold = True
    xlWorksheet.Cells(7, "A") = "Vendor Name: " & txtName.Text
    xlWorksheet.Cells(7, "A").Font.Bold = True

    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0: xBALANCE = 0
    'updated by  arjr requested by HPI
    Set rsLOAD_AP = New ADODB.Recordset
    If cboAR_ACCT_CODE.Text = "ALL" Then
        ' rsLOAD_AP.Open "SELECT JDATE,VOUCHERNO AS APVOUCHERNO,INVOICETYPE + '-' + INVOICENO AS INVOICE,AMOUNT2PAY,AMOUNTPAID,INVOICENO,INVOICETYPE,VENDOR_CODE,AP.ACCT_CODE,RIGHT(VOUCHERNO,6) AS VOUCHERNO FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE WHERE VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "'", gconDMIS, adOpenKeyset
        rsLOAD_AP.Open "SELECT ap.JDATE,ap.VOUCHERNO AS APVOUCHERNO,ap.INVOICETYPE + '-' + ap.INVOICENO AS INVOICE,ap.AMOUNT2PAY,ap.AMOUNTPAID,ap.INVOICENO,ap.INVOICETYPE,ap.VENDOR_CODE,AP.ACCT_CODE,RIGHT(ap.VOUCHERNO,6) AS VOUCHERNO,hd.remarks FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE inner join amis_journal_hd hd on ap.voucherno=hd.jtype+'-'+hd.voucherno  WHERE ap.VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND ap.Jdate >= '" & dtFrom.Value & "'and ap.Jdate <= '" & dtTO.Value & "'", gconDMIS, adOpenKeyset
    Else
        '  rsLOAD_AP.Open "SELECT JDATE,VOUCHERNO AS APVOUCHERNO,INVOICETYPE + '-' + INVOICENO AS INVOICE,AMOUNT2PAY,AMOUNTPAID,INVOICENO,INVOICETYPE,VENDOR_CODE,AP.ACCT_CODE,RIGHT(VOUCHERNO,6) AS VOUCHERNO FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE WHERE VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' AND " & _
           '                  "ACCT_CODE = (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "')", gconDMIS, adOpenKeyset
        rsLOAD_AP.Open "SELECT ap.JDATE,ap.VOUCHERNO AS APVOUCHERNO,ap.INVOICETYPE + '-' + ap.INVOICENO AS INVOICE,ap.AMOUNT2PAY,ap.AMOUNTPAID,ap.INVOICENO,ap.INVOICETYPE,ap.VENDOR_CODE,AP.ACCT_CODE,RIGHT(ap.VOUCHERNO,6) AS VOUCHERNO, hd.remarks FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE inner join amis_journal_hd hd on ap.voucherno=hd.jtype+'-'+hd.voucherno WHERE ap.VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND ap.Jdate >= '" & dtFrom.Value & "'and ap.Jdate <= '" & dtTO.Value & "' AND " & _
                       "ap.ACCT_CODE = (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "')", gconDMIS, adOpenKeyset
    End If
    '       Set rsLOAD_AP = New ADODB.Recordset
    '        If cboAR_ACCT_CODE.Text = "ALL" Then
    '        rsLOAD_AP.Open "SELECT JDATE,VOUCHERNO AS APVOUCHERNO,INVOICETYPE + '-' + INVOICENO AS INVOICE,AMOUNT2PAY,AMOUNTPAID,INVOICENO,INVOICETYPE,VENDOR_CODE,AP.ACCT_CODE,RIGHT(VOUCHERNO,6) AS VOUCHERNO FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE WHERE VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "'", gconDMIS, adOpenKeyset
    '         Else
    '        rsLOAD_AP.Open "SELECT JDATE,VOUCHERNO AS APVOUCHERNO,INVOICETYPE + '-' + INVOICENO AS INVOICE,AMOUNT2PAY,AMOUNTPAID,INVOICENO,INVOICETYPE,VENDOR_CODE,AP.ACCT_CODE,RIGHT(VOUCHERNO,6) AS VOUCHERNO FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE WHERE VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "' AND " & _
             '                       "ACCT_CODE = (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "')", gconDMIS, adOpenKeyset
    '        End If

    Call FORWARDED_BALANCE

    xlWorksheet.Cells(xCounter, "A") = Format(dtFrom.Value, "mm/dd/yyyy")
    xlWorksheet.Cells(xCounter, "B") = "FWD BALANCE"
    xlWorksheet.Cells(xCounter, "D") = "0.00"
    xlWorksheet.Cells(xCounter, "E") = "0.00"
    xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(FWD_BALANCE)))

    If Not rsLOAD_AP.EOF And Not rsLOAD_AP.BOF Then
        Do While Not rsLOAD_AP.EOF
            If Len(rsLOAD_AP!APVoucherno) = 10 Then
                xVOUCHERNO = Left(Null2String(rsLOAD_AP!APVoucherno), 3)
            Else
                xVOUCHERNO = Left(Null2String(rsLOAD_AP!APVoucherno), 2)
            End If
            If xVOUCHERNO = "VPJ" Or xVOUCHERNO = "COB" Or xVOUCHERNO = "APJ" Or xVOUCHERNO = "CDJ" Or xVOUCHERNO = "SJ" Or xVOUCHERNO = "CRJ" Or xVOUCHERNO = "GJ" Then
                xCounter = xCounter + 1
                xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsLOAD_AP!JDATE))), "mm/dd/yyyy")
                xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_AP!APVoucherno)))
                xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsLOAD_AP!INVOICE)))
                If NumericVal(rsLOAD_AP!AMOUNT2PAY) <> 0 Then
                    xlWorksheet.Cells(xCounter, "D") = (Trim("0.00"))
                    xlWorksheet.Cells(xCounter, "E") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_AP!AMOUNT2PAY))))

                    TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))

                    xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
                Else
                    xlWorksheet.Cells(xCounter, "D") = (Trim(Null2String(rsLOAD_AP!AMOUNTPAID)))
                    xlWorksheet.Cells(xCounter, "E") = (Trim("0.00"))


                    TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_AP!AMOUNTPAID)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_AP!AMOUNTPAID)), 2))

                    xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
                End If
                'update by arjr
                xlWorksheet.Cells(xCounter, "G") = (Trim(Null2String(rsLOAD_AP!remarks)))
                Call REFERENCE_VOUCHER_PRINTING(Null2String(rsLOAD_AP!APVoucherno), txtCode.Text, Null2String(rsLOAD_AP!ACCT_CODE))
                'Call ADJUSTMENT_BYVOUHCHERNO_PRINTING(Null2String(rsLOAD_AP!VOUCHERNO), xSJVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AP!Account_code))
                '            ElseIf xSJVOUCHERNO = "APJ" Or xSJVOUCHERNO = "CDJ" Then
                '                xCounter = xCounter + 1
                '                xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsLOAD_AP!JDate))), "mm/dd/yyyy")
                '                xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_AP!SJVoucherno)))
                '                xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsLOAD_AP!invoice)))
                '                If NumericVal(rsLOAD_AP!AMOUNT_TOPAY) <> 0 Then
                '                    xlWorksheet.Cells(xCounter, "D") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_AP!AMOUNT_TOPAY))))
                '                    xlWorksheet.Cells(xCounter, "E") = (Trim("0.00"))
                '
                '                    TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_AP!AMOUNT_TOPAY)), 2))
                '                    xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_AP!AMOUNT_TOPAY)), 2))
                '
                '                    xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
                '                Else
                '                    xlWorksheet.Cells(xCounter, "D") = (Trim("0.00"))
                '                    xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_AP!AMOUNT_PAID)))
                '
                '                    TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_AP!AMOUNT_TOPAY)), 2))
                '                    xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_AP!AMOUNT_PAID)), 2))
                '
                '                    xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
                '                End If

                '                Call REFERENCE_INVOICE_PRINTING(Right(Null2String(rsLOAD_AP!SJVoucherno), 6), Left(Null2String(rsLOAD_AP!SJVoucherno), 3), txtCode.Text, Null2String(rsLOAD_AP!Account_code))
                '                Call ADJUSTMENT_BYVOUHCHERNO_PRINTING(Null2String(rsLOAD_AP!VOUCHERNO), xSJVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AP!Account_code))
            End If
            rsLOAD_AP.MoveNext
            DoEvents
        Loop
    End If
    xCounter = xCounter + 2
    xlWorksheet.Cells(xCounter, "C") = "TOTAL"
    xlWorksheet.Cells(xCounter, "C").Font.Bold = True
    xlWorksheet.Cells(xCounter, "D") = ToDoubleNumber(TOTAL_DEBIT)
    xlWorksheet.Cells(xCounter, "D").Font.Bold = True
    xlWorksheet.Cells(xCounter, "E") = ToDoubleNumber(TOTAL_CREDIT)
    xlWorksheet.Cells(xCounter, "E").Font.Bold = True
    xlWorksheet.Cells(xCounter, "F") = ToDoubleNumber(Round(NumericVal(xBALANCE + FWD_BALANCE), 2))
    xlWorksheet.Cells(xCounter, "F").Font.Bold = True
    xlApplication.Visible = True
    Set xlApplication = Nothing
    Set xlWorkbook = Nothing
    Set xlWorksheet = Nothing
    Set rsLOAD_AP = Nothing
    Screen.MousePointer = 0
End Sub

Sub REFERENCE_VOUCHER(xVOUCHERNO As String, xVENDORCODE As String, xAcctCode As String)
    Dim rsPAYMENT                                      As ADODB.Recordset
    Set rsPAYMENT = New ADODB.Recordset
    rsPAYMENT.Open "SELECT JDATE,JTYPE,VOUCHERNO,ISNULL(INVOICETYPE,'') + CASE WHEN INVOICETYPE IS NULL THEN '' ELSE '-' END + ISNULL(INVOICENO,'') AS INVOICENO,AMOUNTPAID FROM AMIS_DETAILS WHERE PV_VOUCHERNO = '" & xVOUCHERNO & "' AND VENDORCODE = '" & xVENDORCODE & "' AND ACCT_CODE = '" & xAcctCode & "' " & _
                   "AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "'", gconDMIS, adOpenKeyset
    If Not rsPAYMENT.EOF And Not rsPAYMENT.BOF Then
        Do While Not rsPAYMENT.EOF
            Set REC = rptLEDGER.Records.Add
            REC.AddItem (Trim(Null2String(rsPAYMENT!JDATE)))
            REC.AddItem (Trim(Null2String(rsPAYMENT!jtype) & "-" & Null2String(rsPAYMENT!VOUCHERNO)))
            REC.AddItem (Trim(Null2String(rsPAYMENT!INVOICENO)))
            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsPAYMENT!AMOUNTPAID))))
            REC.AddItem (Trim("0.00"))

            TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsPAYMENT!AMOUNTPAID)), 2))
            xBALANCE = Trim(ToDoubleNumber(xBALANCE - NumericVal(rsPAYMENT!AMOUNTPAID)))

            REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
            rptLEDGER.Populate
            Set REC = Nothing
            rsPAYMENT.MoveNext
            DoEvents
        Loop
        'Call ADJUSTMENT_BYVOUHCHERNO(xINVOICENO, xINVOICETYPE, txtCode.Text, xAcctCode)
    End If
    Set rsPAYMENT = Nothing
End Sub

Sub DIRECT_DISBURSEMENT(xVENDORCODE As String, xAcctCode As String)
    Dim rsDIRECT_DISBURSEMENT                          As ADODB.Recordset
    Set rsDIRECT_DISBURSEMENT = New ADODB.Recordset
    If cboAR_ACCT_CODE.Text = "ALL" Then
        rsDIRECT_DISBURSEMENT.Open "SELECT JDATE,VOUCHERNO,AMOUNT2PAY,AMOUNTPAID,INVOICENO FROM AMIS_AP WHERE VENDOR_CODE='" & xVENDORCODE & "' AND LEFT(VOUCHERNO,3)='CDJ' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "'", gconDMIS, adOpenForwardOnly
    Else
        rsDIRECT_DISBURSEMENT.Open "SELECT JDATE,VOUCHERNO,AMOUNT2PAY,AMOUNTPAID,INVOICENO FROM AMIS_AP WHERE VENDOR_CODE='" & xVENDORCODE & "' AND ACCT_CODE='" & xAcctCode & "' AND LEFT(VOUCHERNO,3)='CDJ' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "'", gconDMIS, adOpenForwardOnly
    End If

    If Not rsDIRECT_DISBURSEMENT.EOF And Not rsDIRECT_DISBURSEMENT.BOF Then
        Do While Not rsDIRECT_DISBURSEMENT.EOF
            Set REC = rptLEDGER.Records.Add
            REC.AddItem (Trim(Null2String(rsDIRECT_DISBURSEMENT!JDATE)))
            REC.AddItem (Trim(Null2String(rsDIRECT_DISBURSEMENT!VOUCHERNO)))
            REC.AddItem (Trim(Null2String(rsDIRECT_DISBURSEMENT!INVOICENO)))
            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsDIRECT_DISBURSEMENT!AMOUNTPAID))))
            REC.AddItem (Trim("0.00"))

            TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsDIRECT_DISBURSEMENT!AMOUNTPAID)), 2))
            xBALANCE = Trim(ToDoubleNumber(xBALANCE - NumericVal(rsDIRECT_DISBURSEMENT!AMOUNTPAID)))

            REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
            rptLEDGER.Populate
            Set REC = Nothing
            rsDIRECT_DISBURSEMENT.MoveNext
            DoEvents
        Loop
    End If
    rptLEDGER.Columns(3).FooterText = ToDoubleNumber(TOTAL_DEBIT)
    rptLEDGER.Columns(4).FooterText = ToDoubleNumber(TOTAL_CREDIT)
    rptLEDGER.Columns(5).FooterText = ToDoubleNumber(Round(NumericVal(xBALANCE + FWD_BALANCE), 2))
    Set rsDIRECT_DISBURSEMENT = Nothing
End Sub

Sub REFERENCE_VOUCHER_PRINTING(xVOUCHERNO As String, xVENDORCODE As String, xAcctCode As String)
    Dim rsPAYMENT                                      As ADODB.Recordset
    Set rsPAYMENT = New ADODB.Recordset
    rsPAYMENT.Open "SELECT JDATE,JTYPE,VOUCHERNO,ISNULL(INVOICETYPE,'') + CASE WHEN INVOICETYPE IS NULL THEN '' ELSE '-' END + ISNULL(INVOICENO,'') AS INVOICENO,AMOUNTPAID FROM AMIS_DETAILS WHERE PV_VOUCHERNO = '" & xVOUCHERNO & "' AND VENDORCODE = '" & xVENDORCODE & "' AND ACCT_CODE = '" & xAcctCode & "' " & _
                   "AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "'", gconDMIS, adOpenKeyset
    If Not rsPAYMENT.EOF And Not rsPAYMENT.BOF Then
        Do While Not rsPAYMENT.EOF
            xCounter = xCounter + 1
            xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsPAYMENT!JDATE))), "mm/dd/yyyy")
            xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsPAYMENT!jtype) & "-" & Null2String(rsPAYMENT!VOUCHERNO)))
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
        'Call ADJUSTMENT_BYVOUHCHERNO(xINVOICENO, xINVOICETYPE, txtCode.Text, xAcctCode)
    End If
    Set rsPAYMENT = Nothing
End Sub

Sub ADJUSTMENT_BYVOUHCHERNO(xADJVOUCHERNO, xADJTYPE, xVENDORCODE As String, xACCT_CODE As String)
    Dim rsADJ                                          As ADODB.Recordset
    Set rsADJ = New ADODB.Recordset
    rsADJ.Open "SELECT JDATE, VOUCHERNO, INVOICETYPE + '-' + INVOICENO AS REF_INVOICE,AMOUNT2PAY,AMOUNTPAID " & _
               "FROM AMIS_AP WHERE VENDOR_CODE = '" & xVENDORCODE & "' AND " & _
               "Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' AND INVOICENO='" & xADJVOUCHERNO & "' AND INVOICETYPE='" & xADJTYPE & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND LEFT(VOUCHERNO,2) = 'GJ'", gconDMIS, adOpenKeyset
    'rsADJ.Open "SELECT JDATE, SJVOUCHERNO, INVOICETYPE + '-' + INVOICENO AS REF_INVOICE,AMOUNT_TOPAY,AMOUNT_PAID " & _
     "FROM AMIS_AP WHERE VENDOR_CODE = '" & xCUSCDE & "' AND INVOICETYPE = '" & xJtype & "' AND " & _
     "Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' AND  INVOICENO = '" & XVOUCHERNO & "' AND ACCOUNT_CODE = '" & xACCT_CODE & "' AND LEFT(SJVOUCHERNO,2) = '" & xxSJVOUCHERNO & "'", gconDMIS, adOpenKeyset
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
    Dim rsADJ                                          As ADODB.Recordset
    Set rsADJ = New ADODB.Recordset
    rsADJ.Open "SELECT JDATE, SJVOUCHERNO, INVOICETYPE + '-' + INVOICENO AS REF_INVOICE,AMOUNT_TOPAY,AMOUNT_PAID " & _
               "FROM AMIS_AP WHERE VENDOR_CODE = '" & xCUSCDE & "' AND INVOICETYPE = '" & xJType & "' AND " & _
               "Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' AND  INVOICENO = '" & xVOUCHERNO & "' AND ACCOUNT_CODE = '" & xACCT_CODE & "' AND LEFT(SJVOUCHERNO,2) = 'GJ'", gconDMIS, adOpenKeyset
    'rsADJ.Open "SELECT JDATE, SJVOUCHERNO, INVOICETYPE + '-' + INVOICENO AS REF_INVOICE,AMOUNT_TOPAY,AMOUNT_PAID " & _
     "FROM AMIS_AP WHERE VENDOR_CODE = '" & xCUSCDE & "' AND INVOICETYPE = '" & xJtype & "' AND " & _
     "Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' AND  INVOICENO = '" & XVOUCHERNO & "' AND ACCOUNT_CODE = '" & xACCT_CODE & "' AND LEFT(SJVOUCHERNO,2) = '" & xxSJVOUCHERNO & "'", gconDMIS, adOpenKeyset
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
        rptENTITY.SetFocus
    End If
End Sub

Private Sub TextSearch_LostFocus()
    TextSearch.BackColor = vbWhite
End Sub

Sub FORWARDED_BALANCE()
    Dim rsLOAD_AP                                      As ADODB.Recordset
    Dim xVOUCHERNO                                     As String
    Set rsLOAD_AP = New ADODB.Recordset

    FWD_BALANCE = 0: FWD_CREDIT = 0: FWD_DEBIT = 0

    If cboAR_ACCT_CODE.Text = "ALL" Then
        rsLOAD_AP.Open "SELECT JDATE,VOUCHERNO AS APVOUCHERNO,INVOICETYPE + '-' + INVOICENO AS INVOICE,AMOUNT2PAY,AMOUNTPAID,INVOICENO,INVOICETYPE,VENDOR_CODE,ACCT_CODE,RIGHT(VOUCHERNO,6) AS VOUCHERNO " & _
                       "FROM AMIS_AP WHERE VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate < '" & dtFrom.Value & "'", gconDMIS, adOpenKeyset
    Else
        rsLOAD_AP.Open "SELECT JDATE,VOUCHERNO AS APVOUCHERNO,INVOICETYPE + '-' + INVOICENO,AMOUNT2PAY,AMOUNTPAID,INVOICENO,INVOICETYPE,VENDOR_CODE,ACCT_CODE,RIGHT(VOUCHERNO,6) AS VOUCHERNO " & _
                       "FROM AMIS_AP WHERE VENDOR_CODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate < '" & dtFrom.Value & "' AND " & _
                       "ACCT_CODE = (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "')", gconDMIS, adOpenKeyset
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

                'Call FWD_REFERENCE_INVOICE(Null2String(rsLOAD_AP!INVOICENO), Null2String(rsLOAD_AP!InvoiceType), txtCode.Text, Null2String(rsLOAD_AP!Account_code))
                'Call FWD_ADJUSTMENT_BYVOUHCHERNO(Null2String(rsLOAD_AP!VOUCHERNO), xVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AP!Account_code))

                '            ElseIf xVOUCHERNO = "SJ" Or xVOUCHERNO = "CRJ" Then
                '                If NumericVal(rsLOAD_AP!AMOUNT_TOPAY) <> 0 Then
                '                    FWD_DEBIT = ToDoubleNumber(Round((FWD_DEBIT + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))
                '                    FWD_BALANCE = ToDoubleNumber(Round((FWD_BALANCE + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))
                '                Else
                '                    FWD_CREDIT = ToDoubleNumber(Round((FWD_CREDIT + NumericVal(rsLOAD_AP!AMOUNT2PAY)), 2))
                '                    FWD_BALANCE = ToDoubleNumber(Round((FWD_BALANCE - NumericVal(rsLOAD_AP!AMOUNTPAID)), 2))
                '                End If

                'Call FWD_REFERENCE_INVOICE(Right(Null2String(rsLOAD_AP!VOUCHERNO), 6), Left(Null2String(rsLOAD_AP!VOUCHERNO), 3), txtCode.Text, Null2String(rsLOAD_AP!Account_code))
            End If
            rsLOAD_AP.MoveNext
        Loop
    End If
End Sub

Sub FWD_REFERENCE_INVOICE(xInvoiceNo As String, xInvoiceType As String, xCUSCODE As String, xAcctCode As String)
    Dim rsINVOICE                                      As ADODB.Recordset
    Set rsINVOICE = New ADODB.Recordset
    rsINVOICE.Open "SELECT * FROM AMIS_DETAIL WHERE INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "' AND VENDOR_CODE = '" & xCUSCODE & "' AND ACCT_CODE = '" & xAcctCode & "' " & _
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
    Dim rsADJ                                          As ADODB.Recordset
    Set rsADJ = New ADODB.Recordset
    rsADJ.Open "SELECT JDATE, SJVOUCHERNO, INVOICETYPE + '-' + INVOICENO AS REF_INVOICE,AMOUNT_TOPAY,AMOUNT_PAID " & _
               "FROM AMIS_AP WHERE VENDOR_CODE = '" & xCUSCDE & "' AND INVOICETYPE = '" & xJType & "' AND " & _
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
    rsREF.Open "SELECT DISTINCT VENDOR_NAME,VENDOR_CODE FROM AMIS_AP ORDER BY VENDOR_NAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsREF.EOF And Not rsREF.BOF Then
        txtCode.Text = UCase(Null2String(rsREF!VENDOR_CODE))
        txtName.Text = UCase(Null2String(rsREF!VENDOR_NAME))
    End If
    Call LOAD_AP_ENTITY
End Sub

Function GetInvoices(VOUCHERNO As String, jtype As String) As String
    Dim rsInvoices As ADODB.Recordset
    Set rsInvoices = New ADODB.Recordset
    rsInvoices.Open "SELECT INV_NO FROM AMIS_PV_DETAIL WHERE VOUCHERNO = '" & VOUCHERNO & "' AND JTYPE='" & jtype & "'", gconDMIS, adOpenForwardOnly
    If Not rsInvoices.EOF And Not rsInvoices.BOF Then
        Do While Not rsInvoices.EOF
            GetInvoices = GetInvoices + "," + Null2String(rsInvoices!INV_NO)
            rsInvoices.MoveNext
        Loop
    End If
    GetInvoices = Mid(GetInvoices, 2, Len(GetInvoices))
    Set rsInvoices = Nothing
End Function

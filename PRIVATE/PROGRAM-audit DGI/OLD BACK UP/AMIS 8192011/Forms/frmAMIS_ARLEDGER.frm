VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmAMIS_ARLEDGER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers A/R Ledger"
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
   Icon            =   "frmAMIS_ARLEDGER.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   11820
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Picture         =   "frmAMIS_ARLEDGER.frx":1082
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
      MouseIcon       =   "frmAMIS_ARLEDGER.frx":2104
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_ARLEDGER.frx":2256
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
      MouseIcon       =   "frmAMIS_ARLEDGER.frx":25BC
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_ARLEDGER.frx":270E
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
      MouseIcon       =   "frmAMIS_ARLEDGER.frx":2A74
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_ARLEDGER.frx":2BC6
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
      MouseIcon       =   "frmAMIS_ARLEDGER.frx":2EC0
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_ARLEDGER.frx":3012
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
      MouseIcon       =   "frmAMIS_ARLEDGER.frx":336A
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_ARLEDGER.frx":34BC
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
      Format          =   48889857
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
      Picture         =   "frmAMIS_ARLEDGER.frx":381B
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
         Caption         =   "Account Name"
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
         Caption         =   "Account Code"
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
         Width           =   1605
         _Version        =   655364
         _ExtentX        =   2831
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
         Width           =   1605
         _Version        =   655364
         _ExtentX        =   2831
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
            Size            =   9.01
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
      Format          =   48889857
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
Attribute VB_Name = "frmAMIS_ARLEDGER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLOAD_CUSTOMER                                    As ADODB.Recordset
Dim rsLOAD_AR                                          As ADODB.Recordset
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
        .Columns.Add 2, "INVOICE#/OR", 110, True: .Columns(2).Alignment = xtpAlignmentCenter: .Columns(2).AllowRemove = False
        .Columns.Add 3, "DEBIT", 110, True: .Columns(3).Alignment = xtpAlignmentRight: .Columns(3).AllowRemove = False
        .Columns.Add 4, "CREDIT", 110, True: .Columns(4).Alignment = xtpAlignmentRight: .Columns(4).AllowRemove = False
        .Columns.Add 5, "BALANCE", 80, True: .Columns(5).Alignment = xtpAlignmentRight: .Columns(5).AllowRemove = False
        .Columns.Add 6, "ID", 0, True: .Columns(6).Alignment = xtpAlignmentIconRight: .Columns(6).AllowRemove = False: .Columns(6).Visible = False
        .Columns.Add 7, "jtype", 0, True: .Columns(7).Alignment = xtpAlignmentIconRight: .Columns(7).AllowRemove = False: .Columns(7).Visible = False

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

    Set REC = rptLEDGER.Records.Add
    REC.AddItem(Trim("")).Record.Visible = True
    REC.AddItem(Trim("")).Record.Visible = True
    REC.AddItem(Trim("TOTAL")).Record.Visible = True
    REC.AddItem(Trim(ToDoubleNumber(TOTAL_DEBIT))).Record.Visible = True
    REC.AddItem(Trim(ToDoubleNumber(TOTAL_CREDIT))).Record.Visible = True
    REC.AddItem(Trim(ToDoubleNumber(xBALANCE + FWD_BALANCE))).Record.Visible = True
    REC.AddItem(Trim("")).Record.Visible = True
    REC.AddItem(Trim("")).Record.Visible = True
    rptLEDGER.Populate
    Set REC = Nothing

    rptLEDGER.PrintOptions.Header.TextLeft = "" & COMPANY_NAME & "" & vbCrLf & "" & COMPANY_ADDRESS & "" & vbCrLf & vbCrLf & "CUSTOMER LEDGER " & vbCrLf & vbCrLf & "Period Covered From: " & "" & dtFrom.Value & " " & "To " & "" & dtTo.Value & "" & vbCrLf & vbCrLf & "Customer code: " & "" & txtCode.Text & "" & vbCrLf & "Customer Name: " & "" & txtName.Text & ""
    'rptLEDGER.PrintPreview True
    Call LOAD_AR_ENTITY
    Call LOAD_AR_ENTITY_PRINTING

    LogAudit "V", "CUSTOMERS A/R LEDGER", txtCode
End Sub

Private Sub cmdOk_Click()
    If cboAR_ACCT_CODE.Text = "" Then
        MessagePop InfoFriend, "SYSTEM MESSAGE", "Please select the account code."
        cboAR_ACCT_CODE.SetFocus
        Exit Sub
    End If

    If CDate(dtFrom.Value) > CDate(dtTo.Value) Then
        MessagePop InfoFriend, "System Message", "Invalid date range Date from is greater than date to."
        Exit Sub
    End If

    If txtCode.Text = "" Then
        MessagePop InfoFriend, "System Message", "Entity name not yet selected.Please select entity name."
        Exit Sub
    End If
    Call LOAD_AR_ENTITY
End Sub

Private Sub Form_Load()
'UPDATED BY: JUN
'DATE UPDATED: 02132010
'DESCRIPTION: THIS CUSTOMER AR LEDGER WAS BASE ON AMIS_AR AND AMIS_DETAIL TABLE
'             WHICH AMIS_AR IS THE AMOUNT TO PAY BY THE CUSTOMER AND AMIS_DETAIL IS THE AMOUNT PAID DETAIL BY THE CUSTOMER
    CenterMe frmMain, Me, 1
    Call INIT_CTRL_ENTITY
    Call INIT_CTRL_LEDGER
    Call LOAD_ACCT_CODE
    Call LOAD_MAX_MIN_DATE
    Call LOAD_CUSTOMER
    cboAR_ACCT_CODE.ListIndex = 0
    Call rsRefresh
    Call StoreMemVars
End Sub

Sub LOAD_ACCT_CODE()
    Dim rsLOAD_ACCT_CODE                               As ADODB.Recordset
    Set rsLOAD_ACCT_CODE = New ADODB.Recordset
    rsLOAD_ACCT_CODE.Open "SELECT DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE LEFT(ACCTCODE,5) IN ('11-02','11-03')", gconDMIS, adOpenKeyset
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
    On Error Resume Next
    Dim rsMAX_MIN_DATE                                 As ADODB.Recordset
    Set rsMAX_MIN_DATE = New ADODB.Recordset
    rsMAX_MIN_DATE.Open "SELECT * FROM (SELECT MAX(JDATE)AS MAX_JDATE, MIN(JDATE) AS MIN_JDATE FROM AMIS_JOURNAL_HD WHERE STATUS='P') T WHERE MAX_JDATE IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsMAX_MIN_DATE.EOF And Not rsMAX_MIN_DATE.BOF Then
        dtFrom.Value = Null2String(rsMAX_MIN_DATE!MIN_JDATE)
        dtTo.Value = Null2String(rsMAX_MIN_DATE!MAX_JDATE)
    Else
        dtFrom.Value = LOGDATE
        dtTo.Value = LOGDATE
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
    Call LOAD_MAX_MIN_DATE
    cmdOk_Click
End Sub

Sub LOAD_CUSTOMER()
    Set rsLOAD_CUSTOMER = New ADODB.Recordset
    rsLOAD_CUSTOMER.Open "SELECT DISTINCT TOP 22 CUSTOMERCODE,CUSTOMERNAME FROM AMIS_AR WHERE CUSTOMERCODE IS NOT NULL ORDER BY CUSTOMERNAME ASC", gconDMIS, adOpenKeyset

    rptENTITY.Records.DeleteAll

    If Not rsLOAD_CUSTOMER.EOF And Not rsLOAD_CUSTOMER.BOF Then
        Do While Not rsLOAD_CUSTOMER.EOF
            Set REC = rptENTITY.Records.Add
            REC.AddItem (Trim(UCase(Null2String(rsLOAD_CUSTOMER!CUSTOMERNAME))))
            REC.AddItem (Trim(UCase(Null2String(rsLOAD_CUSTOMER!CustomerCode))))
            rptENTITY.Populate
            Set REC = Nothing
            rsLOAD_CUSTOMER.MoveNext
        Loop
    End If
    Set rsLOAD_CUSTOMER = Nothing
End Sub

Private Sub rptLEDGER_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record(4).Value <> "0.00" And RTrim(LTrim(Row.Record(2).Value)) <> "TOTAL" Then
        Metrics.BackColor = vbWhite
    Else
        Metrics.BackColor = RGB(214, 234, 246)
        '&HFFFFC0
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
    ElseIf xJOURNAL = "DRJ" Then
        Call frmAMISJournalEntry_DRJ.LOADJOURNAL("DRJ")
        FormExistsShow frmAMISJournalEntry_DRJ
        Call frmAMISJournalEntry_DRJ.SearchVoucherNo(xVOUCHER)
    ElseIf xJOURNAL = "GJ" Then
        Call frmAMISJournalEntry_GJ.LOADJOURNAL("GJ")
        FormExistsShow frmAMISJournalEntry_GJ
        Call frmAMISJournalEntry_GJ.SearchVoucherNo(xVOUCHER)
    ElseIf xJOURNAL = "COB" Then
        'JOURNALTYPE = xJOURNAL
        Call frmAMISCustomerAROpening.LOADJOURNAL("COB")
        FormExistsShow frmAMISCustomerAROpening
        Call frmAMISCustomerAROpening.SearchVoucherNo(xVOUCHER)
    End If
End Sub

Private Sub textSearch_Change()
    Dim rssearch                                       As ADODB.Recordset
    Set rssearch = New ADODB.Recordset
    If TextSearch.Text <> "" Then
        rssearch.Open "SELECT DISTINCT CUSTOMERCODE,CUSTOMERNAME FROM AMIS_AR WHERE CUSTOMERCODE IS NOT NULL AND CUSTOMERNAME LIKE '" & Replace(TextSearch.Text, "'", "") & "%' ORDER BY CUSTOMERNAME ASC", gconDMIS, adOpenKeyset
    Else
        rssearch.Open "SELECT DISTINCT TOP 22 CUSTOMERCODE,CUSTOMERNAME FROM AMIS_AR WHERE CUSTOMERCODE IS NOT NULL ORDER BY CUSTOMERNAME ASC", gconDMIS, adOpenKeyset
    End If

    rptENTITY.Records.DeleteAll
    If Not rssearch.EOF And Not rssearch.BOF Then
        Do While Not rssearch.EOF
            Set REC = rptENTITY.Records.Add
            REC.AddItem (Trim(UCase(Null2String(rssearch!CUSTOMERNAME))))
            REC.AddItem (Trim(UCase(Null2String(rssearch!CustomerCode))))
            rptENTITY.Populate
            Set REC = Nothing
            rssearch.MoveNext
        Loop
    End If
    Set rssearch = Nothing
End Sub

Sub LOAD_AR_ENTITY()
    Screen.MousePointer = 11
    Dim xSJVOUCHERNO                                   As String
    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0: xBALANCE = 0
    Set rsLOAD_AR = New ADODB.Recordset
    If cboAR_ACCT_CODE.Text = "ALL" Then
        rsLOAD_AR.Open "SELECT JDATE,SJVOUCHERNO,INVOICETYPE + '-' + INVOICENO AS INVOICE,AMOUNT_TOPAY,AMOUNT_PAID,INVOICENO,INVOICETYPE,CUSTOMERCODE,ACCOUNT_CODE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO FROM AMIS_AR AR INNER JOIN AMIS_CHARTACCOUNT AC ON AR.ACCOUNT_CODE=AC.ACCTCODE WHERE STATUS='P' AND IS_SCHEDULE_ACCNT=1 AND CUSTOMERCODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "' ORDER BY JDATE,SJVOUCHERNO", gconDMIS, adOpenKeyset
    Else
        rsLOAD_AR.Open "SELECT JDATE,SJVOUCHERNO,INVOICETYPE + '-' + INVOICENO AS INVOICE,AMOUNT_TOPAY,AMOUNT_PAID,INVOICENO,INVOICETYPE,CUSTOMERCODE,ACCOUNT_CODE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO FROM AMIS_AR AR INNER JOIN AMIS_CHARTACCOUNT AC ON AR.ACCOUNT_CODE=AC.ACCTCODE WHERE STATUS='P' AND IS_SCHEDULE_ACCNT=1 AND CUSTOMERCODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "' AND " & _
                       "ACCOUNT_CODE = (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "') ORDER BY JDATE,SJVOUCHERNO", gconDMIS, adOpenKeyset
    End If

    rptLEDGER.Records.DeleteAll

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

    If Not rsLOAD_AR.EOF And Not rsLOAD_AR.BOF Then
        Do While Not rsLOAD_AR.EOF
            If Len(rsLOAD_AR!SJVOUCHERNO) = 10 Then
                xSJVOUCHERNO = Left(Null2String(rsLOAD_AR!SJVOUCHERNO), 3)
            Else
                xSJVOUCHERNO = Left(Null2String(rsLOAD_AR!SJVOUCHERNO), 2)
            End If
            If xSJVOUCHERNO = "COB" Or xSJVOUCHERNO = "SJ" Then
                Set REC = rptLEDGER.Records.Add
                REC.AddItem (Trim(Null2String(rsLOAD_AR!JDATE)))
                REC.AddItem (Trim(Null2String(rsLOAD_AR!SJVOUCHERNO)))
                REC.AddItem (Trim(Null2String(rsLOAD_AR!INVOICE)))
                If NumericVal(rsLOAD_AR!AMOUNT_TOPAY) <> 0 Then
                    REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsLOAD_AR!AMOUNT_TOPAY))))
                    REC.AddItem (Trim("0.00"))

                    TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))

                    REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
                Else
                    REC.AddItem (Trim("0.00"))
                    REC.AddItem (Trim(Null2String(rsLOAD_AR!AMOUNT_PAID)))

                    TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_AR!AMOUNT_PAID)), 2))

                    REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
                End If


                Call REFERENCE_INVOICE(Null2String(rsLOAD_AR!INVOICENO), Null2String(rsLOAD_AR!InvoiceType), txtCode.Text, Null2String(rsLOAD_AR!Account_code))
                Call ADJUSTMENT_BYVOUHCHERNO(Null2String(rsLOAD_AR!VOUCHERNO), xSJVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AR!Account_code))
                'Call ADJUSTMENT_DETAILS(Null2String(rsLOAD_AR!VOUCHERNO), xSJVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AR!Account_code))
                rptLEDGER.Populate
                Set REC = Nothing
            ElseIf xSJVOUCHERNO = "APJ" Or xSJVOUCHERNO = "CDJ" Then
                Set REC = rptLEDGER.Records.Add
                REC.AddItem (Trim(Null2String(rsLOAD_AR!JDATE)))
                REC.AddItem (Trim(Null2String(rsLOAD_AR!SJVOUCHERNO)))
                REC.AddItem (Trim(Null2String(rsLOAD_AR!INVOICE)))
                If NumericVal(rsLOAD_AR!AMOUNT_TOPAY) <> 0 Then
                    REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsLOAD_AR!AMOUNT_TOPAY))))
                    REC.AddItem (Trim("0.00"))

                    TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))

                    REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
                Else
                    REC.AddItem (Trim("0.00"))
                    REC.AddItem (Trim(Null2String(ToDoubleNumber(NumericVal(rsLOAD_AR!AMOUNT_PAID)))))

                    TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_AR!AMOUNT_PAID)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_AR!AMOUNT_PAID)), 2))

                    REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
                End If


                Call REFERENCE_INVOICE(Right(Null2String(rsLOAD_AR!SJVOUCHERNO), 6), Left(Null2String(rsLOAD_AR!SJVOUCHERNO), 3), txtCode.Text, Null2String(rsLOAD_AR!Account_code))
                Call ADJUSTMENT_BYVOUHCHERNO(Null2String(rsLOAD_AR!VOUCHERNO), xSJVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AR!Account_code))
                rptLEDGER.Populate
                Set REC = Nothing
            End If
            rsLOAD_AR.MoveNext
        Loop
    End If

    rptLEDGER.Columns(3).FooterText = ToDoubleNumber(TOTAL_DEBIT)
    rptLEDGER.Columns(4).FooterText = ToDoubleNumber(TOTAL_CREDIT)
    rptLEDGER.Columns(5).FooterText = ToDoubleNumber(Round(NumericVal(xBALANCE + FWD_BALANCE), 2))
    Screen.MousePointer = 0
    'Set rsLOAD_AR = Nothing
End Sub

Sub LOAD_AR_ENTITY_PRINTING()
    Dim xSJVOUCHERNO                                   As String
    If Len(Dir(AMIS_REPORT_PATH & "\Ledgers\CustomersSubsidiaryLedger.xlt")) = 0 Then
        MsgBox "No Excel file found", vbInformation, "System Message"
        Exit Sub
    End If
    xCounter = 10
    Set xlApplication = CreateObject("Excel.Application")
    Set xlWorkbook = xlApplication.Workbooks.Open(AMIS_REPORT_PATH & "\Ledgers\CustomersSubsidiaryLedger.xlt")
    Set xlWorksheet = xlWorkbook.Worksheets(1)
    xlWorksheet.Cells(1, "A") = COMPANY_NAME
    xlWorksheet.Cells(1, "A").Font.Bold = True
    xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
    xlWorksheet.Cells(2, "A").Font.Bold = True
    xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
    xlWorksheet.Cells(3, "A").Font.Bold = True
    xlWorksheet.Cells(4, "A") = "CUSTOMER SUBSIDIARY LEDGER"
    xlWorksheet.Cells(4, "A").Font.Bold = True
    xlWorksheet.Cells(6, "A") = "Customer Code: " & txtCode.Text
    xlWorksheet.Cells(6, "A").Font.Bold = True
    xlWorksheet.Cells(7, "A") = "Customer Name: " & txtName.Text
    xlWorksheet.Cells(7, "A").Font.Bold = True
    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0: xBALANCE = 0
    Set rsLOAD_AR = New ADODB.Recordset
    If cboAR_ACCT_CODE.Text = "ALL" Then
        rsLOAD_AR.Open "SELECT JDATE,SJVOUCHERNO,INVOICETYPE + '-' + INVOICENO AS INVOICE,AMOUNT_TOPAY,AMOUNT_PAID,INVOICENO,INVOICETYPE,CUSTOMERCODE,ACCOUNT_CODE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO FROM AMIS_AR AR INNER JOIN AMIS_CHARTACCOUNT AC ON AC.ACCTCODE=AR.ACCOUNT_CODE WHERE IS_SCHEDULE_ACCNT=1 AND CUSTOMERCODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "' ORDER BY JDATE,SJVOUCHERNO", gconDMIS, adOpenKeyset
    Else
        rsLOAD_AR.Open "SELECT JDATE,SJVOUCHERNO,INVOICETYPE + '-' + INVOICENO AS INVOICE,AMOUNT_TOPAY,AMOUNT_PAID,INVOICENO,INVOICETYPE,CUSTOMERCODE,ACCOUNT_CODE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO FROM AMIS_AR AR INNER JOIN AMIS_CHARTACCOUNT AC ON AC.ACCTCODE=AR.ACCOUNT_CODE WHERE IS_SCHEDULE_ACCNT=1 AND CUSTOMERCODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "' AND " & _
                       "ACCOUNT_CODE = (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "') ORDER BY JDATE,SJVOUCHERNO", gconDMIS, adOpenKeyset
    End If

    Call FORWARDED_BALANCE

    xlWorksheet.Cells(xCounter, "A") = Format(dtFrom.Value, "mm/dd/yyyy")
    xlWorksheet.Cells(xCounter, "B") = "FWD BALANCE"
    xlWorksheet.Cells(xCounter, "D") = "0.00"
    xlWorksheet.Cells(xCounter, "E") = "0.00"
    xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(FWD_BALANCE)))

    If Not rsLOAD_AR.EOF And Not rsLOAD_AR.BOF Then
        Do While Not rsLOAD_AR.EOF
            If Len(rsLOAD_AR!SJVOUCHERNO) = 10 Then
                xSJVOUCHERNO = Left(Null2String(rsLOAD_AR!SJVOUCHERNO), 3)
            Else
                xSJVOUCHERNO = Left(Null2String(rsLOAD_AR!SJVOUCHERNO), 2)
            End If
            If xSJVOUCHERNO = "COB" Or xSJVOUCHERNO = "SJ" Then
                xCounter = xCounter + 1
                xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsLOAD_AR!JDATE))), "mm/dd/yyyy")
                xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_AR!SJVOUCHERNO)))
                xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsLOAD_AR!INVOICE)))
                If NumericVal(rsLOAD_AR!AMOUNT_TOPAY) <> 0 Then
                    xlWorksheet.Cells(xCounter, "D") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_AR!AMOUNT_TOPAY))))
                    xlWorksheet.Cells(xCounter, "E") = (Trim("0.00"))

                    TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))

                    xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
                Else
                    xlWorksheet.Cells(xCounter, "D") = (Trim("0.00"))
                    xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_AR!AMOUNT_PAID)))

                    TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_AR!AMOUNT_PAID)), 2))

                    xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
                End If

                Call REFERENCE_INVOICE_PRINTING(Null2String(rsLOAD_AR!INVOICENO), Null2String(rsLOAD_AR!InvoiceType), txtCode.Text, Null2String(rsLOAD_AR!Account_code))
                Call ADJUSTMENT_BYVOUHCHERNO_PRINTING(Null2String(rsLOAD_AR!VOUCHERNO), xSJVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AR!Account_code))
                'Call ADJUSTMENT_DETAILS_PRINTING(Null2String(rsLOAD_AR!VOUCHERNO), xSJVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AR!Account_code))
            ElseIf xSJVOUCHERNO = "APJ" Or xSJVOUCHERNO = "CDJ" Then
                xCounter = xCounter + 1
                xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsLOAD_AR!JDATE))), "mm/dd/yyyy")
                xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_AR!SJVOUCHERNO)))
                xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsLOAD_AR!INVOICE)))
                If NumericVal(rsLOAD_AR!AMOUNT_TOPAY) <> 0 Then
                    xlWorksheet.Cells(xCounter, "D") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_AR!AMOUNT_TOPAY))))
                    xlWorksheet.Cells(xCounter, "E") = (Trim("0.00"))

                    TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))

                    xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
                Else
                    xlWorksheet.Cells(xCounter, "D") = (Trim("0.00"))
                    xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_AR!AMOUNT_PAID)))

                    TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                    xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_AR!AMOUNT_PAID)), 2))

                    xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
                End If

                Call REFERENCE_INVOICE_PRINTING(Right(Null2String(rsLOAD_AR!SJVOUCHERNO), 6), Left(Null2String(rsLOAD_AR!SJVOUCHERNO), 3), txtCode.Text, Null2String(rsLOAD_AR!Account_code))
                Call ADJUSTMENT_BYVOUHCHERNO_PRINTING(Null2String(rsLOAD_AR!VOUCHERNO), xSJVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AR!Account_code))
            End If
            rsLOAD_AR.MoveNext
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
    Set rsLOAD_AR = Nothing
End Sub

Sub REFERENCE_INVOICE(xInvoiceNo As String, xInvoiceType As String, xCUSCODE As String, xAcctCode As String)
    Dim rsINVOICE                                      As ADODB.Recordset
    Dim xVOUCHERNO_ADJ                                 As String
    Dim xJTYPE_ADJ                                     As String
    Set rsINVOICE = New ADODB.Recordset
    rsINVOICE.Open "SELECT * FROM AMIS_DETAIL WHERE STATUS='P' AND INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "' AND CUSTOMERCODE = '" & xCUSCODE & "' AND ACCT_CODE = '" & xAcctCode & "' " & _
                   "AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "'", gconDMIS, adOpenKeyset
    If Not rsINVOICE.EOF And Not rsINVOICE.BOF Then
        Do While Not rsINVOICE.EOF
            Set REC = rptLEDGER.Records.Add
            REC.AddItem (Trim(Null2String(rsINVOICE!JDATE)))
            REC.AddItem (Trim(Null2String(rsINVOICE!jtype) & "-" & Null2String(rsINVOICE!VOUCHERNO)))
            REC.AddItem (Trim(Null2String(rsINVOICE!InvoiceType) & "-" & Null2String(rsINVOICE!INVOICENO)))
            xVOUCHERNO_ADJ = Null2String(rsINVOICE!VOUCHERNO)
            xJTYPE_ADJ = Null2String(rsINVOICE!jtype)
            '            If Trim(Null2String(rsINVOICE!jtype)) = "CRJ" Then
            '                REC.AddItem "OR#-" & (Trim(GET_ORNUM(Null2String(rsINVOICE!jtype), Null2String(rsINVOICE!VOUCHERNO)))) & "/" & (Trim(Null2String(rsINVOICE!InvoiceType) & "-" & Null2String(rsINVOICE!INVOICENO)))
            '            Else
            '                REC.AddItem (Trim(Null2String(rsINVOICE!InvoiceType) & "-" & Null2String(rsINVOICE!INVOICENO)))
            '            End If
            REC.AddItem (Trim("0.00"))
            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsINVOICE!invoiceamount))))

            TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsINVOICE!invoiceamount)), 2))
            xBALANCE = Trim(((xBALANCE) - (rsINVOICE!invoiceamount)))

            REC.AddItem (Trim(ToDoubleNumber(NumericVal(xBALANCE))))
            rptLEDGER.Populate
            Set REC = Nothing
            rsINVOICE.MoveNext
        Loop
        Call ADJUSTMENT_BYVOUHCHERNO(xVOUCHERNO_ADJ, xJTYPE_ADJ, txtCode.Text, xAcctCode)
    End If
    Set rsINVOICE = Nothing
End Sub

Function GET_ORNUM(XXX As String, YYY As String) As String
    Dim rsJournalHD                                    As ADODB.Recordset
    Set rsJournalHD = New ADODB.Recordset
    rsJournalHD.Open "SELECT * FROM AMIS_JOURNAL_HD WHERE JTYPE='" & XXX & "' AND VOUCHERNO='" & YYY & "' AND STATUS='P'", gconDMIS, adOpenKeyset
    If Not rsJournalHD.EOF And Not rsJournalHD.BOF Then
        GET_ORNUM = Null2String(rsJournalHD!INVOICENO)
    Else
        GET_ORNUM = ""
    End If
    Set rsJournalHD = Nothing
End Function

Sub REFERENCE_INVOICE_PRINTING(xInvoiceNo As String, xInvoiceType As String, xCUSCODE As String, xAcctCode As String)
    Dim rsINVOICE                                      As ADODB.Recordset
    Dim xVOUCHERNO_ADJ                                 As String
    Dim xJTYPE_ADJ                                     As String
    Set rsINVOICE = New ADODB.Recordset
    rsINVOICE.Open "SELECT * FROM AMIS_DETAIL WHERE STATUS='P' AND INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "' AND CUSTOMERCODE = '" & xCUSCODE & "' AND ACCT_CODE = '" & xAcctCode & "' " & _
                   "AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "'", gconDMIS, adOpenKeyset
    If Not rsINVOICE.EOF And Not rsINVOICE.BOF Then
        Do While Not rsINVOICE.EOF
            xCounter = xCounter + 1
            xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsINVOICE!JDATE))), "mm/dd/yyyy")
            If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                If Trim(Null2String(rsINVOICE!jtype)) = "CRJ" Then
                    xlWorksheet.Cells(xCounter, "B") = "OR#-" & (Trim(GET_ORNUM(Null2String(rsINVOICE!jtype), Null2String(rsINVOICE!VOUCHERNO))))
                Else
                    xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsINVOICE!jtype) & "-" & Null2String(rsINVOICE!VOUCHERNO)))
                End If
            Else
                xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsINVOICE!jtype) & "-" & Null2String(rsINVOICE!VOUCHERNO)))
            End If
            xVOUCHERNO_ADJ = Null2String(rsINVOICE!VOUCHERNO)
            xJTYPE_ADJ = Null2String(rsINVOICE!jtype)
            '            If Trim(Null2String(rsINVOICE!jtype)) = "CRJ" Then
            '                xlWorksheet.Cells(xCounter, "C") = "OR#-" & (Trim(GET_ORNUM(Null2String(rsINVOICE!jtype), Null2String(rsINVOICE!VOUCHERNO))))
            '            Else
            xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsINVOICE!InvoiceType) & "-" & Null2String(rsINVOICE!INVOICENO)))
            '            End If
            xlWorksheet.Cells(xCounter, "D") = (Trim("0.00"))
            xlWorksheet.Cells(xCounter, "E") = (Trim(ToDoubleNumber(NumericVal(rsINVOICE!invoiceamount))))

            TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsINVOICE!invoiceamount)), 2))
            xBALANCE = Trim(ToDoubleNumber(xBALANCE - NumericVal(rsINVOICE!invoiceamount)))

            xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
            rsINVOICE.MoveNext
        Loop
        Call ADJUSTMENT_BYVOUHCHERNO_PRINTING(xVOUCHERNO_ADJ, xJTYPE_ADJ, txtCode.Text, xAcctCode)
    End If
    Set rsINVOICE = Nothing
End Sub

Sub ADJUSTMENT_BYVOUHCHERNO(xVOUCHERNO As String, xJType As String, xCUSCDE As String, xACCT_CODE As String)
    Dim rsADJ                                          As ADODB.Recordset
    Set rsADJ = New ADODB.Recordset
    rsADJ.Open "SELECT JDATE, SJVOUCHERNO, INVOICETYPE + '-' + INVOICENO AS REF_INVOICE,AMOUNT_TOPAY,AMOUNT_PAID " & _
               "FROM AMIS_AR WHERE CUSTOMERCODE = '" & xCUSCDE & "' AND INVOICETYPE = '" & xJType & "' AND " & _
               "Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "' AND  INVOICENO = '" & xVOUCHERNO & "' AND ACCOUNT_CODE = '" & xACCT_CODE & "' AND LEFT(SJVOUCHERNO,2) = 'GJ'", gconDMIS, adOpenKeyset
    If Not rsADJ.EOF And Not rsADJ.BOF Then
        Do While Not rsADJ.EOF
            Set REC = rptLEDGER.Records.Add
            REC.AddItem (Trim(Null2String(rsADJ!JDATE)))
            REC.AddItem (Trim(Null2String(rsADJ!SJVOUCHERNO)))
            REC.AddItem (Trim(Null2String(rsADJ!REF_INVOICE)))

            If NumericVal(rsADJ!AMOUNT_TOPAY) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsADJ!AMOUNT_TOPAY))))
                REC.AddItem (Trim("0.00"))

                TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsADJ!AMOUNT_TOPAY)), 2))
                xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsADJ!AMOUNT_TOPAY)), 2))

                REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
            Else
                REC.AddItem (Trim("0.00"))
                REC.AddItem (Trim(Null2String(rsADJ!AMOUNT_PAID)))

                TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsADJ!AMOUNT_PAID)), 2))
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

Sub ADJUSTMENT_DETAILS(xVOUCHERNO As String, xJType As String, xCUSCODE As String, xACCT_CODE As String)
    Dim rsINVOICE                                      As ADODB.Recordset
    Set rsINVOICE = New ADODB.Recordset
    rsINVOICE.Open "SELECT * FROM AMIS_DETAIL WHERE STATUS='P' AND INVOICENO = '" & xVOUCHERNO & "' AND INVOICETYPE = '" & xJType & "' AND CUSTOMERCODE = '" & xCUSCODE & "' AND ACCT_CODE = '" & xACCT_CODE & "' " & _
                   "AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "'", gconDMIS, adOpenKeyset
    If Not rsINVOICE.EOF And Not rsINVOICE.BOF Then
        Do While Not rsINVOICE.EOF
            Set REC = rptLEDGER.Records.Add
            REC.AddItem (Trim(Null2String(rsINVOICE!JDATE)))
            REC.AddItem (Trim(Null2String(rsINVOICE!jtype) & "-" & Null2String(rsINVOICE!VOUCHERNO)))
            REC.AddItem (Trim(Null2String(rsINVOICE!InvoiceType) & "-" & Null2String(rsINVOICE!INVOICENO)))
            '            If Trim(Null2String(rsINVOICE!jtype)) = "CRJ" Then
            '                REC.AddItem "OR#-" & (Trim(GET_ORNUM(Null2String(rsINVOICE!jtype), Null2String(rsINVOICE!VOUCHERNO)))) & "/" & (Trim(Null2String(rsINVOICE!InvoiceType) & "-" & Null2String(rsINVOICE!INVOICENO)))
            '            Else
            '                REC.AddItem (Trim(Null2String(rsINVOICE!InvoiceType) & "-" & Null2String(rsINVOICE!INVOICENO)))
            '            End If
            REC.AddItem (Trim("0.00"))
            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsINVOICE!invoiceamount))))

            TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsINVOICE!invoiceamount)), 2))
            xBALANCE = Trim(ToDoubleNumber(xBALANCE - NumericVal(rsINVOICE!invoiceamount)))

            REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
            rptLEDGER.Populate
            Set REC = Nothing
            rsINVOICE.MoveNext
        Loop
        'Call ADJUSTMENT_BYVOUHCHERNO_PRINTING(xVOUCHERNO, xJType, xCUSCODE, xACCT_CODE)
    End If
    Set rsINVOICE = Nothing
End Sub

Sub ADJUSTMENT_DETAILS_PRINTING(xVOUCHERNO As String, xJType As String, xCUSCODE As String, xACCT_CODE As String)
    Dim rsINVOICE                                      As ADODB.Recordset
    Set rsINVOICE = New ADODB.Recordset
    rsINVOICE.Open "SELECT * FROM AMIS_DETAIL WHERE STATUS='P' AND INVOICENO = '" & xVOUCHERNO & "' AND INVOICETYPE = '" & xJType & "' AND CUSTOMERCODE = '" & xCUSCODE & "' AND ACCT_CODE = '" & xACCT_CODE & "' " & _
                   "AND Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "'", gconDMIS, adOpenKeyset
    If Not rsINVOICE.EOF And Not rsINVOICE.BOF Then
        '        Do While Not rsINVOICE.EOF
        '            Set REC = rptLEDGER.Records.Add
        '            REC.AddItem (Trim(Null2String(rsINVOICE!JDate)))
        '            REC.AddItem (Trim(Null2String(rsINVOICE!jtype) & "-" & Null2String(rsINVOICE!VOUCHERNO)))
        '            REC.AddItem (Trim(Null2String(rsINVOICE!InvoiceType) & "-" & Null2String(rsINVOICE!INVOICENO)))
        ''            If Trim(Null2String(rsINVOICE!jtype)) = "CRJ" Then
        ''                REC.AddItem "OR#-" & (Trim(GET_ORNUM(Null2String(rsINVOICE!jtype), Null2String(rsINVOICE!VOUCHERNO)))) & "/" & (Trim(Null2String(rsINVOICE!InvoiceType) & "-" & Null2String(rsINVOICE!INVOICENO)))
        ''            Else
        ''                REC.AddItem (Trim(Null2String(rsINVOICE!InvoiceType) & "-" & Null2String(rsINVOICE!INVOICENO)))
        ''            End If
        '            REC.AddItem (Trim("0.00"))
        '            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsINVOICE!invoiceamount))))
        '
        '            TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsINVOICE!invoiceamount)), 2))
        '            xBALANCE = Trim(ToDoubleNumber(xBALANCE - NumericVal(rsINVOICE!invoiceamount)))
        '
        '            REC.AddItem (Trim(ToDoubleNumber(xBALANCE)))
        '            rptLEDGER.Populate
        '            Set REC = Nothing
        '            rsINVOICE.MoveNext
        '        Loop
        '============================================================
        Do While Not rsINVOICE.EOF
            xCounter = xCounter + 1
            xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsLOAD_AR!JDATE))), "mm/dd/yyyy")
            If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                If Trim(Null2String(rsINVOICE!jtype)) = "CRJ" Then
                    xlWorksheet.Cells(xCounter, "B") = "OR#-" & (Trim(GET_ORNUM(Null2String(rsINVOICE!jtype), Null2String(rsINVOICE!VOUCHERNO))))
                Else
                    xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsINVOICE!jtype) & "-" & Null2String(rsINVOICE!VOUCHERNO)))
                End If
            Else
                xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsINVOICE!jtype) & "-" & Null2String(rsINVOICE!VOUCHERNO)))
            End If
            '            If Trim(Null2String(rsINVOICE!jtype)) = "CRJ" Then
            '                xlWorksheet.Cells(xCounter, "C") = "OR#-" & (Trim(GET_ORNUM(Null2String(rsINVOICE!jtype), Null2String(rsINVOICE!VOUCHERNO))))
            '            Else
            xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsINVOICE!InvoiceType) & "-" & Null2String(rsINVOICE!INVOICENO)))
            '            End If
            xlWorksheet.Cells(xCounter, "D") = (Trim("0.00"))
            xlWorksheet.Cells(xCounter, "E") = (Trim(ToDoubleNumber(NumericVal(rsINVOICE!invoiceamount))))

            TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsINVOICE!invoiceamount)), 2))
            xBALANCE = Trim(ToDoubleNumber(xBALANCE - NumericVal(rsINVOICE!invoiceamount)))

            xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(xBALANCE)))
            rsINVOICE.MoveNext
        Loop
        Call ADJUSTMENT_BYVOUHCHERNO_PRINTING(xVOUCHERNO, xJType, xCUSCODE, xACCT_CODE)
    End If
    Set rsINVOICE = Nothing
End Sub

Sub ADJUSTMENT_BYVOUHCHERNO_PRINTING(xVOUCHERNO As String, xJType As String, xCUSCDE As String, xACCT_CODE As String)
    Dim rsADJ                                          As ADODB.Recordset
    Set rsADJ = New ADODB.Recordset
    rsADJ.Open "SELECT JDATE, SJVOUCHERNO, INVOICETYPE + '-' + INVOICENO AS REF_INVOICE,AMOUNT_TOPAY,AMOUNT_PAID " & _
               "FROM AMIS_AR WHERE STATUS='P' AND CUSTOMERCODE = '" & xCUSCDE & "' AND INVOICETYPE = '" & xJType & "' AND " & _
               "Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTo.Value & "' AND  INVOICENO = '" & xVOUCHERNO & "' AND ACCOUNT_CODE = '" & xACCT_CODE & "' AND LEFT(SJVOUCHERNO,2) = 'GJ'", gconDMIS, adOpenKeyset
    'rsADJ.Open "SELECT JDATE, SJVOUCHERNO, INVOICETYPE + '-' + INVOICENO AS REF_INVOICE,AMOUNT_TOPAY,AMOUNT_PAID " & _
     "FROM AMIS_AR WHERE CUSTOMERCODE = '" & xCUSCDE & "' AND INVOICETYPE = '" & xJtype & "' AND " & _
     "Jdate >= '" & dtFrom.Value & "'and Jdate <= '" & dtTO.Value & "' AND  INVOICENO = '" & XVOUCHERNO & "' AND ACCOUNT_CODE = '" & xACCT_CODE & "' AND LEFT(SJVOUCHERNO,2) = '" & xxSJVOUCHERNO & "'", gconDMIS, adOpenKeyset
    If Not rsADJ.EOF And Not rsADJ.BOF Then
        Do While Not rsADJ.EOF
            xCounter = xCounter + 1
            xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsADJ!JDATE))), "mm/dd/yyyy")
            xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsADJ!SJVOUCHERNO)))
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
    Dim rsLOAD_AR                                      As ADODB.Recordset
    Dim xSJVOUCHERNO                                   As String
    Set rsLOAD_AR = New ADODB.Recordset

    FWD_BALANCE = 0: FWD_CREDIT = 0: FWD_DEBIT = 0

    If cboAR_ACCT_CODE.Text = "ALL" Then
        rsLOAD_AR.Open "SELECT JDATE,SJVOUCHERNO,INVOICETYPE + '-' + INVOICENO AS INVOICE,AMOUNT_TOPAY,AMOUNT_PAID,INVOICENO,INVOICETYPE,CUSTOMERCODE,ACCOUNT_CODE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO " & _
                       "FROM AMIS_AR WHERE CUSTOMERCODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate < '" & dtFrom.Value & "'", gconDMIS, adOpenKeyset
    Else
        rsLOAD_AR.Open "SELECT JDATE,SJVOUCHERNO,INVOICETYPE + '-' + INVOICENO,AMOUNT_TOPAY,AMOUNT_PAID,INVOICENO,INVOICETYPE,CUSTOMERCODE,ACCOUNT_CODE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO " & _
                       "FROM AMIS_AR WHERE CUSTOMERCODE = '" & RTrim(LTrim(txtCode.Text)) & "' AND Jdate < '" & dtFrom.Value & "' AND " & _
                       "ACCOUNT_CODE = (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = '" & (RTrim(LTrim(cboAR_ACCT_CODE.Text))) & "')", gconDMIS, adOpenKeyset
    End If

    If Not rsLOAD_AR.EOF And Not rsLOAD_AR.BOF Then
        Do While Not rsLOAD_AR.EOF
            If Len(rsLOAD_AR!SJVOUCHERNO) = 10 Then
                xSJVOUCHERNO = Left(Null2String(rsLOAD_AR!SJVOUCHERNO), 3)
            Else
                xSJVOUCHERNO = Left(Null2String(rsLOAD_AR!SJVOUCHERNO), 2)
            End If
            If xSJVOUCHERNO = "COB" Or xSJVOUCHERNO = "SJ" Then
                If NumericVal(rsLOAD_AR!AMOUNT_TOPAY) <> 0 Then
                    FWD_DEBIT = ToDoubleNumber(Round((FWD_DEBIT + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                    FWD_BALANCE = ToDoubleNumber(Round((FWD_BALANCE + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                Else
                    FWD_CREDIT = ToDoubleNumber(Round((FWD_CREDIT + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                    FWD_BALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_AR!AMOUNT_PAID)), 2))
                End If

                Call FWD_REFERENCE_INVOICE(Null2String(rsLOAD_AR!INVOICENO), Null2String(rsLOAD_AR!InvoiceType), txtCode.Text, Null2String(rsLOAD_AR!Account_code))
                Call FWD_ADJUSTMENT_BYVOUHCHERNO(Null2String(rsLOAD_AR!VOUCHERNO), xSJVOUCHERNO, txtCode.Text, Null2String(rsLOAD_AR!Account_code))

            ElseIf xSJVOUCHERNO = "APJ" Or xSJVOUCHERNO = "CDJ" Then
                If NumericVal(rsLOAD_AR!AMOUNT_TOPAY) <> 0 Then
                    FWD_DEBIT = ToDoubleNumber(Round((FWD_DEBIT + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                    FWD_BALANCE = ToDoubleNumber(Round((FWD_BALANCE + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                Else
                    FWD_CREDIT = ToDoubleNumber(Round((FWD_CREDIT + NumericVal(rsLOAD_AR!AMOUNT_TOPAY)), 2))
                    FWD_BALANCE = ToDoubleNumber(Round((FWD_BALANCE - NumericVal(rsLOAD_AR!AMOUNT_PAID)), 2))
                End If

                Call FWD_REFERENCE_INVOICE(Right(Null2String(rsLOAD_AR!SJVOUCHERNO), 6), Left(Null2String(rsLOAD_AR!SJVOUCHERNO), 3), txtCode.Text, Null2String(rsLOAD_AR!Account_code))
            End If
            rsLOAD_AR.MoveNext
        Loop
    End If
End Sub

Sub FWD_REFERENCE_INVOICE(xInvoiceNo As String, xInvoiceType As String, xCUSCODE As String, xAcctCode As String)
    Dim rsINVOICE                                      As ADODB.Recordset
    Set rsINVOICE = New ADODB.Recordset
    rsINVOICE.Open "SELECT * FROM AMIS_DETAIL WHERE STATUS='P' AND INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "' AND CUSTOMERCODE = '" & xCUSCODE & "' AND ACCT_CODE = '" & xAcctCode & "' " & _
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
               "FROM AMIS_AR WHERE CUSTOMERCODE = '" & xCUSCDE & "' AND INVOICETYPE = '" & xJType & "' AND " & _
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
    rsREF.Open "SELECT DISTINCT CUSTOMERNAME,CUSTOMERCODE FROM AMIS_AR ORDER BY CUSTOMERNAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsREF.EOF And Not rsREF.BOF Then
        txtCode.Text = UCase(Null2String(rsREF!CustomerCode))
        txtName.Text = UCase(Null2String(rsREF!CUSTOMERNAME))
    End If
    Call LOAD_AR_ENTITY
End Sub

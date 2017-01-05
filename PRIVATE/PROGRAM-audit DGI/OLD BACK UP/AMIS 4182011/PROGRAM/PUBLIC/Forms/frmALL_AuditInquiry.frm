VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmALL_AuditInquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audit inquiry"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14310
   DrawWidth       =   10
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmALL_AuditInquiry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "frmALL_AuditInquiry.frx":000C
   ScaleHeight     =   7665
   ScaleWidth      =   14310
   Begin XtremeReportControl.ReportControl rptHIST 
      Height          =   6465
      Left            =   30
      TabIndex        =   18
      Top             =   840
      Width           =   14235
      _Version        =   655364
      _ExtentX        =   25109
      _ExtentY        =   11404
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnReorder=   0   'False
      MultipleSelection=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton cmdPRINT 
      Caption         =   "Print"
      Height          =   405
      Left            =   13080
      TabIndex        =   20
      Top             =   390
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   30
      ScaleHeight     =   285
      ScaleWidth      =   14205
      TabIndex        =   8
      Top             =   7320
      Width           =   14235
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - DISPLAY ALL LOGS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   2
         Left            =   3600
         TabIndex        =   17
         Top             =   30
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " F4 - MORE OPTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   1
         Left            =   1530
         TabIndex        =   11
         Top             =   30
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  ESC - EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   6090
         TabIndex        =   10
         Top             =   30
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  F3 - SEARCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   0
         TabIndex        =   9
         Top             =   30
         Width           =   1065
      End
   End
   Begin VB.TextBox txtSEARCH 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   30
      TabIndex        =   0
      Top             =   390
      Width           =   13005
   End
   Begin VB.PictureBox picOPTION 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2565
      Left            =   3570
      ScaleHeight     =   2535
      ScaleWidth      =   7065
      TabIndex        =   12
      Top             =   2580
      Width           =   7095
      Begin VB.ComboBox cboTRAN 
         Height          =   330
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   5385
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6240
         MousePointer    =   99  'Custom
         Picture         =   "frmALL_AuditInquiry.frx":685E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancel"
         Top             =   1620
         Width           =   705
      End
      Begin VB.ComboBox cboAPP 
         Height          =   330
         ItemData        =   "frmALL_AuditInquiry.frx":6B9C
         Left            =   1590
         List            =   "frmALL_AuditInquiry.frx":6BAF
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   1695
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Display"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5550
         MousePointer    =   99  'Custom
         Picture         =   "frmALL_AuditInquiry.frx":6BD1
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Move to Previous Record"
         Top             =   1620
         Width           =   705
      End
      Begin VB.ComboBox cboModule 
         Height          =   330
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   810
         Width           =   5385
      End
      Begin VB.ComboBox cboModuleType 
         Height          =   330
         Left            =   4830
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   450
         Width           =   2145
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSACTION ID"
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   1
         Top             =   1290
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MODULE NAME"
         Height          =   210
         Index           =   2
         Left            =   390
         TabIndex        =   16
         Top             =   540
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHOOSE MODULE"
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   900
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MODULE TYPE"
         Height          =   210
         Index           =   0
         Left            =   3660
         TabIndex        =   14
         Top             =   540
         Width           =   1065
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   14085
         _Version        =   655364
         _ExtentX        =   24844
         _ExtentY        =   397
         _StockProps     =   14
         Caption         =   "CHOOSE OPTION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   8388608
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption6 
      Height          =   315
      Left            =   30
      TabIndex        =   19
      Top             =   0
      Width           =   14235
      _Version        =   655364
      _ExtentX        =   25109
      _ExtentY        =   556
      _StockProps     =   14
      Caption         =   "Type Your keyword here (if you want to search in Different column, separate it by space key)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "frmALL_AuditInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim COLSTRING                                          As String

Sub DisabledBack(COND As Boolean)
    rptHIST.Enabled = COND
    txtSearch.Enabled = COND
    cmdPrint.Enabled = COND
End Sub

Sub FillTransactionID()
    Dim rsTMP                                          As New ADODB.Recordset

    cboTRAN.Clear
    Set rsTMP = gconAudit.Execute("SELECT TRANNO FROM DMIS_AUDIT WHERE MODULE_NAME = '" & cboModule.Text & "' ORDER BY TRANNO")
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        Do While Not rsTMP.EOF
            cboTRAN.AddItem Null2String(rsTMP!TRANNO)

            rsTMP.MoveNext
        Loop
    End If

    Set rsTMP = Nothing
End Sub

Sub FillCboModules()
    Dim rsTMP                                          As New ADODB.Recordset

    cboModule.Clear
    Set rsTMP = gconDMIS.Execute("SELECT DESCRIPTIONS FROM ALL_RAMS_MODULES WHERE MODULE_TYPE = '" & cboModuleType.Text & "' AND MAINMODULENAME = '" & cboAPP.Text & "' ORDER BY DESCRIPTIONS")
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        Do While Not rsTMP.EOF
            cboModule.AddItem Null2String(rsTMP!DESCRIPTIONS)

            rsTMP.MoveNext
        Loop
    End If

    Set rsTMP = Nothing
End Sub

Sub LoadAllModule()
    Dim rsTMP                                          As New ADODB.Recordset

    cboModuleType.Clear
    Set rsTMP = gconDMIS.Execute("SELECT DISTINCT MODULE_TYPE FROM ALL_RAMS_MODULES WHERE MAINMODULENAME = '" & cboAPP.Text & "' ORDER BY MODULE_TYPE")
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        Do While Not rsTMP.EOF
            cboModuleType.AddItem Null2String(rsTMP!MODULE_TYPE)
            rsTMP.MoveNext
        Loop
    End If
    Set rsTMP = Nothing
End Sub


Private Sub cboAPP_Change()
    Call LoadAllModule
End Sub

Private Sub cboAPP_Click()
    Call LoadAllModule
End Sub

Private Sub cboModule_Change()
    Call FillTransactionID
End Sub

Private Sub cboModule_Click()
    FillTransactionID
End Sub

Private Sub cboModuleType_Change()
    Call FillCboModules
End Sub

Private Sub cboModuleType_Click()
    Call FillCboModules
End Sub

Private Sub cmdCancel_Click()
    picoption.ZOrder 1
    picoption.Visible = False
    Call DisabledBack(True)
End Sub

Private Sub cmdPrint_Click()
    If rptHIST.Records.Count <= 0 Then Exit Sub
    rptHIST.PrintOptions.Header.TextCenter = "AUDIT PRINT FOR " & Replace(frmALL_AuditInquiry.Caption, "Audit Inquiry", "")
    rptHIST.PrintPreview True
End Sub

Private Sub cmdRefresh_Click()
'If cboTRAN.Text = "" Then
'    ShowIsRequiredMsg ("Transaction ID cannot be Blank")
'    cboTRAN.SetFocus
'    Exit Sub
'End If

    Dim rsTMP                                          As New ADODB.Recordset
    Dim VTRANID                                        As String
    Set rsTMP = gconAudit.Execute("SELECT TRANSACTION_ID FROM DMIS_AUDIT WHERE TRANNO = '" & cboTRAN.Text & "' AND MODULE_NAME = '" & cboModule.Text & "'")
    If Not (rsTMP.EOF And rsTMP.BOF) Then
        VTRANID = Null2String(rsTMP!TRANSACTION_ID)
    End If
    Set rsTMP = Nothing

    Call cmdCancel_Click
    Call DisplayHistory("", cboModule.Text)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
    Case vbKeyEscape:
        Unload Me

    Case vbKeyF3:
        txtSearch.SetFocus

    Case vbKeyF4:
        Call DisabledBack(False)
        picoption.ZOrder 0
        picoption.Visible = True
        cboAPP.SetFocus

    Case vbKeyF5:
        MsgBox "Function Under Revision", vbInformation, "DMIS 2.0"
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Null2String(rptHIST.SelectedRows(0).Record(10).Value)
End Sub

Private Sub mnuPaste_Click()

End Sub

Private Sub rptHIST_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.GroupRow = True Then Exit Sub
    If (frmALL_AUDIT_Details.SHOW_AUDITDETAILS(Row.Record(9).Value)) = True Then
        frmALL_AUDIT_Details.ZOrder 0
        frmALL_AUDIT_Details.Show
    End If
End Sub

Private Sub rptHIST_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If LOGCODE = "NET" Then Call PopupMenu(mnuOption)
End Sub

Private Sub txtSearch_Change()
    rptHIST.FilterText = txtSearch.Text
    rptHIST.Populate
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.BackColor = &HC0FFC0
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch.BackColor = vbWhite
End Sub

Public Sub PASS_COLVAL(XXX As String)
    COLSTRING = XXX
End Sub

Public Sub DisplayHistory(xTRANID As String, xMODNAME As String, Optional VTYPE As String)
    Screen.MousePointer = 11
    Dim RSUPLOAD                                       As New ADODB.Recordset
    Dim REC                                            As XtremeReportControl.ReportRecord

    DoEvents
    If VTYPE = "PRINTING" Then
        DoEvents
        'If CHANGE_USER = True Then
        If COMPANY_CODE = COMPANY_VERSION Then
            Set RSUPLOAD = gconAudit.Execute("SELECT User_Name AS UserName,USERACTION,ACTION_DATE,Description,Type,TRANNO,TRANSTYPE,DETID,ID,TRACKING_MEMO FROM ALL_VW_AUDIT WHERE DESCRIPTION  = '" & xMODNAME & "' ORDER BY ACTION_DATE ASC")
        Else
            Set RSUPLOAD = gconAudit.Execute("SELECT UserName,USERACTION,ACTION_DATE,Description,Type,TRANNO,TRANSTYPE,DETID,ID,TRACKING_MEMO FROM ALL_VW_AUDIT WHERE DESCRIPTION  = '" & xMODNAME & "' ORDER BY ACTION_DATE ASC")
        End If
    Else
        'Set RSUPLOAD = gconAudit.Execute("SELECT USERNAME, USERACTION, convert(varchar, ACTION_DATE,101), SUBSTRING(convert(varchar, ACTION_DATE,109),13,LEN(convert(varchar, ACTION_DATE,109)) - 12), DESCRIPTION, TYPE, TRANNO, TRANSTYPE, DETID, ID FROM ALL_VW_AUDIT WHERE DESCRIPTION  = '" & xMODNAME & "' AND TRANSACTION_ID = " & xTRANID & " ORDER BY convert(varchar, ACTION_DATE,08) ASC")
        DoEvents
        'If CHANGE_USER = True Then
        If COMPANY_CODE = COMPANY_VERSION Then
            Set RSUPLOAD = gconAudit.Execute("SELECT User_Name as UserName,USERACTION,ACTION_DATE,Description,Type,TRANNO,TRANSTYPE,DETID,ID,TRACKING_MEMO FROM ALL_VW_AUDIT WHERE DESCRIPTION  = '" & xMODNAME & "' AND TRANSACTION_ID = " & xTRANID & " ORDER BY ACTION_DATE ASC")
        Else
            Set RSUPLOAD = gconAudit.Execute("SELECT UserName,USERACTION,ACTION_DATE,Description,Type,TRANNO,TRANSTYPE,DETID,ID,TRACKING_MEMO FROM ALL_VW_AUDIT WHERE DESCRIPTION  = '" & xMODNAME & "' AND TRANSACTION_ID = " & xTRANID & " ORDER BY ACTION_DATE ASC")
        End If
    End If

    With rptHIST
        .Columns.DeleteAll

        .Columns.Add 0, "USER ID", 70, True:: .Columns(0).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
        .Columns.Add 1, "USER ACTION", 90, True: .Columns(1).Alignment = xtpAlignmentCenter: .Columns(1).AllowRemove = False
        .Columns.Add 2, "DATE", 70, True: .Columns(2).Alignment = xtpAlignmentCenter: .Columns(2).AllowRemove = False
        .Columns.Add 3, "TIME", 90, True: .Columns(3).Alignment = xtpAlignmentCenter: .Columns(3).AllowRemove = False
        .Columns.Add 4, "MODULE NAME", 150, True: .Columns(4).Alignment = xtpAlignmentLeft: .Columns(4).AllowRemove = False
        .Columns.Add 5, "TYPE", 50, True: .Columns(5).Alignment = xtpAlignmentCenter: .Columns(5).AllowRemove = False
        .Columns.Add 6, "TRANNO", 270, True: .Columns(6).Alignment = xtpAlignmentLeft: .Columns(6).AllowRemove = False
        .Columns.Add 7, "TRANSTYPE", 80, True: .Columns(7).Alignment = xtpAlignmentCenter: .Columns(7).AllowRemove = False: .Columns(7).Resizable = False
        .Columns.Add 8, "DET ID", 80, True: .Columns(8).Alignment = xtpAlignmentCenter: .Columns(8).AllowRemove = False: .Columns(8).Resizable = False
        .Columns.Add 9, "ID", 100, True: .Columns(9).Alignment = xtpAlignmentCenter: .Columns(9).AllowRemove = False: .Columns(9).Resizable = False
        .Columns.Add 10, "SQL CODE", 0, True: .Columns(10).Alignment = xtpAlignmentCenter: .Columns(10).AllowRemove = False: .Columns(10).Resizable = False

        .GroupsOrder.Add rptHIST.Columns(1)
        .Columns(1).Visible = False
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        '.PaintManager.TextFont.Bold = True
    End With

    rptHIST.Records.DeleteAll
    While Not RSUPLOAD.EOF
        DoEvents
        Set REC = rptHIST.Records.Add
        With REC
            .AddItem Null2String(RSUPLOAD!UserName)
            .AddItem Null2String(GetUserAction(RSUPLOAD!USERACTION))
            .AddItem Null2String(DateValue(RSUPLOAD!ACTION_DATE))
            .AddItem Null2String(TimeValue(RSUPLOAD!ACTION_DATE))
            .AddItem Null2String(RSUPLOAD!DESCRIPTION)
            .AddItem Null2String(RSUPLOAD!Type)
            .AddItem Null2String(RSUPLOAD!TRANNO)
            .AddItem Null2String(RSUPLOAD!TRANSTYPE)
            .AddItem Null2String(RSUPLOAD!DETID)
            .AddItem Null2String(RSUPLOAD!ID)
            .AddItem Null2String(RSUPLOAD!TRACKING_MEMO)

        End With
        RSUPLOAD.MoveNext
    Wend
    rptHIST.Populate
    Screen.MousePointer = 0
End Sub



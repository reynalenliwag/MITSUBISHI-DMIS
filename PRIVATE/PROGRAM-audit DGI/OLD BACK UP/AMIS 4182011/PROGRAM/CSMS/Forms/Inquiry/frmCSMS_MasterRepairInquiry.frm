VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCSMS_MasterRepairInquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Repair Order Inquiry"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12795
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_MasterRepairInquiry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   12795
   Begin XtremeReportControl.ReportControl rptRO 
      Height          =   4815
      Left            =   30
      TabIndex        =   0
      Top             =   660
      Width           =   12735
      _Version        =   655364
      _ExtentX        =   22463
      _ExtentY        =   8493
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnReorder=   0   'False
      MultipleSelection=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   2355
      Left            =   30
      TabIndex        =   3
      Top             =   5520
      Width           =   12735
      _Version        =   655364
      _ExtentX        =   22463
      _ExtentY        =   4154
      _StockProps     =   64
      Appearance      =   2
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.FixedTabWidth=   160
      ItemCount       =   4
      Item(0).Caption =   "Jobs"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsvJobs"
      Item(1).Caption =   "Parts"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lsvParts"
      Item(2).Caption =   "Materials"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lsvMat"
      Item(3).Caption =   "Accessories"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "lsvAcc"
      Begin MSComctlLib.ListView lsvJobs 
         Height          =   1905
         Left            =   60
         TabIndex        =   4
         Top             =   360
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Job Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Job Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Std Hrs"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Flat Rate"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Discount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lsvParts 
         Height          =   1905
         Left            =   -69940
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Part No"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Part Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Unit Price"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Discount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lsvMat 
         Height          =   1905
         Left            =   -69940
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mat. Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Material Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Unit Price"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Discount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lsvAcc 
         Height          =   1905
         Left            =   -69940
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Acc.Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Acc. Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Unit Price"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Discount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   6075
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12795
      _Version        =   655364
      _ExtentX        =   22569
      _ExtentY        =   979
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
   End
End
Attribute VB_Name = "frmCSMS_MasterRepairInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub InitializeReportControl()
    Screen.MousePointer = 11
    
    With rptRO
        .Columns.DeleteAll
        .Columns.Add 0, "RO NUMBER", 90, True::     .Columns(0).Alignment = xtpAlignmentCenter:      .Columns(0).Resizable = True:   .Columns(0).AllowRemove = False
        .Columns.Add 1, "INV NUMBER", 90, True:    .Columns(1).Alignment = xtpAlignmentCenter:      .Columns(1).Resizable = True:   .Columns(1).AllowRemove = False
        .Columns.Add 2, "CUSTOMER", 170, True:       .Columns(2).Alignment = xtpAlignmentLeft:      .Columns(2).Resizable = True:   .Columns(2).AllowRemove = False
        .Columns.Add 3, "PLATE NO", 60, True:       .Columns(3).Alignment = xtpAlignmentLeft:      .Columns(3).Resizable = True:   .Columns(3).AllowRemove = False
        .Columns.Add 4, "VIN NO", 130, True:         .Columns(4).Alignment = xtpAlignmentLeft:       .Columns(4).Resizable = True:   .Columns(4).AllowRemove = False
        .Columns.Add 5, "MODEL", 90, True:          .Columns(5).Alignment = xtpAlignmentLeft:     .Columns(5).Resizable = True:   .Columns(5).AllowRemove = False
        .Columns.Add 6, "SA NAME", 90, True:        .Columns(6).Alignment = xtpAlignmentLeft:     .Columns(6).Resizable = True:   .Columns(6).AllowRemove = False
        .Columns.Add 7, "KM RDG", 75, True:         .Columns(7).Alignment = xtpAlignmentCenter:     .Columns(7).Resizable = True:   .Columns(7).AllowRemove = False
        .Columns.Add 8, "INV AMT", 80, True:        .Columns(8).Alignment = xtpAlignmentRight:     .Columns(8).Resizable = True:   .Columns(8).AllowRemove = False
        .Columns.Add 9, "INS AMT", 80, True:        .Columns(9).Alignment = xtpAlignmentRight:     .Columns(9).Resizable = True:   .Columns(9).AllowRemove = False
        .Columns.Add 10, "DATE RECD", 80, True:     .Columns(10).Alignment = xtpAlignmentCenter:    .Columns(10).Resizable = True:  .Columns(10).AllowRemove = False
        .Columns.Add 11, "DATE INV", 80, True:      .Columns(11).Alignment = xtpAlignmentCenter:    .Columns(11).Resizable = True:  .Columns(11).AllowRemove = False
        .Columns.Add 12, "DATE REL", 80, True:      .Columns(12).Alignment = xtpAlignmentCenter:    .Columns(12).Resizable = True:  .Columns(12).AllowRemove = False
        
        
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
    End With
    Screen.MousePointer = 0
End Sub

Sub FillGrid()
    Dim RSUPLOAD                                        As New ADODB.Recordset
    Dim REC                                             As XtremeReportControl.ReportRecord
    
    Set RSUPLOAD = gconDMIS.Execute("SELECT " & _
        " CSMS_Repor.REP_OR AS RO, " & _
        " CSMS_Repor.INVOICE AS INV, " & _
        " CSMS_Repor.NIYM AS NIYM, " & _
        " CSMS_Repor.PLATE_NO AS PLATE, " & _
        " CSMS_Repor.VIN AS VIN, " & _
        " CSMS_Repor.MODEL AS MODEL, " & _
        " CSMS_vw_EMPNO.NAYM AS SA, " & _
        " CSMS_Repor.KM_RDG AS KM, " & _
        " CSMS_Repor.RO_AMOUNT AS AMT, " & _
        " CSMS_Repor.INSAMT AS INS, " & _
        " CSMS_Repor.DTE_RECD AS RECD, " & _
        " CSMS_Repor.DTE_COMP AS COMP, " & _
        " CSMS_Repor.DTE_REL AS REL" & _
        " FROM CSMS_Repor INNER JOIN " & _
        " CSMS_vw_EMPNO ON CSMS_Repor.RECD_BY = CSMS_vw_EMPNO.CODE " & _
        " WHERE (CSMS_Repor.TRANSTYPE = 'R')")
    rptRO.Records.DeleteAll
    While Not RSUPLOAD.EOF
        Set REC = rptRO.Records.Add
        REC.AddItem (Trim(RSUPLOAD!ro))
        REC.AddItem (Trim(RSUPLOAD!INV))
        REC.AddItem (Trim(RSUPLOAD!NIYM))
        REC.AddItem (Trim(RSUPLOAD!PLATE))
        REC.AddItem (Trim(RSUPLOAD!VIN))
        REC.AddItem (Trim(RSUPLOAD!MODEL))
        REC.AddItem (Trim(RSUPLOAD!SA))
        REC.AddItem (Trim(RSUPLOAD!KM))
        REC.AddItem (Trim(Format(RSUPLOAD!AMT, MAXIMUM_DIGIT)))
        REC.AddItem (Trim(Format(RSUPLOAD!INS, MAXIMUM_DIGIT)))
        REC.AddItem (Trim(RSUPLOAD!RECD))
        REC.AddItem (Trim(RSUPLOAD!COMP))
        REC.AddItem (Trim(RSUPLOAD!REL))
        
        RSUPLOAD.MoveNext
        Set REC = Nothing
    Wend
    rptRO.Populate
    
    Set RSUPLOAD = Nothing
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    Call InitializeReportControl
    Call FillGrid
End Sub

Private Sub rptRO_SelectionChanged()
    Call FillDetails(Null2String(rptRO.SelectedRows(0).Record(0).Value))
End Sub

Private Sub txtSearch_Change()
    rptRO.FilterText = txtSearch.Text
    rptRO.Populate
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.BackColor = &HC0FFC0
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch.BackColor = vbWhite
End Sub

Sub FillDetails(xRO As String)
    Dim RSTMP               As New ADODB.Recordset
    
    lsvJobs.ListItems.Clear
    Set RSTMP = gconDMIS.Execute("SELECT DETCDE, DETDSC, CAST(DET_HRS AS DECIMAL(18,2)), CAST(FLATRATE AS DECIMAL(18,2)), CAST(DISCOUNT_2 AS DECIMAL(18,2)), CAST(DET_AMT AS DECIMAL(18,2)), ID FROM CSMS_RO_DET " & _
        " WHERE REP_OR = " & N2Str2Null(xRO) & _
        " AND LIVIL = '1'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Call Listview_Loadval(lsvJobs.ListItems, RSTMP)
    End If
    Set RSTMP = Nothing
    
    Set RSTMP = New ADODB.Recordset
    lsvParts.ListItems.Clear
    Set RSTMP = gconDMIS.Execute("SELECT DETCDE, DETDSC, CAST(DETVOL AS DECIMAL(18,2)), CAST(DETPRC AS DECIMAL(18,2)), CAST(DISCOUNT_2 AS DECIMAL(18,2)), CAST(DET_AMT AS DECIMAL(18,2)), ID FROM CSMS_RO_DET " & _
        " WHERE REP_OR = " & N2Str2Null(xRO) & _
        " AND LIVIL = '2'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Call Listview_Loadval(lsvParts.ListItems, RSTMP)
    End If
    Set RSTMP = Nothing

    Set RSTMP = New ADODB.Recordset
    lsvMat.ListItems.Clear
    Set RSTMP = gconDMIS.Execute("SELECT DETCDE, DETDSC, CAST(DETVOL AS DECIMAL(18,2)), CAST(DETPRC AS DECIMAL(18,2)), CAST(DISCOUNT_2 AS DECIMAL(18,2)), CAST(DET_AMT AS DECIMAL(18,2)), ID FROM CSMS_RO_DET " & _
        " WHERE REP_OR = " & N2Str2Null(xRO) & _
        " AND LIVIL = '3'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Call Listview_Loadval(lsvMat.ListItems, RSTMP)
    End If
    Set RSTMP = Nothing
    
    Set RSTMP = New ADODB.Recordset
    lsvAcc.ListItems.Clear
    Set RSTMP = gconDMIS.Execute("SELECT DETCDE, DETDSC, CAST(DETVOL AS DECIMAL(18,2)), CAST(DETPRC AS DECIMAL(18,2)), CAST(DISCOUNT_2 AS DECIMAL(18,2)), CAST(DET_AMT AS DECIMAL(18,2)), ID FROM CSMS_RO_DET " & _
        " WHERE REP_OR = " & N2Str2Null(xRO) & _
        " AND LIVIL = '4'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Call Listview_Loadval(lsvAcc.ListItems, RSTMP)
    End If
    Set RSTMP = Nothing
End Sub

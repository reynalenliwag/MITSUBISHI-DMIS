VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO50BF~1.OCX"
Begin VB.Form frmSearchAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Account Code / Description"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmSearchAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   9090
   Begin VB.Frame Frame1 
      Caption         =   $"frmSearchAccount.frx":09AA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3975
      Left            =   60
      TabIndex        =   1
      Top             =   690
      Width           =   8985
      Begin XtremeReportControl.ReportControl rptAccounts 
         Height          =   3645
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   8835
         _Version        =   655364
         _ExtentX        =   15584
         _ExtentY        =   6429
         _StockProps     =   64
         BorderStyle     =   4
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   1110
      TabIndex        =   0
      Top             =   120
      Width           =   7785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   240
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   110
      TabIndex        =   3
      Top             =   260
      Width           =   945
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   645
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9975
      _Version        =   655364
      _ExtentX        =   17595
      _ExtentY        =   1138
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
      GradientColorLight=   -2147483635
      GradientColorDark=   -2147483629
   End
End
Attribute VB_Name = "frmSearchAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event RECORDSELECTED(strChartAccount As String)

Sub FillGrid(XXX As String)
    Screen.MousePointer = 11
    Dim rsAccounts                                          As ADODB.Recordset
    Dim REC                                                 As XtremeReportControl.ReportRecord
    Set rsAccounts = New ADODB.Recordset
    rsAccounts.Open "select * from AMIS_CHARTACCOUNT where TITLES = '" & XXX & "'", gconDMIS, adOpenKeyset
    rptAccounts.Records.DeleteAll
    While Not rsAccounts.EOF
        Set REC = rptAccounts.Records.Add
        With REC
            .AddItem Null2String(rsAccounts!AcctCode)
            .AddItem Null2String(rsAccounts!DESCRIPTION)
            .AddItem Null2String(rsAccounts!Trantype2)
            .AddItem Null2String(rsAccounts!TRANTYPE1)
            .AddItem Null2String(rsAccounts!Trantype3)
            .AddItem Null2String(rsAccounts!Trantype4)
            DoEvents
        End With
        rsAccounts.MoveNext
    Wend
    rptAccounts.Populate
    Set rsAccount = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    InitializeFill
    Call FillGrid(xSELECTED)
End Sub

Private Sub rptAccounts_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    RaiseEvent RECORDSELECTED(Row.Record(0).Value)
    Unload Me
End Sub

Private Sub txtSearch_Change()
    rptAccounts.FilterText = txtSearch.Text
    rptAccounts.Populate
End Sub

Sub InitializeFill()
    Screen.MousePointer = 11
    With rptAccounts
        .Columns.DeleteAll
        .Columns.Add 0, "Account Code", 200, True: .Columns(0).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
        .Columns.Add 1, "Description", 450, True: .Columns(1).Alignment = xtpAlignmentLeft: .Columns(1).AllowRemove = False
        .Columns.Add 2, "Department", 200, True: .Columns(2).Alignment = xtpAlignmentLeft: .Columns(2).AllowRemove = False
        .Columns.Add 3, "Models", 200, True: .Columns(3).Alignment = xtpAlignmentLeft: .Columns(3).AllowRemove = False
        .Columns.Add 4, "Applications", 200, True: .Columns(4).Alignment = xtpAlignmentLeft: .Columns(4).AllowRemove = False
        '.Columns.Add 5, "Area", 200, True:              .Columns(5).Alignment = xtpAlignmentLeft:   .Columns(5).AllowRemove = False

        .PaintManager.HorizontalGridStyle = xtpGridSmallDots
        .PaintManager.VerticalGridStyle = xtpGridSmallDots
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = False
        .PaintManager.TextFont.Weight = 500
        .SetCustomDraw xtpCustomBeforeDrawRow
        .AllowColumnRemove = False
    End With
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtSearch.Text = "" Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If rptAccounts.Rows.Count > 0 And rptAccounts.Enabled = True Then: rptAccounts.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCSMS_MasterStockInquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Inquiry"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_MasterStockInquiry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   11295
   Begin XtremeReportControl.ReportControl rptRO 
      Height          =   6405
      Left            =   30
      TabIndex        =   0
      Top             =   630
      Width           =   11235
      _Version        =   655364
      _ExtentX        =   19817
      _ExtentY        =   11298
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnReorder=   0   'False
      MultipleSelection=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.OptionButton Option1 
      Caption         =   "By &Part no"
      Height          =   315
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Value           =   -1  'True
      Width           =   1845
   End
   Begin VB.OptionButton Option2 
      Caption         =   "By &Description"
      Height          =   315
      Left            =   7530
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   150
      Width           =   1845
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   60
      TabIndex        =   2
      Top             =   150
      Width           =   5565
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1860
      TabIndex        =   6
      Top             =   7050
      Width           =   9405
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  F3 - To Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   30
      TabIndex        =   5
      Top             =   7050
      Width           =   1815
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11295
      _Version        =   655364
      _ExtentX        =   19923
      _ExtentY        =   1032
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
Attribute VB_Name = "frmCSMS_MasterStockInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event ShowForm()
Dim xTYPE                                               As String

Public Sub SetType(XXX As String, FORMNAME As String)
    xTYPE = XXX
    Me.Caption = FORMNAME
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtSearch.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    Call InitializeReportControl
    Call txtSearch_Change
    Option1.Value = True
End Sub

Sub FillGrid(XXX As String)
    Dim RSUPLOAD                                        As New ADODB.Recordset
    Dim REC                                             As XtremeReportControl.ReportRecord
    XXX = Replace(XXX, "'", "")
    
    If XXX = "" Then
        Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 ISNULL(STOCKNO,'') AS STOCKNO, " & _
                " ISNULL(STOCKDESC,'') AS STOCKDESC, " & _
                " ISNULL(NEWNO,'') AS NEWNO, " & _
                " CAST(ISNULL(SRP, 0) AS DECIMAL(18,0)) AS SRP, " & _
                " ISNULL(MODELCODE,'') AS MODELCODE, " & _
                " ISNULL(INVCLASS,'') + ISNULL(SUBINVCLAS,'') AS INVCLASS, " & _
                " CASE " & _
                    " WHEN ISNULL(ONHAND,0) > 0 THEN 'Y' " & _
                    " WHEN ISNULL(ONHAND,0) < 1 THEN 'N' " & _
                    " END As ONHAND " & _
                " ,ISNULL(LOCATION,'') AS LOCATION " & _
                " FROM PMIS_STOCKMAS " & _
                " WHERE TYPE = " & N2Str2Null(xTYPE) & _
                " ORDER BY STOCKNO")
    Else
        If Option1.Value = True Then
            XXX = " AND STOCKNO LIKE '%" & XXX & "%'"
            Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 ISNULL(STOCKNO,'') AS STOCKNO, " & _
                " ISNULL(STOCKDESC,'') AS STOCKDESC, " & _
                " ISNULL(NEWNO,'') AS NEWNO, " & _
                " CAST(ISNULL(SRP, 0) AS DECIMAL(18,0)) AS SRP, " & _
                " ISNULL(MODELCODE,'') AS MODELCODE, " & _
                " ISNULL(INVCLASS,'') + ISNULL(SUBINVCLAS,'') AS INVCLASS, " & _
                " CASE " & _
                    " WHEN ISNULL(ONHAND,0) > 0 THEN 'Y' " & _
                    " WHEN ISNULL(ONHAND,0) < 1 THEN 'N' " & _
                    " END As ONHAND " & _
                " ,ISNULL(LOCATION,'') AS LOCATION " & _
                " FROM PMIS_STOCKMAS " & _
                " WHERE TYPE = " & N2Str2Null(xTYPE) & _
                " " & XXX & _
                " ORDER BY STOCKNO")
        Else
            XXX = " AND STOCKDESC LIKE '%" & XXX & "%'"
            Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 ISNULL(STOCKNO,'') AS STOCKNO, " & _
                " ISNULL(STOCKDESC,'') AS STOCKDESC, " & _
                " ISNULL(NEWNO,'') AS NEWNO, " & _
                " CAST(ISNULL(SRP, 0) AS DECIMAL(18,0)) AS SRP, " & _
                " ISNULL(MODELCODE,'') AS MODELCODE, " & _
                " ISNULL(INVCLASS,'') + ISNULL(SUBINVCLAS,'') AS INVCLASS, " & _
                " CASE " & _
                    " WHEN ISNULL(ONHAND,0) > 0 THEN 'Y' " & _
                    " WHEN ISNULL(ONHAND,0) < 1 THEN 'N' " & _
                    " END As ONHAND " & _
                " ,ISNULL(LOCATION,'') AS LOCATION " & _
                " FROM PMIS_STOCKMAS " & _
                " WHERE TYPE = " & N2Str2Null(xTYPE) & _
                " " & XXX & _
                " ORDER BY STOCKNO")
        End If
    End If
    rptRO.Records.DeleteAll
    While Not RSUPLOAD.EOF
        Set REC = rptRO.Records.Add
        REC.AddItem (Trim(RSUPLOAD!STOCKNO))
        REC.AddItem (Trim(RSUPLOAD!STOCKDESC))
        REC.AddItem (Trim(RSUPLOAD!NEWNO))
        REC.AddItem (Trim(Format(RSUPLOAD!SRP, MAXIMUM_DIGIT)))
        REC.AddItem (Trim(RSUPLOAD!MODELCODE))
        REC.AddItem (Trim(RSUPLOAD!INVCLASS))
        REC.AddItem (Trim(RSUPLOAD!ONHAND))
        REC.AddItem (Trim(RSUPLOAD!Location))
        
        RSUPLOAD.MoveNext
        Set REC = Nothing
    Wend
    rptRO.Populate
    
    Set RSUPLOAD = Nothing
End Sub

Sub InitializeReportControl()
    Screen.MousePointer = 11
    
    With rptRO
        .Columns.DeleteAll
        .Columns.Add 0, "PART NUMBER", 90, True::       .Columns(0).Resizable = False:                  .Columns(0).Resizable = True:   .Columns(0).AllowRemove = False
        .Columns.Add 1, "DESCRIPTION", 185, True:       .Columns(1).AllowRemove = False:                .Columns(1).Resizable = True:   .Columns(1).AllowRemove = False
        .Columns.Add 2, "SUPERCESSION", 90, True:      .Columns(2).AllowRemove = False:                .Columns(2).Resizable = True:   .Columns(2).AllowRemove = False
        .Columns.Add 3, "SRP", 60, True:                .Columns(3).Alignment = xtpAlignmentRight:     .Columns(3).Resizable = True:   .Columns(3).AllowRemove = False
        .Columns.Add 4, "MODEL", 90, True:              .Columns(4).Alignment = xtpAlignmentLeft:       .Columns(4).Resizable = True:   .Columns(4).AllowRemove = False
        .Columns.Add 5, "ICC", 60, True:                .Columns(5).Alignment = xtpAlignmentCenter:     .Columns(5).Resizable = True:   .Columns(5).AllowRemove = False
        .Columns.Add 6, "STOCK", 60, True:              .Columns(6).Alignment = xtpAlignmentCenter:     .Columns(6).Resizable = True:   .Columns(6).AllowRemove = False
        .Columns.Add 7, "LOCATION", 90, True:           .Columns(7).Alignment = xtpAlignmentCenter:     .Columns(7).Resizable = True:   .Columns(7).AllowRemove = False
        
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

Private Sub Option1_Click()
    txtSearch.SetFocus
End Sub

Private Sub Option2_Click()
    txtSearch.SetFocus
End Sub

Private Sub txtSearch_Change()
    Call FillGrid(txtSearch)
    'rptRO.FilterText = txtSearch.Text
    'rptRO.Populate
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.BackColor = &HC0FFC0
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch.BackColor = vbWhite
End Sub

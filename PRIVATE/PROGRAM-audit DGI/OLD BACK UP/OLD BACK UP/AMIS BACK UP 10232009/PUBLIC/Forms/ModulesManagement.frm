VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmRAMS_ModulesSheet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DMIS 2.0 Modules  Sheet"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ModulesManagement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl Grid 
      Height          =   4695
      Left            =   90
      TabIndex        =   0
      Top             =   1020
      Width           =   7785
      _Version        =   655364
      _ExtentX        =   13732
      _ExtentY        =   8281
      _StockProps     =   64
      BorderStyle     =   4
      ShowFooter      =   -1  'True
   End
   Begin VB.TextBox txtSearch 
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   510
      Width           =   4425
   End
   Begin XtremeShortcutBar.ShortcutCaption lblDescription 
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   -30
      Width           =   7935
      _Version        =   655364
      _ExtentX        =   13996
      _ExtentY        =   820
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
Attribute VB_Name = "frmRAMS_ModulesSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents AEFormModule        As frmRAMS_AEModule
Attribute AEFormModule.VB_VarHelpID = -1

Private Sub AEFormModule_ChangedRecord(o As Boolean)
    LoadFilGrid
    Set AEFormModule = Nothing
End Sub
Private Sub LoadFilGrid()
    SQL = "SELECT B.DESCRIPTIONS,ISNULL(B.MODULE_TYPE, 'NOT CONFIG') ,B.ID FROM ALL_RAMS_MODULES B INNER JOIN ALL_PROFILE  A ON B.MAINMODULEID=A.ID Where A.MODULENAME=" & N2Str2Null(MODULENAME)
    flex_FillReportView gconDMIS.Execute(SQL), Grid


End Sub
Sub Form_Load()
    Me.Caption = "DMIS 2.0 " & MODULENAME & " Modules Sheet "
    CenterMe frmMain, Me, 1
    With Grid
        .Columns.Add 0, "Module Name", 100, True
        .Columns.Add 1, "Module Type", 100, True
    End With
    LoadFilGrid
    lblDescription.Caption = "SYSTEM MODULE FOR ::" & MODULENAME
End Sub

Private Sub Grid_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record Is Nothing Then: Exit Sub
    If Row.Record(1).Value = "NOT CONFIG" Then
        Metrics.ForeColor = RGB(116, 49, 41)
    End If
End Sub

Private Sub Grid_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal item As XtremeReportControl.IReportRecordItem)
    Set AEFormModule = New frmRAMS_AEModule
    Call AEFormModule.EditModule(item.Record(2).Value)
    AEFormModule.Show
End Sub
Private Function flex_FillReportView(rs As Recordset, grd As ReportControl, Optional ByVal WithSN As Boolean = False)

    Dim fld                            As Field
    Dim j                              As Long
    Dim REC                            As XtremeReportControl.ReportRecord


    grd.Records.DeleteAll


    While Not rs.EOF
        j = j + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem j
        End If
        For Each fld In rs.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        rs.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set rs = Nothing
End Function

Private Sub txtsearch_Change()
    Grid.FilterText = txtSearch.Text
    Grid.Populate
End Sub

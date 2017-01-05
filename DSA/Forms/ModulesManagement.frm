VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmModulesSheet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DMIS 2.0 Modules  Sheet"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   9750
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
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeReportControl.ReportControl Grid 
      Height          =   5805
      Left            =   90
      TabIndex        =   2
      Top             =   1350
      Width           =   9585
      _Version        =   655364
      _ExtentX        =   16907
      _ExtentY        =   10239
      _StockProps     =   64
      BorderStyle     =   4
      ShowFooter      =   -1  'True
   End
   Begin VB.ComboBox cboMainModule 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   3525
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DMIS 2.0Module List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   3
      Top             =   240
      Width           =   1920
   End
   Begin VB.Label lblDescription 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3690
      TabIndex        =   1
      Top             =   960
      Width           =   5925
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   2100
      Picture         =   "ModulesManagement.frx":000C
      Top             =   0
      Width           =   7665
   End
   Begin VB.Image Image2 
      Height          =   885
      Left            =   0
      Picture         =   "ModulesManagement.frx":0B1E
      Top             =   0
      Width           =   7665
   End
End
Attribute VB_Name = "frmModulesSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents AEFormModule                        As frmAEModule
Attribute AEFormModule.VB_VarHelpID = -1
Private Sub AEFormModule_ChangedRecord(o As Boolean)
    cboMainModule_CLICK
    Set AEFormModule = Nothing
End Sub
Private Sub cboMainModule_CLICK()
    SQL = "SELECT B.DESCRIPTIONS,ISNULL(B.MODULE_TYPE, 'NOT CONFIG') ,B.MODULEID FROM ALL_RAMS_MODULES B INNER JOIN ALL_PROFILE  A ON B.MAINMODULEID=A.ID Where A.MODULENAME=" & N2Str2Null(cboMainModule.Text)
    flex_FillReportView gconDMIS.Execute(SQL), Grid
    Dim TEMPRS                                         As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("Select ModuleDescription from ALL_PROFILE WHERE ModuleName=" & N2Str2Null(cboMainModule.Text))

    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        lblDescription = TEMPRS!ModuleDescription
    End If

End Sub
Sub Form_Load()

    Call FillCombo("Select ID, ModuleName from ALL_PROFILE", 0, 1, cboMainModule)
    ReportControlPaintManager Grid
    ReportControlAddColumnHeader Grid, "Module Name, Module Type"
    ConfigHeaders Grid, "100,100"
    cboMainModule.ListIndex = 0
End Sub

Private Sub Grid_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record Is Nothing Then: Exit Sub


    If Row.Record(1).Value = "NOT CONFIG" Then
        '        Metrics.ForeColor = RGB(44, 72, 92)
        '    ElseIf Row.Record(1).Value = "SEARCH" Then
        '        Metrics.ForeColor = vbBlue
        '    ElseIf Row.Record(1).Value = "INQUIRY" Then
        '        Metrics.ForeColor = vbYellow
        '    ElseIf Row.Record(1).Value = "INQUIRY" Then
        '        Metrics.ForeColor = RGB(188, 199, 50)
        '    ElseIf Row.Record(1).Value = "REPORTS" Then
        '        Metrics.ForeColor = RGB(190, 245, 200)
        '    ElseIf Row.Record(1).Value = "ALL" Then
        '        Metrics.ForeColor = RGB(48, 101, 56)
        Metrics.ForeColor = RGB(116, 49, 41)
    Else


    End If
End Sub

Private Sub Grid_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal item As XtremeReportControl.IReportRecordItem)
    Set AEFormModule = New frmAEModule
    Call AEFormModule.EditModule(item.Record(2).Value)
    AEFormModule.Show 1
End Sub

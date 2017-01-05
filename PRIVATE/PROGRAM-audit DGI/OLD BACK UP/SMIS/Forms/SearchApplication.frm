VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMISSearchApplication 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Application"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl lvGrid 
      Height          =   4425
      Left            =   60
      TabIndex        =   4
      Top             =   780
      Width           =   6195
      _Version        =   655364
      _ExtentX        =   10927
      _ExtentY        =   7805
      _StockProps     =   64
      BorderStyle     =   2
   End
   Begin VB.OptionButton optselect 
      Caption         =   "&Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1815
      TabIndex        =   5
      Tag             =   "TelePhone"
      Top             =   60
      Width           =   1245
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   690
      Left            =   4320
      Picture         =   "SearchApplication.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5340
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   690
      Left            =   5340
      Picture         =   "SearchApplication.frx":033C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5340
      Width           =   945
   End
   Begin VB.OptionButton optselect 
      Caption         =   "&Applicant Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Tag             =   "AcctName"
      Top             =   60
      Width           =   1680
   End
   Begin VB.TextBox txtsearch 
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
      Left            =   15
      TabIndex        =   0
      Top             =   390
      Width           =   3150
   End
End
Attribute VB_Name = "frmSMISSearchApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event SelectionMade(oCusRs As ADODB.Recordset)

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdSelect_Click()
    If Not lvGrid.SelectedRows.Count <= 0 Then
        Call lvGrid_RowDblClick(lvGrid.Rows(lvGrid.SelectedRows.Row(0).Index), Nothing)
    Else
        MessagePop InfoVoid, "Selection Required", "There is Nothing To Select from ", 1000, 1
    End If

End Sub


Private Sub Form_Load()
    ' LogInitialInquiry  AcctName,   ContactPerson , Variant,    Address,    Email   TelePhone,  SAE,    CUSCDE,     ProspectType, ProspectID
    ReportControlAddColumnHeader lvGrid, "Date,ProspectName, ContactPerson, Model, SAE"

    ReportControlPaintManager lvGrid
    With lvGrid
        .Columns(0).FooterText = "F3: Add Filter"
        .Columns(1).FooterText = "F8: Remove Filter"
    End With
    ResizeColumnHeader lvGrid, "40, 20,20,20"

    flex_FillReportView gconDMIS.Execute("SELECT TOP 100  DateApplied, Ind_apl_lastname + Ind_apl_lastname + Ind_apl_lastname , Ind_address, Ind_LoanApl_unitModel , ID  From SMIS_LOANINDIV ORDER BY DateApplied DESC "), lvGrid, False

End Sub


Private Sub Form_Unload(Cancel As Integer)
    strfor = vbNullString
    CHKCUSCDE = vbNullString
End Sub

Private Sub lvGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call lvGrid_RowDblClick(lvGrid.Rows(lvGrid.SelectedRows.Row(0).Index), Nothing)
    End If
    If KeyCode = vbKeyF3 Then
        Call frmSMISFilter.ConfigGrid(lvGrid, 0)
        frmSMISFilter.Show 1
    End If
End Sub

Private Sub lvGrid_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then: Exit Sub
    Dim Temprs                                    As ADODB.Recordset
    Set Temprs = gconDMIS.Execute("SELECT * FROM SMIS_LOANINDIV WHERE ID=" & Row.Record(4).Value)
    If Not Temprs Is Nothing Then
        RaiseEvent SelectionMade(Temprs)
    End If

End Sub

'   0             1                                                         2                3               4
'DateApplied, Ind_apl_lastname + Ind_apl_lastname + Ind_apl_lastname , Ind_address, Ind_LoanApl_unitModel , ID

Private Sub optselect_Click(Index As Integer)
    txtsearch_Change
End Sub

Private Sub txtsearch_Change()
    Dim Temprs                                    As ADODB.Recordset
    Dim i                                         As Integer
    Dim KEY                                       As String

    For i = 0 To optselect.Count - 1
        If optselect(i).Value = True Then
            KEY = optselect(i).Tag
        End If
    Next

    Set Temprs = gconDMIS.Execute("select TOP 50 LogInitialInquiry,  AcctName, ContactPerson , Variant, Address, Email, TelePhone, SAE, CUSCDE, ProspectType, ProspectID from CRIS_PROSPECTS where D_S is Null AND " & KEY & " like '%" & ReplaceQuote(txtSearch.Text) & "%'")
    flex_FillReportView Temprs, lvGrid, False

    lvGrid.Populate
End Sub




Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If lvGrid.Records.Count <= 0 Then: Exit Sub
    If KeyCode = vbKeyDown Then
        lvGrid.SetFocus
    End If
End Sub

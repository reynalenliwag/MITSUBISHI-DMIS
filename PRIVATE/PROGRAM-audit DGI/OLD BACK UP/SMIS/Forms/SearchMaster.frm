VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMISSearhMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Customer"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
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
   ScaleHeight     =   6300
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl lvGrid 
      Height          =   4725
      Left            =   120
      TabIndex        =   7
      Top             =   780
      Width           =   8895
      _Version        =   655364
      _ExtentX        =   15690
      _ExtentY        =   8334
      _StockProps     =   64
      BorderStyle     =   2
      ShowFooter      =   -1  'True
   End
   Begin VB.OptionButton optselect 
      Caption         =   "A&ddress"
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
      Index           =   4
      Left            =   2220
      TabIndex        =   9
      Tag             =   "ADDRESS"
      Top             =   120
      Width           =   1035
   End
   Begin VB.OptionButton optselect 
      Caption         =   "&Telephone"
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
      Left            =   3300
      TabIndex        =   8
      Tag             =   "PHONE"
      Top             =   120
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add New"
      Height          =   690
      Left            =   6060
      Picture         =   "SearchMaster.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5565
      Width           =   945
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   690
      Left            =   7065
      Picture         =   "SearchMaster.frx":0313
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5565
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   690
      Left            =   8085
      Picture         =   "SearchMaster.frx":064F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5565
      Width           =   945
   End
   Begin VB.OptionButton optselect 
      Caption         =   "&Email"
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
      Index           =   2
      Left            =   2220
      TabIndex        =   6
      Tag             =   "EMAIL"
      Top             =   420
      Width           =   825
   End
   Begin VB.OptionButton optselect 
      Caption         =   "Account &Name"
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
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Tag             =   "ACCTNAME"
      Top             =   120
      Value           =   -1  'True
      Width           =   1515
   End
   Begin VB.OptionButton optselect 
      Caption         =   "&Customer Name"
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
      Left            =   120
      TabIndex        =   4
      Tag             =   "CUSTOMERNAME"
      Top             =   420
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
      Left            =   3120
      TabIndex        =   0
      Top             =   420
      Width           =   5865
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Account Type:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4410
      TabIndex        =   15
      Top             =   5535
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contact Person:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2265
      TabIndex        =   14
      Top             =   5520
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   765
   End
   Begin VB.Label lblMis 
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
      Height          =   495
      Left            =   4440
      TabIndex        =   12
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lblContactPerson 
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
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   5760
      Width           =   2115
   End
   Begin VB.Label lblAddress 
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
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5760
      Width           =   2115
   End
End
Attribute VB_Name = "frmSMISSearhMaster"
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
Private Sub Command2_Click()
    frmALLCustomer.cmdAdd.Value = True
    frmALLCustomer.Show
End Sub

Private Sub Form_Load()
    ReportControlAddColumnHeader lvGrid, "CustomerName, Email, Mobile, Phone"
    ReportControlPaintManager lvGrid
    With lvGrid
        .Columns(0).FooterText = "F3: Add Filter"
        .Columns(1).FooterText = "F8: Remove Filter"
    End With
    ResizeColumnHeader lvGrid, "40, 20,20,20"

    'CUSTID,CUSCDE, CUSTYPE, , ,  ,CONTACTPESON,  ,   , ,
    flex_FillReportView gconDMIS.Execute("SELECT TOP 100  CUSTOMERNAME, EMAIL, MOBILE, PHONE,  CUSTID , CUSCDE , CUSTYPE , ACCTNAME, ADDRESS, CONTACTPERSON from CRIS_vW_AllProfile"), lvGrid, False

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
    Set Temprs = gconDMIS.Execute("SELECT * FROM ALL_CUSTOMER WHERE ID=" & Row.Record(4).Value)
    If Not Temprs Is Nothing Then
        RaiseEvent SelectionMade(Temprs)
    End If

End Sub

Private Sub lvGrid_SelectionChanged()
    Dim MIS_INFO                                  As String
    '                   0           1       2       3        4         5       6        7           8        9
    '               CUSTOMERNAME,   EMAIL, MOBILE, PHONE,  CUSTID , CUSCDE , CUSTYPE , ACCTNAME, ADDRESS, CONTACTPERSON
    lblAddress = lvGrid.SelectedRows.Row(0).Record(8).Value
    If lvGrid.SelectedRows.Row(0).Record(6).Value = "C" Then
        MIS_INFO = "Company/Agency"
        lblContactPerson = lvGrid.SelectedRows.Row(0).Record(9).Value
    ElseIf lvGrid.SelectedRows.Row(0).Record(6).Value = "P" Then
        MIS_INFO = "Personal"
        lblContactPerson = lvGrid.SelectedRows.Row(0).Record(0).Value
    ElseIf lvGrid.SelectedRows.Row(0).Record(6).Value = "F" Then
        MIS_INFO = "Fleet"
        lblContactPerson = lvGrid.SelectedRows.Row(0).Record(9).Value
    ElseIf lvGrid.SelectedRows.Row(0).Record(6).Value = "G" Then
        MIS_INFO = "Government"
        lblContactPerson = lvGrid.SelectedRows.Row(0).Record(9).Value
    End If

    lblMis = MIS_INFO



End Sub

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
    Set Temprs = gconDMIS.Execute("select TOP 50 CUSTOMERNAME, EMAIL, MOBILE, PHONE,  CUSTID , CUSCDE , CUSTYPE , ACCTNAME, ADDRESS, CONTACTPERSON  from CRIS_vW_AllProfile where " & KEY & " like '%" & ReplaceQuote(txtSearch.Text) & "%'")
    flex_FillReportView Temprs, lvGrid, False

    lvGrid.Populate
End Sub




Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If lvGrid.Records.Count <= 0 Then: Exit Sub
    If KeyCode = vbKeyDown Then
        lvGrid.SetFocus
    End If
End Sub

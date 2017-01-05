VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMISSearchProspects 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Prospects"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
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
   ScaleHeight     =   6135
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl lvGrid 
      Height          =   4425
      Left            =   60
      TabIndex        =   5
      Top             =   780
      Width           =   6195
      _Version        =   655364
      _ExtentX        =   10927
      _ExtentY        =   7805
      _StockProps     =   64
      BorderStyle     =   2
   End
   Begin VB.OptionButton optselect 
      Caption         =   "SAE"
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
      Index           =   5
      Left            =   3255
      TabIndex        =   9
      Tag             =   "SAE"
      Top             =   420
      Width           =   825
   End
   Begin VB.OptionButton optselect 
      Caption         =   "&Vehicles"
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
      Left            =   3255
      TabIndex        =   8
      Tag             =   "Variant"
      Top             =   60
      Width           =   1245
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
      Left            =   4500
      TabIndex        =   7
      Tag             =   "ADDRESS"
      Top             =   60
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
      Left            =   1815
      TabIndex        =   6
      Tag             =   "TelePhone"
      Top             =   60
      Width           =   1245
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   690
      Left            =   4320
      Picture         =   "SearchProspects.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5310
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   690
      Left            =   5340
      Picture         =   "SearchProspects.frx":033C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5310
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
      Left            =   4500
      TabIndex        =   4
      Tag             =   "EMAIL"
      Top             =   360
      Width           =   825
   End
   Begin VB.OptionButton optselect 
      Caption         =   "&Prospect Name"
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
Attribute VB_Name = "frmSMISSearchProspects"
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

    'CUSTID,CUSCDE, CUSTYPE, , ,  ,CONTACTPESON
    flex_FillReportView gconDMIS.Execute("SELECT TOP 100  LogInitialInquiry,  AcctName, ContactPerson , Variant, Address, Email, TelePhone, SAE, CUSCDE, ProspectType, ProspectID from CRIS_PROSPECTS WHERE D_S IS NULL ORDER BY LOGINITIALINQUIRY DESC "), lvGrid, False

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
    
    Set Temprs = gconDMIS.Execute("SELECT * FROM CRIS_PROSPECTS WHERE PROSPECTID=" & Row.Record(10).Value)
    If Not Temprs Is Nothing Then
        RaiseEvent SelectionMade(Temprs)
    End If

End Sub

Private Sub lvGrid_SelectionChanged()

    '               0             1       2                3               4         5        6        7        8            9               10
    'LogInitialInquiry AcctName ContactPerson    Variant,    Address,    Email   TelePhone,  SAE,    CUSCDE,     ProspectType, ProspectID from CRIS_PROSPECTS WHERE D_S IS NULL ORDER BY LOGINITIALINQUIRY DESC "), lvGrid, False

    Dim MIS_INFO                                  As String
    If lvGrid.SelectedRows.Row(0).Record(8).Value = "C" Then
        MIS_INFO = "Company/Agency"
    ElseIf lvGrid.SelectedRows.Row(0).Record(8).Value = "P" Then
        MIS_INFO = "Personal"
    ElseIf lvGrid.SelectedRows.Row(0).Record(8).Value = "F" Then
        MIS_INFO = "Fleet"
    ElseIf lvGrid.SelectedRows.Row(0).Record(8).Value = "G" Then
        MIS_INFO = "Government"
    End If

    lblContactDetails = MIS_INFO

    If Null2String(lvGrid.SelectedRows.Row(0).Record(7).Value) = vbNullString Then
        CustomerInfoCard1.ConvertToCustomer
    Else
        CustomerInfoCard1.CustomerCode = lvGrid.SelectedRows.Row(0).Record(7).Value
    End If

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

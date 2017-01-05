VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmAMISCustomerVendor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Customer Information"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10035
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmAMISCustomerVendor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   $"frmAMISCustomerVendor.frx":1082
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   5055
      Left            =   30
      TabIndex        =   4
      Top             =   690
      Width           =   9975
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4755
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   9855
         _Version        =   655364
         _ExtentX        =   17383
         _ExtentY        =   8387
         _StockProps     =   64
         BorderStyle     =   4
      End
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   210
      Width           =   8715
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
      Left            =   120
      TabIndex        =   0
      Top             =   270
      Width           =   945
   End
   Begin VB.Label Label2 
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
      Left            =   140
      TabIndex        =   3
      Top             =   290
      Width           =   945
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10035
      _Version        =   655364
      _ExtentX        =   17701
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
Attribute VB_Name = "frmAMISCustomerVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event SelectedInfo(strSelected As String)

Sub InitReportControl()
    With rptList
        .Columns.DeleteAll
        If xSELECTED = "Customer" Then
            Me.Caption = "Search Customer Information"
            .Columns.Add 0, "Code", 80, True: .Columns(0).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
            .Columns.Add 1, "Lastname", 180, True: .Columns(1).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
            .Columns.Add 2, "First Name", 180, True: .Columns(2).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
            .Columns.Add 3, "Account Name", 220, True: .Columns(3).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
            .Columns.Add 4, "Mobile", 90, True: .Columns(4).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
            .Columns.Add 5, "Home Phone", 90, True: .Columns(5).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
        ElseIf xSELECTED = "Vendor" Then
            Me.Caption = "Search Vendor Information"
            .Columns.Add 0, "Code", 80, True: .Columns(0).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
            .Columns.Add 1, "Name of Vendor", 180, True: .Columns(1).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
            .Columns.Add 2, "Contact Person", 180, True: .Columns(2).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
            .Columns.Add 3, "Address", 220, True: .Columns(3).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
            .Columns.Add 4, "Phone No.", 90, True: .Columns(4).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
            .Columns.Add 5, "Fax No.", 90, True: .Columns(5).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False
        End If
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots
        .PaintManager.VerticalGridStyle = xtpGridSmallDots
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.TextFont.Weight = 500
        .SetCustomDraw xtpCustomBeforeDrawRow
    End With
End Sub

Sub FillGrid()
    Screen.MousePointer = 11
    Dim REC                                       As XtremeReportControl.ReportRecord
    If xSELECTED = "Customer" Then
        Dim rsCustomer                            As ADODB.Recordset
        Set rsCustomer = New ADODB.Recordset
        rsCustomer.Open "Select CUSTCODE, LASTNAME, FIRSTNAME, ACCTNAME, MOBILE, TELEPHONENO from All_CustMaster_AMIS order by AcctName", gconDMIS, adOpenKeyset
        rptList.Records.DeleteAll
        While Not rsCustomer.EOF
            Set REC = rptList.Records.Add
            With REC
                .AddItem Null2String(rsCustomer!CUSTCODE)
                .AddItem Null2String(rsCustomer!lastname)
                .AddItem Null2String(rsCustomer!Firstname)
                .AddItem Null2String(rsCustomer!AcctName)
                .AddItem Null2String(rsCustomer!Mobile)
                .AddItem Null2String(rsCustomer!TelephoneNo)
            End With
            DoEvents
            rsCustomer.MoveNext
        Wend
        rptList.Populate
        Set rsCustomer = Nothing
    ElseIf xSELECTED = "Vendor" Then
        Dim rsVENDOR                              As ADODB.Recordset
        Set rsVENDOR = New ADODB.Recordset
        rsVENDOR.Open "Select CODE,NAMEOFVENDOR,CONTACT,ADDRESS,PHONE,FAX from All_Vendor order by NameofVendor", gconDMIS, adOpenKeyset
        rptList.Records.DeleteAll
        While Not rsVENDOR.EOF
            Set REC = rptList.Records.Add
            With REC
                .AddItem Null2String(rsVENDOR!code)
                .AddItem Null2String(rsVENDOR!nameofvendor)
                .AddItem Null2String(rsVENDOR!CONTACT)
                .AddItem Null2String(rsVENDOR!Address)
                .AddItem Null2String(rsVENDOR!Phone)
                .AddItem Null2String(rsVENDOR!Fax)
            End With
            DoEvents
            rsVENDOR.MoveNext
        Wend
        rptList.Populate
        Set rsVENDOR = Nothing
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    InitReportControl
    FillGrid
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    RaiseEvent SelectedInfo(Row.Record(0).Value)
    Unload Me
End Sub

Private Sub txtSEARCH_Change()
    rptList.FilterText = txtSearch.Text
    rptList.Populate
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtSearch.Text = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If

    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If rptList.Rows.Count > 0 And rptList.Enabled = True Then: rptList.SetFocus
    End If

    If KeyCode = vbKeyEscape Then Unload Me
End Sub

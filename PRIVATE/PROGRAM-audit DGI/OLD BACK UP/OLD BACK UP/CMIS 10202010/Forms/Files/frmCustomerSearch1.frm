VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCustomerSearch1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Customer"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ForeColor       =   &H8000000F&
   Icon            =   "frmCustomerSearch1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl rptCustomer 
      Height          =   4305
      Left            =   60
      TabIndex        =   1
      Top             =   870
      Width           =   7215
      _Version        =   655364
      _ExtentX        =   12726
      _ExtentY        =   7594
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnReorder=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   7230
   End
   Begin VB.TextBox txtActiveForm 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Text            =   "txtActiveForm"
      Top             =   -510
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5250
      Width           =   5625
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5250
      Width           =   1545
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   6420
      MouseIcon       =   "frmCustomerSearch1.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmCustomerSearch1.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Add Customer"
      Top             =   5670
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   5580
      MouseIcon       =   "frmCustomerSearch1.frx":076F
      MousePointer    =   99  'Custom
      Picture         =   "frmCustomerSearch1.frx":08C1
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   5670
      Width           =   855
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4740
      MouseIcon       =   "frmCustomerSearch1.frx":0BFF
      MousePointer    =   99  'Custom
      Picture         =   "frmCustomerSearch1.frx":0D51
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Select this Customer"
      Top             =   5670
      Width           =   855
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   8
      Top             =   15
      Width           =   7230
      _Version        =   655364
      _ExtentX        =   12753
      _ExtentY        =   503
      _StockProps     =   14
      Caption         =   "SEARCH (Type your keyword here)"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColorLight=   16744576
      GradientColorDark=   12582912
      ForeColor       =   16777215
   End
End
Attribute VB_Name = "frmCustomerSearch1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    Unload Me
    frmAllCustomer.cmdAdd.Value = True
    frmAllCustomer.Show 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    If SelectCustomer = "Bank" Then
        If Trim(txtCode.Text) <> "" Then
            frmCreditCardCompany.txtCuscde.Text = txtCode.Text
            frmCreditCardCompany.cboCUSNAME.Text = txtName.Text
            Unload Me
        End If
    ElseIf SelectCustomer = "Customer" Then
        If Trim(txtCode.Text) <> "" Then
            frmCMISOREntry.txtCuscde.Text = txtCode.Text
            'frmCMISOREntry.SetCustomer
            frmCMISOREntry.cboCUSNAME.Text = txtName.Text
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 11
    DisplayCustomer
End Sub

Sub DisplayCustomer()
    Screen.MousePointer = 11
    Dim rsAllCustomer           As ADODB.Recordset
    Dim REC                     As XtremeReportControl.ReportRecord
'    Call ReportControlAddColumnHeader(rptCustomer, " CODE, Lastname, First Name, Account Name, Mobile, Home Phone, Fax, Address, City, Province")
'    Call ReportControlPaintManager(rptCustomer)
'    'rptCustomer.GroupsOrder.Add rptCustomer.Columns(0)
'    Call ResizeColumnHeader(rptCustomer, "20, 50, 50, 50, 25, 30, 25, 25, 15, 15, 15")
'    Call flex_FillReportView(gconDMIS.Execute("select CusCde,LastName,FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where cuscde <> '999999' order by lastname asc"), rptCustomer)
    With rptCustomer
        .Columns.DeleteAll
        .Columns.Add 0, "Customer Code", 45, True:   .Columns(0).Alignment = xtpAlignmentLeft:   .Columns(0).AllowRemove = False
        .Columns.Add 1, "Lastname", 150, True:       .Columns(1).Alignment = xtpAlignmentLeft:   .Columns(0).AllowRemove = False
        .Columns.Add 2, "First name", 150, True:     .Columns(2).Alignment = xtpAlignmentLeft:   .Columns(0).AllowRemove = False
        .Columns.Add 3, "Account Name", 200, True:   .Columns(3).Alignment = xtpAlignmentLeft:   .Columns(0).AllowRemove = False
        .Columns.Add 4, "Mobile", 90, True:          .Columns(4).Alignment = xtpAlignmentLeft:   .Columns(0).AllowRemove = False
        .Columns.Add 5, "Fax", 90, True:             .Columns(5).Alignment = xtpAlignmentLeft:   .Columns(0).AllowRemove = False
        .Columns.Add 6, "Home Phone", 90, True:      .Columns(6).Alignment = xtpAlignmentLeft:   .Columns(0).AllowRemove = False
        .Columns.Add 7, "Address", 200, True:        .Columns(7).Alignment = xtpAlignmentLeft:   .Columns(0).AllowRemove = False
        .Columns.Add 8, "City", 200, True:           .Columns(8).Alignment = xtpAlignmentLeft:   .Columns(0).AllowRemove = False
        .Columns.Add 9, "Province", 200, True:       .Columns(9).Alignment = xtpAlignmentLeft:   .Columns(0).AllowRemove = False
        
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots      ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = False
        .PaintManager.TextFont.Weight = 500
        .AllowColumnRemove = False
    End With
    
    Set rsAllCustomer = New ADODB.Recordset
    rsAllCustomer.Open "select CusCde,LastName,FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where cuscde <> '999999' order by lastname asc", gconDMIS, adOpenKeyset
    rptCustomer.Records.DeleteAll
    While Not rsAllCustomer.EOF
            Set REC = rptCustomer.Records.Add
            With REC
                .AddItem Null2String(rsAllCustomer!CUSCDE)
                .AddItem Null2String(rsAllCustomer!lastname)
                .AddItem Null2String(rsAllCustomer!Firstname)
                .AddItem Null2String(rsAllCustomer!AcctName)
                .AddItem Null2String(rsAllCustomer!Mobile)
                .AddItem Null2String(rsAllCustomer!HomePhone)
                .AddItem Null2String(rsAllCustomer!Fax)
                .AddItem Null2String(rsAllCustomer!CUSTOMERADD)
                .AddItem Null2String(rsAllCustomer!City)
                .AddItem Null2String(rsAllCustomer!provincialadd)
                DoEvents
            End With
        rsAllCustomer.MoveNext
    Wend
    rptCustomer.Populate
    Screen.MousePointer = 0
End Sub

Sub ReportControlAddColumnHeader(lst As ReportControl, StringHeaders As String)
    Dim ar()                                           As String
    Dim I                                              As Integer

    ar = Split(StringHeaders, ",")
    lst.Columns.DeleteAll
    For I = LBound(ar) To UBound(ar)
        lst.Columns.Add I, ar(I), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Sub ReportControlPaintManager(lst As ReportControl)
    With lst
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnExplorer
    End With
End Sub

Public Function flex_FillReportView(RS As ADODB.Recordset, grd As XtremeReportControl.ReportControl, Optional ByVal WithSN As Boolean = False)
    Dim fld                                            As ADODB.Field
    Dim j                                              As Long
    Dim REC                                            As XtremeReportControl.ReportRecord

    grd.Records.DeleteAll

    While Not RS.EOF
        j = j + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem j
        End If
        For Each fld In RS.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        RS.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RS = Nothing
End Function

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                                           As String
    Dim cWidth                                         As Long
    Dim I                                              As Integer
    Dim scwidth                                        As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For I = LBound(ar) To UBound(ar)
            If I <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.ColumnHeaders(I + 1).Width = scwidth
            End If
        Next
    ElseIf TypeOf grd Is ReportControl Then
        For I = LBound(ar) To UBound(ar)
            If I < grd.Columns.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.Columns(I).Width = scwidth
            End If
        Next

    End If

    Erase ar
    grd.Visible = True
End Sub

Private Sub rptCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtCode.Text = "" Then
        MessagePop InfoFriend, "Customer Info", "Select customer from the list"
        rptCustomer.SetFocus
        Exit Sub
    Else
        If KeyCode = vbKeyReturn Then
            cmdSelect_Click
        End If
    End If
End Sub

Private Sub rptCustomer_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    txtCode = Null2String(rptCustomer.SelectedRows(0).Record(0).Value)
    txtName = Null2String(rptCustomer.SelectedRows(0).Record(3).Value)
End Sub

Private Sub rptCustomer_SelectionChanged()
    txtCode = Null2String(rptCustomer.SelectedRows(0).Record(0).Value)
    txtName = Null2String(rptCustomer.SelectedRows(0).Record(3).Value)
End Sub

Private Sub txtSEARCH_Change()
    rptCustomer.FilterText = txtSearch.Text
    rptCustomer.Populate
End Sub

Private Sub txtSEARCH_GotFocus()
    txtSearch.BackColor = &HC0FFFF
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtSearch.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If rptCustomer.Rows.Count > 0 And rptCustomer.Enabled = True Then: rptCustomer.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtSEARCH_LostFocus()
    txtSearch.BackColor = &HFFFFFF
End Sub

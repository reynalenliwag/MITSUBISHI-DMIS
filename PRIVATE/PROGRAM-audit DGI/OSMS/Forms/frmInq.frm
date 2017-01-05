VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOSMSInquiryReceiving 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplies Received"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11910
   Begin VB.CommandButton mDEExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10740
      TabIndex        =   9
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame Trans_No 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   90
      TabIndex        =   1
      Top             =   300
      Width           =   11745
      Begin VB.OptionButton optNum 
         Caption         =   "MRR &Number"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optDate 
         Caption         =   "MRR &Date"
         Height          =   345
         Left            =   3240
         TabIndex        =   3
         Top             =   120
         Width           =   1845
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   11565
      End
      Begin MSComctlLib.ListView lstReceiving 
         Height          =   4755
         Left            =   60
         TabIndex        =   5
         Top             =   900
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8387
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmInq.frx":0000
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MRR NUMBER"
            Object.Width           =   2559
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "MRR DATE"
            Object.Width           =   2295
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "INVOICE NO."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "PO DATE"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "PO NUMBER"
            Object.Width           =   2416
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "SUPPLIER CODE"
            Object.Width           =   2998
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "RECEIVED BY"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "TOTAL AMOUNT"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Search by:"
         Height          =   345
         Left            =   90
         TabIndex        =   6
         Top             =   150
         Width           =   1065
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdRRDetails 
      Height          =   1755
      Left            =   180
      TabIndex        =   0
      Top             =   4200
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   3096
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
      FormatString    =   "Item #              | Supply                              | Qty                  |  Unit                   "
   End
   Begin VB.CommandButton mDEDetails 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9660
      TabIndex        =   8
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton mDESearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8580
      TabIndex        =   7
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inquiry Supplies Received"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   30
      Width           =   2835
   End
End
Attribute VB_Name = "frmOSMSInquiryReceiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsrrHEADER As ADODB.Recordset
Dim rsrrDETAILS As ADODB.Recordset

Sub rsRefresh()
    Set rsrrHEADER = New ADODB.Recordset
    rsrrHEADER.Open "select * from OSMS_RRHEADER  order by rrnumber asc", gconDMIS
End Sub

Sub StoreMemVars()
    If Not rsrrHEADER.EOF And Not rsrrHEADER.BOF Then
        FillGrid1
    Else
        MsgBoxXP "Record is Empty!", "No Record", XP_OKOnly, msg_Information
    End If
End Sub

Function RecordFound(AAA As Variant) As Boolean
    Dim rsRecordFound As ADODB.Recordset
    Set rsRecordFound = New ADODB.Recordset
    Set rsRecordFound = rsrrHEADER.Clone
    rsRecordFound.Find "RRNumber = '" & AAA & "'"
    If Not rsRecordFound.EOF Then
        rsrrHEADER.Bookmark = rsRecordFound.Bookmark
        RecordFound = True
    Else
        Set rsRecordFound = New ADODB.Recordset
        Set rsRecordFound = rsrrHEADER.Clone
        rsRecordFound.Find "RRDate = '" & CDate(AAA) & "'"
        If Not rsRecordFound.EOF Then
            rsrrHEADER.Bookmark = rsRecordFound.Bookmark
            RecordFound = True
        Else
            RecordFound = False
        End If
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        FillDetailsGrid
    Case vbKeyEscape
        grdRRDetails.ZOrder 1
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    txtSearch.Text = ""
    FillGrid1
End Sub

'Sub FillGrid()
'Set rsrrHEADER = New ADODB.Recordset
'    rsrrHEADER.Open "select * from OSMS_RRHEADER  order by rrnumber asc", gconDMIS
'If Not rsrrHEADER.EOF And Not rsrrHEADER.BOF Then
'      rsrrHEADER.MoveFirst
'      cleargrid grdRRHeader
'      Do While Not rsrrHEADER.EOF
'         grdRRHeader.AddItem Null2Date(rsrrHEADER!rrdate) & Chr(9) & _
          '                           Null2String(rsrrHEADER!rrnumber) & Chr(9) & _
          '                           Null2String(rsrrHEADER!inv_no) & Chr(9) & _
          '                           Null2Date(rsrrHEADER!PO_Date) & Chr(9) & _
          '                           Null2String(rsrrHEADER!PO_No) & Chr(9) & _
          '                           Null2String(rsrrHEADER!Supplier_code) & Chr(9) & _
          '                           Null2String(rsrrHEADER!Receivedby_Code) & Chr(9) & _
          '                           Null2String(rsrrHEADER!Total_Amount)
'         rsrrHEADER.MoveNext
'      Loop
'      If grdRRHeader.Rows > 2 Then grdRRHeader.RemoveItem 1
'Else
'      cleargrid grdRRHeader
'End If
'End Sub

Sub FillDetailsGrid()
'grdRRHeader.Col = 1
    grdRRDetails.ZOrder 0
    Set rsrrDETAILS = New ADODB.Recordset
    'rsrrDETAILS.Open "select * from OSMS_RRDETAILS  where RRNumber = '" & grdRRHeader.Text & "' order by item_no asc", gconDMIS
    rsrrDETAILS.Open "select * from OSMS_RRDETAILS  where RRNumber = '" & lstReceiving.SelectedItem.SubItems(8) & "' order by item_no asc", gconDMIS
    If Not rsrrDETAILS.EOF And Not rsrrDETAILS.BOF Then
        rsrrDETAILS.MoveFirst
        cleargrid grdRRDetails
        Do While Not rsrrDETAILS.EOF
            grdRRDetails.AddItem Null2String(rsrrDETAILS!item_no) & Chr(9) & _
                                 Null2String(rsrrDETAILS!Supply_Code) & Chr(9) & _
                                 Null2String(rsrrDETAILS!rrQUANTITY) & Chr(9) & _
                                 Null2String(rsrrDETAILS!rrunit)
            rsrrDETAILS.MoveNext
        Loop
        If grdRRDetails.Rows > 2 Then grdRRDetails.RemoveItem 1
    Else
        cleargrid grdRRDetails
    End If
End Sub

'Private Sub mDEDelete_Click()
'    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
'        gconDMIS.Execute "delete from OSMS_RRHEADER  where RRNumber = '" & grdRRHeader.Text & "'"
'        gconDMIS.Execute "delete from OSMS_RRDETAILS  where RRNumber = '" & grdRRHeader.Text & "'"
'        'rsRefresh
'        StoreMemvars
'    End If
'End Sub

Private Sub mDEDetails_Click()
    FillDetailsGrid
End Sub

Private Sub mDEExit_Click()
    Unload Me
End Sub

Private Sub mDESearch_Click()
On Error Resume Next
    txtSearch.SetFocus
    
End Sub


Private Sub lstReceiving_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstReceiving
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub txtSearch_Change()
    If optNum.Value = True Then
        If Trim(txtSearch.Text) = "" Then FillGrid1 Else FillSearchGrid1 (txtSearch.Text)
    Else
        If Trim(txtSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (txtSearch.Text)
    End If
End Sub

Sub FillGrid2()
    Dim rsrrHEADER2 As ADODB.Recordset
    lstReceiving.Sorted = False: lstReceiving.ListItems.Clear
    lstReceiving.Enabled = False
    Set rsrrHEADER2 = New ADODB.Recordset
    Set rsrrHEADER2 = gconDMIS.Execute("select RRDate,RRNumber,INV_No,INV_Date,PO_No,PO_Date,Supplier_Code,ReceivedBy_Code,Total_Amount,RRNumber from OSMS_RRHEADER  order by RRDate asc")
    If Not (rsrrHEADER2.EOF And rsrrHEADER2.BOF) Then
        Listview_Loadval Me.lstReceiving.ListItems, rsrrHEADER2
        lstReceiving.Refresh
        lstReceiving.Enabled = True
    End If
    
End Sub

Sub FillSearchGrid2(xxx As String)
    Dim rsrrHEADER2 As ADODB.Recordset
    lstReceiving.Enabled = False
    lstReceiving.Sorted = False: lstReceiving.ListItems.Clear
    Set rsrrHEADER2 = New ADODB.Recordset
    Set rsrrHEADER2 = gconDMIS.Execute("select RRDate,RRNumber,INV_No,INV_Date,PO_No,PO_Date,Supplier_Code,ReceivedBy_Code,Total_Amount,RRNumber from OSMS_RRHEADER  where RRDate like'" & xxx & "%' order by RRDate asc")
    If Not (rsrrHEADER2.EOF And rsrrHEADER2.BOF) Then
        Listview_Loadval Me.lstReceiving.ListItems, rsrrHEADER2
        lstReceiving.Refresh
        lstReceiving.Enabled = False
    End If
    
End Sub

Sub FillGrid1()
    Dim rsrrHEADER2 As ADODB.Recordset
    lstReceiving.Enabled = False
    lstReceiving.Sorted = False: lstReceiving.ListItems.Clear
    Set rsrrHEADER2 = New ADODB.Recordset
    Set rsrrHEADER2 = gconDMIS.Execute("select RRNumber,RRDate,INV_No,INV_Date,PO_No,PO_Date,Supplier_Code,ReceivedBy_Code,Total_Amount,RRNumber from OSMS_RRHEADER  order by RRNumber asc")
    If Not (rsrrHEADER2.EOF And rsrrHEADER2.BOF) Then
        Listview_Loadval Me.lstReceiving.ListItems, rsrrHEADER2
        lstReceiving.Refresh
        lstReceiving.Enabled = True
    End If
End Sub

Sub FillSearchGrid1(xxx As String)
    Dim rsrrHEADER2 As ADODB.Recordset
    lstReceiving.Enabled = False
    lstReceiving.Sorted = False: lstReceiving.ListItems.Clear
    Set rsrrHEADER2 = New ADODB.Recordset
    Set rsrrHEADER2 = gconDMIS.Execute("select RRNumber,RRDate,INV_No,INV_Date,PO_No,PO_Date,Supplier_Code,ReceivedBy_Code,Total_Amount,RRNumber from OSMS_RRHEADER  where RRNumber like'" & xxx & "%' order by RRNumber asc")
    If Not (rsrrHEADER2.EOF And rsrrHEADER2.BOF) Then
        Listview_Loadval Me.lstReceiving.ListItems, rsrrHEADER2
        lstReceiving.Refresh
        lstReceiving.Enabled = True
    End If
End Sub

Private Sub optNum_Click()
    If txtSearch = "" Then FillGrid1 Else FillSearchGrid1 (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
    lstReceiving.ColumnHeaders(1).Text = "MRR NUMBER"
    lstReceiving.ColumnHeaders(1).Width = 1450
    lstReceiving.ColumnHeaders(2).Text = "MRR DATE"
    lstReceiving.ColumnHeaders(2).Width = 1250
End Sub

Private Sub optDate_Click()
    If txtSearch = "" Then FillGrid2 Else FillSearchGrid2 (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
    lstReceiving.ColumnHeaders(1).Text = "MRR DATE"
    lstReceiving.ColumnHeaders(1).Width = 1250
    lstReceiving.ColumnHeaders(2).Text = "MRR NUMBER"
    lstReceiving.ColumnHeaders(2).Width = 1450
End Sub






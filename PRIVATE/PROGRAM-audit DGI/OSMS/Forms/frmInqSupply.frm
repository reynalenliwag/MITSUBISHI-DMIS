VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOSMSInquirySupply 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SupplyInventory"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11910
   Begin VB.Frame Trans_No 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5715
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   11745
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
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   510
         Width           =   11595
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "Supply &Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3270
         TabIndex        =   3
         Top             =   150
         Width           =   1845
      End
      Begin VB.OptionButton optCode 
         Caption         =   "Supply &Code"
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
         Left            =   1110
         TabIndex        =   2
         Top             =   180
         Value           =   -1  'True
         Width           =   2175
      End
      Begin MSComctlLib.ListView lstInventory 
         Height          =   4725
         Left            =   60
         TabIndex        =   1
         Top             =   930
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8334
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
         MouseIcon       =   "frmInqSupply.frx":0000
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SUPPLY CODE"
            Object.Width           =   2823
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SUPPLY DESC"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SUPPLIER CODE"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "LAST RR DATE"
            Object.Width           =   3264
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "LAST ISSUE DATE"
            Object.Width           =   3440
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "SUPPLY ON HAND"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   1065
      End
   End
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
      TabIndex        =   7
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
      Left            =   9660
      TabIndex        =   6
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inquiry Supply Inventory"
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
      Left            =   150
      TabIndex        =   8
      Top             =   60
      Width           =   2835
   End
End
Attribute VB_Name = "frmOSMSInquirySupply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSupply As ADODB.Recordset

Sub rsRefresh()
    Set rsSupply = New ADODB.Recordset
    rsSupply.Open "select * from OSMS_SUPPLY order by Supply_Code asc", gconDMIS
End Sub

Sub StoreMemVars()
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        FillGrid1
    Else
        MsgBoxXP "Record is Empty!", "No Record", XP_OKOnly, msg_Information
    End If
End Sub


Function RecordFound(AAA As Variant) As Boolean
    If AAA <> "" Then
        Dim rsRecordFound As ADODB.Recordset
        Set rsRecordFound = New Recordset
        rsRecordFound.Open "Select Supply_Description from OSMS_SUPPLY order by Supply_Code asc", gconDMIS
        rsRecordFound.Find "Supply_Description like '" & AAA & "%'"
        If Not rsRecordFound.EOF Then
            rsSupply.Bookmark = rsRecordFound.Bookmark
            RecordFound = True
        Else
            Set rsRecordFound = New Recordset
            rsRecordFound.Open "Select * from OSMS_SUPPLY order by Supply_Code asc", gconDMIS
            rsRecordFound.Find "Supply_Code = '" & AAA & "'"
            If Not rsRecordFound.EOF Then
                rsSupply.Bookmark = rsRecordFound.Bookmark
                RecordFound = True
            Else
                RecordFound = False
            End If
        End If
    End If
End Function

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    txtSearch.Text = ""
    FillGrid1
End Sub

'Sub FillGrid()
'Set rsSupply = New ADODB.Recordset
'    rsSupply.Open "select * from OSMS_SUPPLY order by Supply_Code asc", gconDMIS
'If Not rsSupply.EOF And Not rsSupply.BOF Then
'      rsSupply.MoveFirst
'      cleargrid grdSupply
'      Do While Not rsSupply.EOF
'         grdSupply.AddItem Null2String(rsSupply!Supply_Code) & Chr(9) & _
          '                           Null2String(rsSupply!Supply_Description) & Chr(9) & _
          ''                           Null2String(rsSupply!Supplier_code) & Chr(9) & _
          '                           Null2Date(rsSupply!lastrrdate) & Chr(9) & _
          '                           Null2Date(rsSupply!LastIssueDate) & Chr(9) & _
          '                           Null2String(rsSupply!Onhand)
'         rsSupply.MoveNext
'      Loop
'      If grdSupply.Rows > 2 Then grdSupply.RemoveItem 1
'Else
'      cleargrid grdSupply
'End If
'End Sub

'Private Sub mDEDelete_Click()
'    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
'        gconDMIS.Execute "delete from OSMS_SUPPLY where Supply_Code = '" & grdSupply.Text & "'"
'        rsRefresh
'        StoreMemvars
'    End If
'End Sub

Private Sub mDEExit_Click()
    Unload Me
End Sub

Private Sub mDESearch_Click()
On Error Resume Next
    txtSearch.SetFocus
    
End Sub


Private Sub lstInventory_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstInventory
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
    If optCode.Value = True Then
        If Trim(txtSearch.Text) = "" Then FillGrid1 Else FillSearchGrid1 (txtSearch.Text)
    Else
        If Trim(txtSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (txtSearch.Text)
    End If
End Sub

Sub FillGrid2()
    Dim rsSupply2 As ADODB.Recordset
    lstInventory.Sorted = False: lstInventory.ListItems.Clear
    lstInventory.Enabled = False
    Set rsSupply2 = New ADODB.Recordset
    Set rsSupply2 = gconDMIS.Execute("select SUPPLY_DESCRIPTION,SUPPLY_CODE,SUPPLIER_CODE,LASTRRDATE,LASTISSUEDATE,ONHAND,SUPPLY_CODE from OSMS_SUPPLY order by SUPPLY_DESCRIPTION asc")
    If Not (rsSupply2.EOF And rsSupply2.BOF) Then
        Listview_Loadval Me.lstInventory.ListItems, rsSupply2
        lstInventory.Refresh
    End If
    lstInventory.Enabled = True
End Sub

Sub FillSearchGrid2(xxx As String)
    Dim rsSupply2 As ADODB.Recordset
    lstInventory.Sorted = False: lstInventory.ListItems.Clear
    lstInventory.Enabled = False
    Set rsSupply2 = New ADODB.Recordset
    Set rsSupply2 = gconDMIS.Execute("select SUPPLY_DESCRIPTION,SUPPLY_CODE,SUPPLIER_CODE,LASTRRDATE,LASTISSUEDATE,ONHAND,SUPPLY_CODE from OSMS_SUPPLY where SUPPLY_DESCRIPTION like'" & xxx & "%' order by SUPPLY_DESCRIPTION asc")
    If Not (rsSupply2.EOF And rsSupply2.BOF) Then
        Listview_Loadval Me.lstInventory.ListItems, rsSupply2
        lstInventory.Refresh
        lstInventory.Enabled = True
    End If
    
End Sub

Sub FillGrid1()
    Dim rsSupply2 As ADODB.Recordset
    lstInventory.Sorted = False: lstInventory.ListItems.Clear
    lstInventory.Enabled = False
    Set rsSupply2 = New ADODB.Recordset
    Set rsSupply2 = gconDMIS.Execute("select SUPPLY_CODE,SUPPLY_DESCRIPTION,SUPPLIER_CODE,LASTRRDATE,LASTISSUEDATE,ONHAND,SUPPLY_CODE  from OSMS_SUPPLY order by SUPPLY_CODE asc")
    If Not (rsSupply2.EOF And rsSupply2.BOF) Then
        Listview_Loadval Me.lstInventory.ListItems, rsSupply2
        lstInventory.Refresh
        lstInventory.Enabled = True
    End If
    
End Sub

Sub FillSearchGrid1(xxx As String)
    Dim rsSupply2 As ADODB.Recordset
    lstInventory.Sorted = False: lstInventory.ListItems.Clear
    lstInventory.Enabled = False
    Set rsSupply2 = New ADODB.Recordset
    Set rsSupply2 = gconDMIS.Execute("select SUPPLY_CODE,SUPPLY_DESCRIPTION,SUPPLIER_CODE,LASTRRDATE,LASTISSUEDATE,ONHAND,SUPPLY_CODE from OSMS_SUPPLY where SUPPLY_CODE like'" & xxx & "%' order by SUPPLY_CODE asc")
    If Not (rsSupply2.EOF And rsSupply2.BOF) Then
        Listview_Loadval Me.lstInventory.ListItems, rsSupply2
        lstInventory.Refresh
        lstInventory.Enabled = True
    End If
    
End Sub
Private Sub optCode_Click()
    If txtSearch = "" Then FillGrid1 Else FillSearchGrid1 (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
    lstInventory.ColumnHeaders(1).Text = "SUPPLY CODE"
    lstInventory.ColumnHeaders(1).Width = 1600
    lstInventory.ColumnHeaders(2).Text = "SUPPLY DESC"
    lstInventory.ColumnHeaders(2).Width = 3500
End Sub
Private Sub optDesc_Click()
    If txtSearch = "" Then FillGrid2 Else FillSearchGrid2 (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
    lstInventory.ColumnHeaders(1).Text = "SUPPLY DESC"
    lstInventory.ColumnHeaders(1).Width = 3500
    lstInventory.ColumnHeaders(2).Text = "SUPPLY CODE"
    lstInventory.ColumnHeaders(2).Width = 1600
End Sub







VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmSMIS_SearchInvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Vehicles Invoice"
   ClientHeight    =   6030
   ClientLeft      =   2970
   ClientTop       =   3735
   ClientWidth     =   8700
   ClipControls    =   0   'False
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
   ForeColor       =   &H00FFFFFF&
   Icon            =   "SearchInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8700
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1245
      TabIndex        =   1
      Top             =   150
      Width           =   7350
   End
   Begin XtremeSuiteControls.TabControl SearchTab 
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   5625
      Width           =   8475
      _Version        =   655364
      _ExtentX        =   14949
      _ExtentY        =   582
      _StockProps     =   64
      Appearance      =   9
      Color           =   16
      PaintManager.Position=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   8
      Item(0).Caption =   "By &VINO"
      Item(0).ControlCount=   0
      Item(1).Caption =   "By &SONo"
      Item(1).ControlCount=   0
      Item(2).Caption =   "By VDRNo"
      Item(2).ControlCount=   0
      Item(3).Caption =   "By Customer"
      Item(3).ControlCount=   0
      Item(4).Caption =   "By Model"
      Item(4).ControlCount=   0
      Item(5).Caption =   "By CS No"
      Item(5).ControlCount=   0
      Item(6).Caption =   "By Prod No"
      Item(6).ControlCount=   0
      Item(7).Caption =   "By Plate No"
      Item(7).ControlCount=   0
   End
   Begin MSComctlLib.ListView lstSearch 
      Height          =   5055
      Left            =   75
      TabIndex        =   2
      Top             =   555
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "SearchInvoice.frx":000C
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "VOUCHER NO."
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "J. DATE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "VENDOR NAME"
         Object.Width           =   12347
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyword:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   195
      Width           =   1125
   End
End
Attribute VB_Name = "frmSMIS_SearchInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim RsSearch As ADODB.Recordset
'Private m_sSearchBy As String
'Public Property Get SearchBy() As String
'
'    SearchBy = m_sSearchBy
'
'End Property
'
'Public Property Let SearchBy(ByVal sSearchBy As String)
'
'    m_sSearchBy = sSearchBy
'
'End Property
'
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then
'        txtSearch.SetFocus
'    End If
'
'    If Shift = 2 Then
'        Select Case KeyCode
'            Case vbKeyV: SearchTab.SelectedItem = 0: txtSearch_Change
'            Case vbKeyP: SearchTab.SelectedItem = 1: txtSearch_Change
'            Case vbKeyO: SearchTab.SelectedItem = 2: txtSearch_Change
'        End Select
'            m_sSearchBy = SearchTab.SelectedItem
'            SearchTab_SelectedChanged (SearchTab.selected)
'    End If
'End Sub
'Sub getTabName(indx As Integer)
'Select Case indx
'Case 0
'getTabName = "VINNO"
'Case 1
'getTabName = "SNO"
'Case 2
'getTabName = "VDR"
'Case 3
'getTabName = "CUS"
'Case 4
'getTabName = "CUS"
'Case 5
'Case 6
'End Select
'End Sub
'
'Private Sub Form_Load()
'    CenterMe Me, frmMain, 0
'    MsgBox m_sSearchBy
'End Sub
'
'Private Sub lstsearch_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    With lstSearch
'        .Sorted = True
'        If .SortKey = ColumnHeader.Index - 1 Then
'            If .SortOrder = lvwAscending Then
'                .SortOrder = lvwDescending
'            Else
'                .SortOrder = lvwAscending
'            End If
'        Else
'            .SortOrder = lvwAscending
'            .SortKey = ColumnHeader.Index - 1
'        End If
'    End With
'End Sub
'
'Private Sub lstsearch_DblClick()
'    If lstSearch.SelectedItem Is Nothing Then Exit Sub
'
'End Sub
'Private Sub lstsearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    rs.MoveFirst
'    rs.Find ("ScheduleID=" & Item.ListSubItems(2).Text)
'    StoreMemvars
'End Sub
'
'

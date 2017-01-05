VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmCSMSSearchCustomerVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Customer"
   ClientHeight    =   6135
   ClientLeft      =   2835
   ClientTop       =   3390
   ClientWidth     =   10785
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
   ForeColor       =   &H00DEDFDE&
   Icon            =   "SearchCustomerVehicle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl SearchTab 
      Height          =   6195
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   10815
      _Version        =   655364
      _ExtentX        =   19076
      _ExtentY        =   10927
      _StockProps     =   64
      Appearance      =   3
      Color           =   4
      PaintManager.Position=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   130
      ItemCount       =   2
      Item(0).Caption =   "By  &Customer Name"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "By &Plate Number"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "tbPlateNo"
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   5565
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   10755
         _Version        =   655364
         _ExtentX        =   18971
         _ExtentY        =   9816
         _StockProps     =   0
         Begin VB.TextBox txtCustomerName 
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
            Left            =   1350
            TabIndex        =   2
            Top             =   30
            Width           =   9285
         End
         Begin MSComctlLib.ListView ListCustomerName 
            Height          =   5115
            Left            =   0
            TabIndex        =   3
            Top             =   450
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   9022
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
            MouseIcon       =   "SearchCustomerVehicle.frx":000C
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CODE"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "NAME"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "VIN"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "PLATE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "CS#"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "COLOR CODE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "YEAR"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "MAKE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "MODEL"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "ENGINE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "SELLING DEALER"
               Object.Width           =   2540
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
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage tbPlateNo 
         Height          =   5565
         Left            =   -69970
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   10755
         _Version        =   655364
         _ExtentX        =   18971
         _ExtentY        =   9816
         _StockProps     =   0
         Begin VB.TextBox txtPlateNumber 
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
            Left            =   1350
            TabIndex        =   5
            Top             =   60
            Width           =   9375
         End
         Begin MSComctlLib.ListView ListPlateNumber 
            Height          =   5115
            Left            =   0
            TabIndex        =   8
            Top             =   450
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   9022
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
            MouseIcon       =   "SearchCustomerVehicle.frx":0326
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CODE"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "NAME"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "VIN"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "PLATE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "CS#"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "COLOR CODE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "YEAR"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "MAKE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "MODEL"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "ENGINE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "SELLING DEALER"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label4 
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
            Left            =   120
            TabIndex        =   7
            Top             =   60
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "frmCSMSSearchCustomerVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FUNCTION FEATURE   :to be able to search for customer vehicle information for searching infomation
'DATE STARTED       : 10/18/2007
'LAST UPDATED       : 10/18/2007
'WHO UPDATED        : AXP
'UPDATING CODE      : AXP10/18/200712:06
'REQUEST NO         : NONE
Option Explicit
Dim rsCusVeh                                          As New ADODB.Recordset
Dim Y                                                  As Long
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
    
        Select Case SEARCH_TAB
            Case 0:
                If Trim(txtCustomerName) <> "" Then
                    On Error Resume Next
                    txtCustomerName.SetFocus
                Else
                    Unload Me
                End If
            Case 5:
                If Trim(txtPlateNumber) <> "" Then
                    On Error Resume Next
                    txtPlateNumber.SetFocus
                Else
                    Unload Me
                End If
        End Select
    End If
    If Shift = 2 Then
        On Error GoTo ErrorCode:
        Select Case KeyCode
            Case vbKeyC: SearchTab.SelectedItem = 0
            Case vbKeyE: SearchTab.SelectedItem = 1
            Case vbKeyI: SearchTab.SelectedItem = 2
            Case vbKeyV: SearchTab.SelectedItem = 3
            Case vbKeyS: SearchTab.SelectedItem = 4
            Case vbKeyP: SearchTab.SelectedItem = 5
        End Select
        SEARCH_TAB = SearchTab.Selected.Index
        SearchTab.SelectedItem = SEARCH_TAB
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    txtCustomerName_Change
End Sub

Private Sub ListCustomerName_DblClick()
      With frmCSMSAddVehicle
            .CustomerCode = ListCustomerName.SelectedItem.Text
            .labCustCode.Caption = ListCustomerName.SelectedItem.Text
            .labCustomer.Caption = ListCustomerName.SelectedItem.ListSubItems(1).Text
    End With
    Unload Me
    frmCSMSAddVehicle.Show 1
    
End Sub

Private Sub ListCustomerName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtCustomerName.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListCustomerName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ListCustomerName_DblClick
    End If
End Sub

Private Sub ListPlateNumber_DblClick()
      With frmCSMSAddVehicle
      
            .CustomerCode = ListPlateNumber.SelectedItem.Text
            .labCustCode.Caption = ListPlateNumber.SelectedItem.Text
            .labCustomer.Caption = ListPlateNumber.SelectedItem.ListSubItems(1).Text
    End With
    Unload Me
    frmCSMSAddVehicle.Show 1
End Sub

Private Sub ListPlateNumber_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtPlateNumber.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListPlateNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ListPlateNumber_DblClick
    End If
End Sub






Private Sub SearchTab_Click(PreviousTab As Integer)
    SEARCH_TAB = SearchTab.Tab
    DoEvents
    Select Case SEARCH_TAB
        Case 0
            txtCustomerName.Enabled = True: ListCustomerName.Enabled = True
            Me.Caption = "Search Item by Customer Name"
            On Error Resume Next
            txtCustomerName.SetFocus
        Case 1
        Case 2
        Case 3
            txtPlateNumber.Enabled = True: ListPlateNumber.Enabled = True
            Me.Caption = "Search Item by Plate Number Order"
            On Error Resume Next
            txtPlateNumber.SetFocus
        Case 4
        Case 5
    End Select
End Sub



Private Sub txtCustomerName_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCustomerName.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListCustomerName.Enabled = True And ListCustomerName.ListItems.Count > 0 Then
            ListCustomerName.SetFocus
        End If
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtCustomerName_Change()
      If txtCustomerName = "" Then
        ListCustomerName.Enabled = False
        Me.ListCustomerName.Sorted = False: Me.ListCustomerName.ListItems.Clear
        Set rsCusVeh = New ADODB.Recordset
        Set rsCusVeh = gconDMIS.Execute("SELECT [CUSCDE] ,[NIYM],[VIN],[PLATE_NO],[VCOND_NO],[CLRCDE],[YER],[MAKE],[MODEL],[ENGINE],[SELLING_DEALER],[ID]  From [DMIS].[dbo].[CSMS_CusVeh] order by niym asc")
        If Not (rsCusVeh.EOF And rsCusVeh.BOF) Then
            Listview_Loadval Me.ListCustomerName.ListItems, rsCusVeh
            ListCustomerName.Enabled = True
        End If
    Else
        Me.ListCustomerName.Sorted = False: Me.ListCustomerName.ListItems.Clear
        Set rsCusVeh = New ADODB.Recordset
        Set rsCusVeh = gconDMIS.Execute("select  [CUSCDE] ,[NIYM],[VIN],[PLATE_NO],[VCOND_NO],[CLRCDE],[YER],[MAKE],[MODEL],[ENGINE],[SELLING_DEALER],[ID]  From [DMIS].[dbo].[CSMS_CusVeh] Where niym like '" & Trim(Me.txtCustomerName) & "%' order by niym asc")
        If Not (rsCusVeh.EOF And rsCusVeh.BOF) Then
            Listview_Loadval Me.ListCustomerName.ListItems, rsCusVeh
            ListCustomerName.Enabled = True
        End If
    End If
End Sub

Private Sub txtPlateNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtPlateNumber.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListPlateNumber.Enabled = True And ListPlateNumber.ListItems.Count > 0 Then
            ListPlateNumber.SetFocus
        End If
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtPlateNumber_Change()
    If txtPlateNumber = "" Then
        ListPlateNumber.Enabled = False
        Me.ListPlateNumber.Sorted = False: Me.ListPlateNumber.ListItems.Clear
        Set rsCusVeh = New ADODB.Recordset
        Set rsCusVeh = gconDMIS.Execute("SELECT [CUSCDE] ,[NIYM],[VIN],[PLATE_NO],[VCOND_NO],[CLRCDE],[YER],[MAKE],[MODEL],[ENGINE],[SELLING_DEALER],[ID]  From [DMIS].[dbo].[CSMS_CusVeh] order by niym asc")
        If Not (rsCusVeh.EOF And rsCusVeh.BOF) Then
            Listview_Loadval Me.ListPlateNumber.ListItems, rsCusVeh
            ListPlateNumber.Enabled = True
        End If
    Else
        Me.ListPlateNumber.Sorted = False: Me.ListPlateNumber.ListItems.Clear
        Set rsCusVeh = New ADODB.Recordset
        Set rsCusVeh = gconDMIS.Execute("select  [CUSCDE] ,[NIYM],[VIN],[PLATE_NO],[VCOND_NO],[CLRCDE],[YER],[MAKE],[MODEL],[ENGINE],[SELLING_DEALER],[ID]  From [DMIS].[dbo].[CSMS_CusVeh] Where plate_no like '" & Trim(Me.txtPlateNumber) & "%' order by niym asc")
        If Not (rsCusVeh.EOF And rsCusVeh.BOF) Then
            Listview_Loadval Me.ListPlateNumber.ListItems, rsCusVeh
            ListPlateNumber.Enabled = True
        End If
    End If
End Sub


Sub clearListView()
    For Y = 1 To Me.ListCustomerName.ListItems.Count
        If Me.ListCustomerName.ListItems.Count <= 0 Then Exit For
        Me.ListCustomerName.Sorted = False
        Me.ListCustomerName.ListItems.Remove Me.ListCustomerName.SelectedItem.Index
    Next Y
  
   
    For Y = 1 To Me.ListPlateNumber.ListItems.Count
        If Me.ListPlateNumber.ListItems.Count <= 0 Then Exit For
        Me.ListPlateNumber.Sorted = False
        Me.ListPlateNumber.ListItems.Remove Me.ListPlateNumber.SelectedItem.Index
    Next Y
   
   
End Sub

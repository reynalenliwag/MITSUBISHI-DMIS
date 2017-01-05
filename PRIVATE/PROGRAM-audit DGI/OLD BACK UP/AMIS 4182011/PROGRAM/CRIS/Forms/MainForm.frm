VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#10.4#0"; "CODEJO~1.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicles Sales Monitoring"
   ClientHeight    =   8100
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11835
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
   ForeColor       =   &H8000000F&
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11835
   Begin VB.CommandButton Command8 
      Caption         =   "Sales &Calendar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   105
      TabIndex        =   15
      Top             =   5130
      Width           =   2610
   End
   Begin XtremeCalendarControl.DatePicker DatePicker1 
      Height          =   2580
      Left            =   90
      TabIndex        =   14
      Top             =   90
      Width           =   2625
      _Version        =   655364
      _ExtentX        =   4630
      _ExtentY        =   4551
      _StockProps     =   64
      AutoSize        =   0   'False
      ShowNoneButton  =   0   'False
      Show3DBorder    =   2
      TextTodayButton =   "Go to today"
   End
   Begin VB.CommandButton Command7 
      Caption         =   "AOR Computation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   105
      TabIndex        =   13
      Top             =   6225
      Width           =   2610
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Inquiry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   105
      TabIndex        =   7
      Top             =   5670
      Width           =   2610
   End
   Begin VB.CommandButton cmdSearchCustomer 
      Caption         =   "&Search Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   105
      TabIndex        =   6
      Top             =   4575
      Width           =   2610
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Applications"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   105
      TabIndex        =   5
      Top             =   3960
      Width           =   2610
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sales Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   105
      TabIndex        =   4
      Top             =   3420
      Width           =   2610
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Prospecting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   105
      TabIndex        =   3
      Top             =   2865
      Width           =   2610
   End
   Begin MSComctlLib.ListView lstSalesOrder 
      Height          =   2280
      Left            =   2835
      TabIndex        =   0
      Top             =   2625
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   4022
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "MainForm.frx":030A
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CODE"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Address"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tel. No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Unit Model"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "SAE"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView lstIndividual 
      Height          =   2595
      Left            =   2850
      TabIndex        =   1
      Top             =   5175
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "MainForm.frx":046C
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Address"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tel. No."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Unit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date Applied"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Assigned SAE"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView lstCorporate 
      Height          =   1695
      Left            =   2850
      TabIndex        =   2
      Top             =   6090
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "MainForm.frx":05CE
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Business Name"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Office Address"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tel. No."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Unit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date Applied"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Assigned SAE"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView lstProspect 
      Height          =   2055
      Left            =   2835
      TabIndex        =   8
      Top             =   285
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
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
      MouseIcon       =   "MainForm.frx":0730
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CODE"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name "
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Subject"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Model"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "SA"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Car Loan Application for Corporate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2865
      TabIndex        =   12
      Top             =   5835
      Width           =   4185
   End
   Begin VB.Label Label1 
      Caption         =   "Car Loan Application for Individual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2865
      TabIndex        =   11
      Top             =   4905
      Width           =   4665
   End
   Begin VB.Label Label1 
      Caption         =   "Sales Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2850
      TabIndex        =   10
      Top             =   2340
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Prospective Clients"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2865
      TabIndex        =   9
      Top             =   15
      Width           =   2055
   End
   Begin VB.Menu mnuSalesOrder 
      Caption         =   "Sales Order"
      Visible         =   0   'False
      Begin VB.Menu mnuSalesOrder_Edit 
         Caption         =   "&Edit this Customer"
      End
      Begin VB.Menu mnuSalesOrder_Delete 
         Caption         =   "&Delete this Customer"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents SearchMaster As frmCRIS_SearhMaster
Attribute SearchMaster.VB_VarHelpID = -1
Private Sub cmdSearchCustomer_Click()
'    frmCustomerSearch.Show
frmCRIS_SearhMaster.Show
End Sub

Private Sub Command1_Click()
    
    frmCRIS_EntryMain.Show
    
End Sub

Private Sub Command2_Click()
    Set SearchMaster = New frmCRIS_SearhMaster
        SearchMaster.Show
    'Set SearchMaster = Nothing
    
End Sub

Private Sub Command3_Click()
    frmIndivAplForm.Show 1
End Sub

Private Sub Command4_Click()
    'MsgBox "Corporate Loan Not Available in Demo Mode...", vbInformation, "Info"
    
End Sub

Private Sub Command5_Click()
    ''''
End Sub

Private Sub Command6_Click()
frmCRIS_Inquiry.Show
'    frmPropectViewandReport.Show
End Sub

Private Sub Command7_Click()
    frmCRIS_AOR.Show 1
End Sub

Private Sub DatePicker1_SelectionChanged()
    If DatePicker1.Selection.BlocksCount = 1 Then
    
       
          
        FillLoanApplicationIndividual "WHERE DateApplied Between '" & DatePicker1.Selection(0).DateBegin & "'  and '" & DatePicker1.Selection(0).DateEnd & "'"
        FillProspect "WHERE DEYT Between '" & DatePicker1.Selection(0).DateBegin & "'  and '" & DatePicker1.Selection(0).DateEnd & "'"
        FillSalesOrder Empty
        
    Else
        FillLoanApplicationIndividual Empty
        FillProspect Empty
        FillSalesOrder Empty
    End If
    
    

End Sub


Private Sub Form_Load()
    CenterMe frmMain, Me, 0
'    monView.Value = Now
    FillAllGrid
    
End Sub
Sub FillAllGrid()
    FillProspect Empty
    FillSalesOrder Empty
    FillLoanApplicationIndividual Empty
End Sub
Sub FillProspect(xDate As String)
    Dim rsUploadData     As ADODB.Recordset
    lstProspect.Sorted = False: lstProspect.ListItems.Clear
    Set rsUploadData = New ADODB.Recordset
    Set rsUploadData = gconDMIS.Execute("SELECT  ProspectID, " & _
                                     "AcctName , Subject ," & _
                                     "VehicleModel, SAE , LogQuote, LogEmail, LogAppointment, LogTestDrive,  " & _
                                     "LogCall, LogJournal, LogLetter, ProfileType , ProfileID, U_S FROM CRIS_Prospects Where D_S is NULL and LOGCLOSINGDATE IS NULL ")
    If Not rsUploadData.EOF And Not rsUploadData.BOF Then
        Listview_Loadval Me.lstProspect.ListItems, rsUploadData
    End If
End Sub
Sub FillSalesOrder(xDate As String)
    Dim rsUploadData     As ADODB.Recordset
    lstSalesOrder.Sorted = False: lstSalesOrder.ListItems.Clear
    Set rsUploadData = New ADODB.Recordset
        If xDate = vbNullString Then
            Set rsUploadData = gconDMIS.Execute(" Select code , CustName,HomeAddress,HomeTelNo,Model,SalesAE from SMIS_SALESORDER  order by CustName asc")
        Else
            Set rsUploadData = gconDMIS.Execute(" Select code , CustName,HomeAddress,HomeTelNo,Model,SalesAE from SMIS_SALESORDER " & xDate & " order by CustName asc")
        End If
    If Not rsUploadData.EOF And Not rsUploadData.BOF Then
        Listview_Loadval Me.lstSalesOrder.ListItems, rsUploadData
    End If

End Sub
Sub FillLoanApplicationIndividual(xDate As String)
   Dim rsUploadData     As ADODB.Recordset
    lstIndividual.Sorted = False: lstIndividual.ListItems.Clear
    Set rsUploadData = New ADODB.Recordset
    If xDate = vbNullString Then
        Set rsUploadData = gconDMIS.Execute("Select isnull(Ind_Apl_LastName,'')  + ' . ' + isnull(Ind_Apl_FirstName, '')  ,Ind_Address,Ind_TelNo,Ind_LoanApl_UnitModel,DateApplied,Ind_LoanApl_SAE,Status  from SMIS_LoanIndiv ")
    Else
        Set rsUploadData = gconDMIS.Execute("Select Ind_Apl_LastName,Ind_Apl_FirstName,Ind_Address,Ind_TelNo,Ind_LoanApl_UnitModel,DateApplied,Ind_LoanApl_SAE,Status  from SMIS_LoanIndiv " & xDate)
    End If

    
    If Not rsUploadData.EOF And Not rsUploadData.BOF Then
       Listview_Loadval Me.lstIndividual.ListItems, rsUploadData
    End If

End Sub
Sub FillLoanApplicationCorp(xDate As String)
'Corporate Appl
    'Dim rsUploadData     As ADODB.Recordset
    'lstCorporate.Sorted = False: lstCorporate.ListItems.Clear
    'Set rsUploadData = New ADODB.Recordset
    'Set rsUploadData = gconDMIS.Execute("Select Apl_Corp_Busname,Apl_Corp_OfficeAdd,Apl_Corp_TelNo,LAF_UnitModel,DateApplied,LAF_SAE,Status  from SMIS_LoanCorp")
    'If Not rsUploadData.EOF And Not rsUploadData.BOF Then
     '   Listview_Loadval Me.lstCorporate.ListItems, rsUploadData
    'End If

End Sub
Private Sub lstSalesOrder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        '     PopupMenu mnuSalesOrder
    End If
End Sub
Private Sub mnuSalesOrder_Edit_Click()
    frmSalesOrder.txtSaveMe = lstSalesOrder.SelectedItem.SubItems(5)
    frmSalesOrder.Show
End Sub

Private Sub Timer1_Timer()
    labClock.caption = Time
End Sub

Private Sub SearchMaster_SelectionMade(oCusRs As ADODB.Recordset)
    If Not oCusRs Is Nothing Then
        frmcris_SalesOrder.AddNewSO 0, oCusRs!CUSTYPE, 0, oCusRs!ID
        frmcris_SalesOrder.Show
    End If
End Sub
